import * as React from "react";
import styles from "./WpStarRating.module.scss";
import type { IWpStarRatingProps } from "./IWpStarRatingProps";
import { escape } from "@microsoft/sp-lodash-subset";
import StarRatings from "react-star-ratings";

import { getSP } from "../pnpjsConfig";
import { SPFI } from "@pnp/sp";

interface IWpStarRatingState {
  rating: number;
  isFeedbackVisible: boolean;
  isSuccessLabelVisible: boolean;
}

export default class WpStarRating extends React.Component<
  IWpStarRatingProps,
  IWpStarRatingState
> {
  public _sp: SPFI;

  constructor(props: IWpStarRatingProps) {
    super(props);

    this._sp = getSP();
    this.state = {
      rating: 0, // Default rating value
      isFeedbackVisible: false, // Add a state to control visibility
      isSuccessLabelVisible: false,
    };

    this.changeRating = this.changeRating.bind(this);
    this.hideFeedbackDiv = this.hideFeedbackDiv.bind(this);
  }

  private changeRating(newRating: number) {
    this.setState({
      rating: newRating,
    });
    this.showFeedbackDiv();
  }

  private showFeedbackDiv() {
    this.setState({ isFeedbackVisible: true }); // Method to show the div
  }

  private hideFeedbackDiv() {
    (document.getElementById("texPropertyValue") as HTMLTextAreaElement).value =
      "";
    this.setState({
      rating: 0,
      isFeedbackVisible: false,
      isSuccessLabelVisible: false,
    });
  }

  public render(): React.ReactElement<IWpStarRatingProps> {
    const {
      description,
      userDisplayName,
      numberOfStars,
      starRatedColor,
      starHoverColor,
      starEmptyColor,
      starDimension,
      starSpacing,
    } = this.props;

    return (
      <section className={`${styles.wpStarRating}`}>
        <div className={styles.welcome}>
          <h2>Hello , {escape(userDisplayName)} !</h2>
          <div>
            <strong>{escape(description)}</strong>
          </div>
        </div>
        <div>
          <div id="starsDiv" className={`${styles.starsDiv}`}>
            <StarRatings
              rating={this.state.rating}
              numberOfStars={numberOfStars}
              starRatedColor={starRatedColor}
              starHoverColor={starHoverColor}
              starEmptyColor={starEmptyColor}
              starDimension={starDimension}
              starSpacing={starSpacing}
              changeRating={this.changeRating}
            />
          </div>

          {this.state.isFeedbackVisible && (
            <div
              id="optionalFeedbackDiv"
              className={`${styles.optionalFeedbackDiv}`}
            >
              <h3>Please add your optional comment here</h3>
              <textarea id="texPropertyValue" rows={5}></textarea>
              <div id="buttonDiv">
                <button
                  className={styles.submitButton}
                  onClick={() =>
                    this.createNewItemUserFeedback(
                      this.props.selectedListUserFeedback,
                      this.props.selectedListFeedbackDetail,
                      this.state.rating
                    )
                  }
                >
                  Submit
                </button>
                <button onClick={this.hideFeedbackDiv}>Cancel</button>
              </div>
              {this.state.isSuccessLabelVisible && (
                <label id="successLabel">
                  <h3>Your feedback successfully submitted ! Thank you ✌️</h3>
                </label>
              )}
            </div>
          )}
        </div>
      </section>
    );
  }

  // method to use pnp objects and create new item
  private async createNewItemUserFeedback(
    itemTitle: string,
    feedbackPageDetailListTitle: string,
    ratingValue: number
  ): Promise<void> {
    try {
      const currentPageUrl = this.props.currentPageUrl;

      const parsedUserFeedbackList = itemTitle
        ? JSON.parse(itemTitle)
        : { id: "", title: "" };
      const userFeedbackLisTitle = parsedUserFeedbackList.title;

      // Adding an item to the list
      const list = this._sp.web.lists.getByTitle(userFeedbackLisTitle);
      const result = await list.items.add({
        Title: currentPageUrl,
        Rating: ratingValue,
        Comment: (
          document.getElementById("texPropertyValue") as HTMLTextAreaElement
        ).value,
      });
      this.CreateNewItemFeedbackDetails(
        feedbackPageDetailListTitle,
        currentPageUrl,
        ratingValue
      );
      console.log("Item added successfully:", result);
      this.setState({ isSuccessLabelVisible: true });
    } catch (error) {
      console.error("Error in createNewItem:", error);
    }
  }

  private async CreateNewItemFeedbackDetails(
    feedbackListTitle: string,
    pageURL: string,
    ratingCount: number
  ): Promise<void> {
    try {
      //const listId = parsedList.id;

      const parsedListFeedbackDetail = feedbackListTitle
        ? JSON.parse(feedbackListTitle)
        : { id: "", title: "" };
      const feedbackDetailListTitle = parsedListFeedbackDetail.title;
      //const listId = parsedList.id;

      const list = this._sp.web.lists.getByTitle(feedbackDetailListTitle);

      // Query the list to see if an item with the same Title (page URL) already exists
      const items = await list.items.filter(`Title eq '${pageURL}'`)();

      if (items.length > 0) {
        // If the item exists, update it
        const itemId = items[0].Id;
        const currentTotalRating = items[0].TotalRating;
        const currentRatedCount = items[0].RatedCount;

        const updatedTotalRating = currentTotalRating + ratingCount;
        const updatedRatedCount = currentRatedCount + 1;
        const updatedAverageRating = updatedTotalRating / updatedRatedCount;

        // Update the existing item with new values
        await list.items.getById(itemId).update({
          TotalRating: updatedTotalRating,
          RatedCount: updatedRatedCount,
          AverageRating: updatedAverageRating,
        });

        console.log("Item updated successfully:", itemId);
      } else {
        // If the item does not exist, create a new one
        const result = await list.items.add({
          Title: pageURL,
          TotalRating: ratingCount, // The first rating
          RatedCount: 1, // First rating, so the count is 1
          AverageRating: ratingCount, // Since it's the first rating, it's the same as the rating
        });

        console.log("New item created successfully:", result);
      }
    } catch (error) {
      console.error("Error in CreateNewItemFeedbackDetails:", error);
    }
  }
}

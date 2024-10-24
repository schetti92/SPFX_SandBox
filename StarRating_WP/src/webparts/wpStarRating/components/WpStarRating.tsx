import * as React from 'react';
import styles from './WpStarRating.module.scss';
import type { IWpStarRatingProps } from './IWpStarRatingProps';
import { escape } from '@microsoft/sp-lodash-subset';
import StarRatings from 'react-star-ratings';

import { getSP } from '../pnpjsConfig';
import { SPFI } from "@pnp/sp";


export default class WpStarRating extends React.Component<IWpStarRatingProps> {

  public _sp: SPFI;

  constructor(props: IWpStarRatingProps) {
    super(props);
    //this.state = { showmessageBar: false, message: "", itemID: 0 };
    this._sp = getSP();
  }

  public render(): React.ReactElement<IWpStarRatingProps> {
    const {
      selectedList,
      description,
      userDisplayName,
      rating,
      numberOfStars, 
      starRatedColor,
      starHoverColor,
      starEmptyColor,
      starDimension,
      starSpacing,
      changeRating,
    } = this.props;

    const parsedList = selectedList ? JSON.parse(selectedList) : { id: '', title: '' };
    const listTitle = parsedList.title;
    //const listId = parsedList.id; 

    this.state = {
      isVisible: true,
    };


    const handleRatingChange = (newRating: number) => {
      alert(`Current Rating: ${newRating}`);
      changeRating(newRating); 
    };

    
    const resetDiv= () => {
      this.setState({ isVisible: false });
    }

    return (
      <section className={`${styles.wpStarRating}`}>
        <div className={styles.welcome}>
          <h2>Well done, {escape(userDisplayName)}!</h2>
          <h2>Selected List: {escape(listTitle)}</h2>
          <div>Web part property value: <strong>{escape(description)}</strong></div>
        </div>
        <div>
          <div id='starsDiv' className={`${styles.starsDiv}`}>
            <StarRatings
              rating={rating}
              numberOfStars={numberOfStars}
              starRatedColor={starRatedColor}
              starHoverColor={starHoverColor}
              starEmptyColor={starEmptyColor}
              starDimension={starDimension}
              starSpacing={starSpacing}
              changeRating={handleRatingChange}
            />
          </div>
         
          <div id='optionalFeedbackDiv' className={`${styles.optionalFeedbackDiv}`}>
            <h3>Please add your optional comment here</h3>
            <textarea id='texPropertyValue' rows={5} ></textarea>
            <button className={styles.submitButton} onClick={() => this.createNewItem(listTitle)}>Submit</button>
            <button onClick={resetDiv}>Cancel</button>
          </div>
       
        </div>
      </section>
    );
  }

      // method to use pnp objects and create new item
      private async createNewItem(itemTitle: string): Promise<void> {
        try {
          // Adding an item to the list
          const list = this._sp.web.lists.getByTitle(itemTitle);
          const result = await list.items.add({
            Title: (document.getElementById('texPropertyValue') as HTMLTextAreaElement).value
          });
          console.log("Item added successfully:", result);
        } catch (error) {
          console.error("Error in createNewItem:", error);
        }
      }

}

import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneSlider,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'WpStarRatingWebPartStrings';
import WpStarRating from './components/WpStarRating';
import { IWpStarRatingProps } from './components/IWpStarRatingProps';
import { getSP } from './pnpjsConfig';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';


export interface IWpStarRatingWebPartProps {
  selectedListUserFeedback: string;
  selectedListFeedbackDetail: string;
  description: string;
  rating: number;
  numberOfStars: number;
  starRatedColor: string;
  starHoverColor: string;
  starEmptyColor: string;
  starDimension: string;
  starSpacing: string;
}

export default class WpStarRatingWebPart extends BaseClientSideWebPart<IWpStarRatingWebPartProps> {
  private listsDropdownOptions: IPropertyPaneDropdownOption[] = [];

  public render(): void {
    const element: React.ReactElement<IWpStarRatingProps> = React.createElement(
      WpStarRating,
      {
        selectedListUserFeedback: this.properties.selectedListUserFeedback,
        selectedListFeedbackDetail: this.properties.selectedListFeedbackDetail,
        description: this.properties.description,
        userDisplayName: this.context.pageContext.user.displayName,
        rating: 5,
        numberOfStars: this.properties.numberOfStars,
        starRatedColor: this.properties.starRatedColor,
        starHoverColor: this.properties.starHoverColor,
        starEmptyColor: this.properties.starEmptyColor,
        starDimension: this.properties.starDimension,
        starSpacing: this.properties.starSpacing,
        webURL: this.context.pageContext.web.absoluteUrl,
        context: this.context,
        currentPageUrl: this.context.pageContext.web.absoluteUrl,
      }
    );
    ReactDom.render(element, this.domElement);
  }

  public async onInit(): Promise<void> {
    await super.onInit();
    getSP(this.context);
  }

  private fetchLists(): Promise<IPropertyPaneDropdownOption[]> {
    const webURL = this.context.pageContext.web.absoluteUrl;
    const endpoint = `${webURL}/_api/web/lists?$filter=Hidden eq false`;

    return this.context.spHttpClient.get(endpoint, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      })
      .then((data) => {
        const options: IPropertyPaneDropdownOption[] = data.value.map((list: any) => {
          return {
            key: JSON.stringify({ id: list.Id, title: list.Title }),
            text: list.Title
          };
        });
        return options;
      });
  }

  protected onPropertyPaneConfigurationStart(): void {
    if (this.listsDropdownOptions.length === 0) {
      this.fetchLists().then((options) => {
        this.listsDropdownOptions = options;
        this.context.propertyPane.refresh();
      });
    }
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                },),
                PropertyPaneDropdown('selectedListUserFeedback', {
                  label: 'User feedback List',
                  options: this.listsDropdownOptions,
                  disabled: this.listsDropdownOptions.length === 0,
                }),
                PropertyPaneDropdown('selectedListFeedbackDetail', {
                  label: 'Feedback List',
                  options: this.listsDropdownOptions,
                  disabled: this.listsDropdownOptions.length === 0,
                })
              ]
            }
            ,
            {
              //add new group
              groupName: strings.StarRatingPropGroupName,
              groupFields: [
                PropertyPaneSlider('numberOfStars', {
                  label: 'Star Count',
                  min: 1,        // Minimum size of the star in pixels
                  max: 10,       // Maximum size of the star in pixels
                  step: 1,        // Increment by 1 pixel

                }),
                PropertyPaneTextField('starRatedColor', {
                  label: 'Star Rated Color',
                  description: 'Set the color for rated stars (e.g., blue, #0000FF)',
                }),
                PropertyPaneTextField('starHoverColor', {
                  label: 'Star Hover Color',
                  description: 'Set the color for stars on hover (e.g., red, #FF0000)',
                  // value: 'blue'  // Default color for stars on hover
                }),
                PropertyPaneTextField('starEmptyColor', {
                  label: 'Star Empty Color',
                  description: 'Set the color for empty stars (e.g., gray, #808080)',
                  // value: 'red'  // Default color for empty stars
                }),
                PropertyPaneSlider('starDimension', {
                  label: 'Star Dimension',
                  min: 30,        // Minimum size of the star in pixels
                  max: 120,       // Maximum size of the star in pixels
                  step: 10,        // Increment by 1 pixel

                }),
                PropertyPaneSlider('starSpacing', {
                  label: 'Star Spacing',
                  min: 0,        // Minimum size of the star in pixels
                  max: 30,       // Maximum size of the star in pixels
                  step: 5,        // Increment by 1 pixel
                }),
              ]
            }
          ]
        }
      ]
    };
  }
}

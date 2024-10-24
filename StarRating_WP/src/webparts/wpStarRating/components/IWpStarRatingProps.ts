import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IWpStarRatingProps {
  selectedList: string;
  description: string;
  userDisplayName: string;
  rating: number;
  numberOfStars: number;
  starRatedColor: string;
  starHoverColor: string;
  starEmptyColor: string;
  starDimension: string;
  starSpacing: string;
  changeRating: (newRating: number) => void;
  webURL: string;
  context: WebPartContext;
  
  
}

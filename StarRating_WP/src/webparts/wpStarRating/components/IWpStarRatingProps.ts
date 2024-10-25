import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IWpStarRatingProps {
  selectedListUserFeedback: string;
  selectedListFeedbackDetail: string;
  description: string;
  userDisplayName: string;
  rating: number;
  numberOfStars: number;
  starRatedColor: string;
  starHoverColor: string;
  starEmptyColor: string;
  starDimension: string;
  starSpacing: string;
  webURL: string;
  context: WebPartContext;
  currentPageUrl: string;
}

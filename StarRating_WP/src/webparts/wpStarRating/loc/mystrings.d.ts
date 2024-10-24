declare interface IWpStarRatingWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  StarRatingPropGroupName: string;
  DescriptionFieldLabel: string;
  RatingFieldLabel: "Rating"
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppLocalEnvironmentOffice: string;
  AppLocalEnvironmentOutlook: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
  AppOfficeEnvironment: string;
  AppOutlookEnvironment: string;
  UnknownEnvironment: string;
}

declare module 'WpStarRatingWebPartStrings' {
  const strings: IWpStarRatingWebPartStrings;
  export = strings;
}

declare interface IMyTeamsAdaptiveCardExtensionStrings {
  PropertyPaneDescription: string;
  TitleFieldLabel: string;
  Title: string;
  SubTitle: string;
  PrimaryText: string;
  Description: string;
  QuickViewButton: string;
}

declare module 'MyTeamsAdaptiveCardExtensionStrings' {
  const strings: IMyTeamsAdaptiveCardExtensionStrings;
  export = strings;
}

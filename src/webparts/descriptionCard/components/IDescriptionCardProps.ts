import { SPHttpClient } from "@microsoft/sp-http";

export interface IDescriptionCardProps {
  listName: string;
  siteUrl: string;
  spHttpClient: SPHttpClient;
  showTitle: boolean;
  iconSize: string;
  itemHeight: number;
  itemBgColor: string;
  titleFontSize: number;
  titleColor: string;
  descFontSize: number;
  descColor: string;

  enableSlider: boolean;
  desktopItems: number;
  tabletItems: number;
  mobileItems: number;
}

export interface ICardItem {
  id: number;
  title: string;
  description: string;
  link: string;
  iconUrl: string;
  linkText: string;
}

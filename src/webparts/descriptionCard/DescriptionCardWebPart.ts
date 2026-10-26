import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  IPropertyPaneConfiguration,
  IPropertyPaneField,
  PropertyPaneTextField,
  PropertyPaneToggle,
  PropertyPaneSlider,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";

import DescriptionCard from "./components/DescriptionCard";
import { IDescriptionCardProps } from "./components/IDescriptionCardProps";

export interface IDescriptionCardWebPartProps {
  listName: string;
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

// Single place for default values (used by render() and the property pane)
const DEFAULTS = {
  iconSize: '42px',
  itemHeight: 150,
  itemBgColor: "#232b36",
  titleFontSize: 15,
  titleColor: "#ffffff",
  descFontSize: 13,
  descColor: "#e6e6e6",
  enableSlider: false,
  desktopItems: 5,
  tabletItems: 3,
  mobileItems: 1,
};

export default class DescriptionCardWebPart extends BaseClientSideWebPart<IDescriptionCardWebPartProps> {
  public render(): void {
    const p = this.properties;

    const element: React.ReactElement<IDescriptionCardProps> =
      React.createElement(DescriptionCard, {
        listName: p.listName,
        siteUrl: this.context.pageContext.web.absoluteUrl,
        spHttpClient: this.context.spHttpClient,
        showTitle: p.showTitle !== false,
        iconSize: p.iconSize || DEFAULTS.iconSize,
        itemHeight: p.itemHeight || DEFAULTS.itemHeight,
        itemBgColor: p.itemBgColor || DEFAULTS.itemBgColor,
        titleFontSize: p.titleFontSize || DEFAULTS.titleFontSize,
        titleColor: p.titleColor || DEFAULTS.titleColor,
        descFontSize: p.descFontSize || DEFAULTS.descFontSize,
        descColor: p.descColor || DEFAULTS.descColor,
        enableSlider: p.enableSlider === true,
        desktopItems: p.desktopItems || DEFAULTS.desktopItems,
        tabletItems: p.tabletItems || DEFAULTS.tabletItems,
        mobileItems: p.mobileItems || DEFAULTS.mobileItems,
      });

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  // Re-draw the pane when the slider toggle changes, so the fields show/hide immediately
  protected onPropertyPaneFieldChanged(
    propertyPath: string,
    oldValue: unknown,
    newValue: unknown
  ): void {
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    if (propertyPath === "enableSlider") {
      this.context.propertyPane.refresh();
    }
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const p = this.properties;

    // Slider settings are added only when the slider is enabled
    const sliderFields: IPropertyPaneField<any>[] = [
      PropertyPaneToggle("enableSlider", {
        label: "Show as slider",
        checked: p.enableSlider === true,
      }),
    ];

    if (p.enableSlider === true) {
      sliderFields.push(
        PropertyPaneSlider("desktopItems", {
          label: "Items per view (Desktop)",
          min: 1,
          max: 8,
          step: 1,
          value: p.desktopItems || DEFAULTS.desktopItems,
        }),
        PropertyPaneSlider("tabletItems", {
          label: "Items per view (Tablet)",
          min: 1,
          max: 6,
          step: 1,
          value: p.tabletItems || DEFAULTS.tabletItems,
        }),
        PropertyPaneSlider("mobileItems", {
          label: "Items per view (Mobile)",
          min: 1,
          max: 3,
          step: 1,
          value: p.mobileItems || DEFAULTS.mobileItems,
        })
      );
    }

    return {
      pages: [
        {
          header: {
            description: "Description Card Web Part Settings",
          },
          groups: [
            {
              groupName: "List Settings",
              groupFields: [
                PropertyPaneTextField("listName", {
                  label: "SharePoint List Name",
                  description:
                    "Columns: Title, Description, ItemLink,OrderNumber, Attachments(for Image - not a column)",
                  value: p.listName,
                }),
              ],
            },
            {
              groupName: "Style Settings",
              groupFields: [
                PropertyPaneToggle("showTitle", {
                  label: "Show Title",
                  checked: p.showTitle !== false,
                }),
                PropertyPaneTextField("iconSize", { 
                  label: "Size of Icon",
                  description: "e.g. 50px or 50%",
                  value: p.iconSize || DEFAULTS.iconSize,
                }),
                PropertyPaneSlider("itemHeight", {
                  label: "Item height (px)",
                  min: 100,
                  max: 300,
                  step: 1,
                  value: p.itemHeight || DEFAULTS.itemHeight,
                }),
                PropertyPaneTextField("itemBgColor", {
                  label: "Item background color",
                  description: "e.g. #232b36 or rgb(35,43,54)",
                  value: p.itemBgColor,
                }),
                PropertyPaneSlider("titleFontSize", {
                  label: "Title font size (px)",
                  min: 10,
                  max: 40,
                  step: 1,
                  value: p.titleFontSize || DEFAULTS.titleFontSize,
                }),
                PropertyPaneTextField("titleColor", {
                  label: "Title color",
                  value: p.titleColor,
                }),
                PropertyPaneSlider("descFontSize", {
                  label: "Description font size (px)",
                  min: 10,
                  max: 30,
                  step: 1,
                  value: p.descFontSize || DEFAULTS.descFontSize,
                }),
                PropertyPaneTextField("descColor", {
                  label: "Description color",
                  value: p.descColor,
                }),
              ],
            },
            {
              groupName: "Layout Settings",
              groupFields: sliderFields,
            },
          ],
        },
      ],
    };
  }
}
import { Log } from "@microsoft/sp-core-library";
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName,
} from "@microsoft/sp-application-base";
import * as React from "react";
import * as ReactDom from "react-dom";
import { spfi, SPFx } from "@pnp/sp";

import * as strings from "NewsTrackerApplicationCustomizerStrings";
import SpService from "./service/SpService";
import Constants from "./helpers/Constants";
import INewsTickerProps from "./Component/INewsTickerProps";
import NewsTicker from "./Component/NewsTicker";

const LOG_SOURCE: string = "NewsTickerApplicationCustomizer";

export interface INewsTickerApplicationCustomizerProperties {
  listTitle: string;
  listViewTitle: string;
  bgColor: string;
  textColor: string;
}

export default class NewsTickerApplicationCustomizer extends BaseApplicationCustomizer<INewsTickerApplicationCustomizerProperties> {
  private _topPlaceholder: PlaceholderContent | undefined;
  private _spService: SpService;

  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);
    const sp = spfi().using(SPFx(this.context));
    this._spService = new SpService(sp);

    this.context.placeholderProvider.changedEvent.add(
      this,
      this._renderPlaceHolders
    );

    return Promise.resolve();
  }

  private async _renderPlaceHolders(): Promise<void> {
    // Handling the top placeholder
    if (!this._topPlaceholder) {
      this._topPlaceholder = this.context.placeholderProvider.tryCreateContent(
        PlaceholderName.Top,
        { onDispose: this._onDispose }
      );

      // The extension should not assume that the expected placeholder is available.
      if (!this._topPlaceholder) {
        console.error("The expected placeholder (Top) was not found.");
        return;
      }

      if (
        !this.properties ||
        !this.properties.listTitle ||
        !this.properties.listViewTitle
      ) {
        console.error(
          "listTitle or listViewTitle properties value was not found or empty"
        );
        return;
      }

      if (this._topPlaceholder.domElement) {
        // Get news items
        const newsItems = await this._spService.getNewsItems(
          this.properties.listTitle,
          this.properties.listViewTitle
        );

        // Doesn't need to show news ticker if there is no news for now
        if (!newsItems || newsItems.length === 0) return;

        // Find existing element
        const existingElement = document.getElementById(Constants.ROOT_ID);

        // Stop if another news ticker found
        if (document.body.contains(existingElement)) return;

        const element = React.createElement(NewsTicker, <INewsTickerProps>{
          items: newsItems,
          bgColor: this.properties.bgColor,
          textColor: this.properties.textColor,
          spService: this._spService,
        });
        ReactDom.render(element, this._topPlaceholder.domElement);
      }
    }
  }

  private _onDispose(): void {
    console.log(
      "[NewsTickerApplicationCustomizer._onDispose] Disposed custom top placeholders."
    );

    if (this._topPlaceholder && this._topPlaceholder.domElement) {
      ReactDom.unmountComponentAtNode(this._topPlaceholder.domElement);
    }
  }
}

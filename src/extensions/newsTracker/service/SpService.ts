import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/views";

import { SPFI } from "@pnp/sp";
import { News } from "../models/News";

interface INewsItem {
  Title: string;
  Content: string;
  PublishDate: string;
}

export default class SpService {
  private _sp: SPFI;

  constructor(sp: SPFI) {
    this._sp = sp;
  }

  public async getNewsItems(
    listTitle: string,
    listViewTitle: string
  ): Promise<News[]> {
    // Get xml schema for the "Published News" view
    const list = this._sp.web.lists.getByTitle(listTitle);
    const view = await list.views.getByTitle(listViewTitle)();
    if (!view) return [];

    const items = await list.getItemsByCAMLQuery({ ViewXml: view.ListViewXml });

    return items.map((item: INewsItem) =>
      <News>{
        title: item.Title,
        content: item.Content,
        publishDate: new Date(item.PublishDate),
      }
    );
  }
}

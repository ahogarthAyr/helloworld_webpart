import { ISPListItem } from './HelloWorldWebPart';

export default class MockHttpClient  {

  private static _items: ISPListItem[] = [{ Title: 'Mock List', Id: '1', EncodedAbsUrl: "www", Description: ""  },
                                      { Title: 'Mock List 2', Id: '2', EncodedAbsUrl: "www", Description: ""  },
                                      { Title: 'Mock List 3', Id: '3', EncodedAbsUrl: "www", Description: "" }];

  public static get(): Promise<ISPListItem[]> {
    return new Promise<ISPListItem[]>((resolve) => {
      resolve(MockHttpClient._items);
    });
  } 
}
//Mock data in offline mode

import { ISPList } from './JayWebPart';

export default class MockHttpClient  {

    private static _items: ISPList[] = [{ Title: 'Mock List', Id: '1' ,Created: 'Lib', EntityTypeName: 'http://www.google.com', DecodedUrl: ''},
                                        { Title: 'Mock List 2', Id: '2', Created: 'List', EntityTypeName: 'http://www.google.com', DecodedUrl: ''},
                                        { Title: 'Mock List 3', Id: '3', Created: 'List', EntityTypeName: 'http://www.google.com', DecodedUrl: ''}];

    public static get(): Promise<ISPList[]> {
    return new Promise<ISPList[]>((resolve) => {
            resolve(MockHttpClient._items);
        });
    }
}
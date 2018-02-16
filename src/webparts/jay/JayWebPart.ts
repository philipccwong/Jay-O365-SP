//#region header
import {
  SPHttpClient,
  SPHttpClientResponse   
} from '@microsoft/sp-http';
import {
  Environment,
  EnvironmentType
} from '@microsoft/sp-core-library';

import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import MockHttpClient from './MockHttpClient';
import InternalValue, { InternalKey, InternalKeys } from './SPInternalField';

import styles from './JayWebPart.module.scss';
import * as strings from 'JayWebPartStrings';
import { PropertyPaneCheckbox } from '@microsoft/sp-webpart-base/lib/propertyPane/propertyPaneFields/propertyPaneCheckBox/PropertyPaneCheckbox';
import { PropertyPaneDropdown } from '@microsoft/sp-webpart-base/lib/propertyPane/propertyPaneFields/propertyPaneDropdown/PropertyPaneDropdown';
import { Button } from 'office-ui-fabric-react/lib/Button';

export interface IJayWebPartProps {
  description: string;
  test: string;
  test2: string;
  test3: string;
}

export interface ISPLists {
  value: ISPList[];
}

export interface StorageMetric {

  StorageMetrics: {
  TotalFileCount: string;
  TotalFileStreamSize: string;
  TotalSize: string;
  }

}

export interface Subsites {
  value: Subsite[];
}

export interface Subsite{
  Title: string;
  ServerRelativeUrl: string;
  Url: string;
}


export interface ISPList {
  Title: string;
  Id: string;
  Created: string;
  EntityTypeName: string;
  DecodedUrl: string;

}

export interface SPNumber {

  Title: string;
  Id: string; 
  Created: string;
  EntityTypeName: string;
  DecodedUrl: string;
  NumOfContentType: string;
  NumOfListItem: string;
  NumOfField: string;
  TotalFileCount: string;
  TotalFileStreamSize: string;
  TotalSize: string;

}
//#endregion

//Please refer to Readme.txt
//POC on Site Extraction

export default class JayWebPart extends BaseClientSideWebPart<IJayWebPartProps> {

  public render(): void {

    this.domElement.innerHTML = `
    <div class="${ styles.Jay }">
      <div class="${ styles.container }">
        <div class="${ styles.row }">
          <div class="${ styles.column }">
            <span class="${ styles.title }">Your Site Information </span>
          </div>
        </div>
        <div id="subsiteList" />
        <div/><div id="spListContainer" />
      </div>
    </div>`;

    this._renderListAsync();
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('Description', {
                  label: strings.DescriptionFieldLabel
                }),
                PropertyPaneCheckbox('Show List', {
                  text: "Allow to show List",
                  checked: false,
                  disabled: false
                }),
               PropertyPaneCheckbox('Show Storage', {
                  text: "Allow to show Storage",
                  checked: false,
                  disabled: false
                }),
               PropertyPaneCheckbox('Show Sub-site', {
                  text: "Allow to show Subsite",
                  checked: false,
                  disabled: false
                }),
                PropertyPaneCheckbox('Show Library', {
                  text: "Allow to Library",
                  checked: false,
                  disabled: false
                }),
               PropertyPaneCheckbox('Show stream location from console', {
                  text: "Allow to stream location from console",
                  checked: false,
                  disabled: false
                }),
               PropertyPaneCheckbox('LVT Detection', {
                  text: "Allow to detect LVT",
                  checked: false,
                  disabled: false
                })
              ]
              
            }
          ]
        }
      ]
    };
  }

  //#region SP API Call

  private _getListData(): Promise<ISPLists> {
    return  this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + `/_api/web/lists?$filter=Hidden eq false`, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      });
  }

  private _getContentTypeData(_listId : string): Promise<ISPLists> {
    return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + `/_api/lists(guid'`+ _listId + `')/contenttypes`, SPHttpClient.configurations.v1)
    .then( (response: SPHttpClientResponse) => {
        return response.json();
      }, (e: Error) => {e.message});
    
  }

  private _getItemCountData(_listId : string): Promise<ISPLists> {
    return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + `/_api/lists(guid'`+ _listId + `')/itemcount`, SPHttpClient.configurations.v1)
    .then( (response: SPHttpClientResponse) => {
        return response.json();
      }, (e: Error) => {e.message});
    
  }

  private _getFieldCountData(_listId : string): Promise<ISPLists> {
    return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + `/_api/lists(guid'`+ _listId + `')/fields`, SPHttpClient.configurations.v1)
    .then( (response: SPHttpClientResponse) => {
        return response.json();
      }, (e: Error) => {e.message});
    
  }

  private  _getStorageMatrix(_libraryName : string): Promise<StorageMetric> {
    
    return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + `/_api/web/getFolderByServerRelativeUrl('`+ _libraryName + `')?$select=StorageMetrics&$expand=StorageMetrics`, SPHttpClient.configurations.v1)
    .then( (response: SPHttpClientResponse) => {
        return response.json();
      }, (e: Error) => {e.message});
      
  }

  private  _getSubSite(): Promise<Subsites> {
    
    return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + `/_api/web/webs/?$select=title,ServerRelativeUrl,URL`, SPHttpClient.configurations.v1)
    .then( (response: SPHttpClientResponse) => {
        return response.json();
      }, (e: Error) => {e.message});
      
  }
  //#endregion


  private async _renderListAsync() {
    // Local environment

    let SPNumbers: Array<SPNumber> = new Array();

    if (Environment.type === EnvironmentType.Local) {
      //Offline
      this._getMockListData().then((response) => {
        this._storeSPInfo(response.value, SPNumbers);
        this._renderList(SPNumbers);
      });
    }
    else if (Environment.type == EnvironmentType.SharePoint || 
              Environment.type == EnvironmentType.ClassicSharePoint) {
               
        //Get All Subsite
        
       this._getSubSite().then((response)=> {
          this._storeSubSiteInfo(response.value);
        });

        //Get Current Site Lib/List Info
        await this._getListData()
        .then((response) => {
          this._storeSPInfo(response.value, SPNumbers);
        });

        //Get Content Type Info
        let promises = SPNumbers.map(element => {
           return this._getContentTypeData(element.Id)
            .then((response) => {
              return this._updateContentTypeCount(response.value, element.Id, SPNumbers);
           });
        });

        //Get Number of items
        let promises2 = SPNumbers.map(element => {
           return this._getItemCountData(element.Id)
            .then((response) => {
              return this._updateItemCount(response.value, element.Id, SPNumbers);
           });
        });

        //Get Number of columns
        let promises3 = SPNumbers.map(element => {
           return this._getFieldCountData(element.Id)
            .then((response) => {
              return this._updateFieldCount(response.value, element.Id, SPNumbers);
           });
        });
        
        //Get Storage Info
        let promises4 = SPNumbers.map(element => {
          let x = element.EntityTypeName;
          this._getInternalFieldMapping().then((response) => {
            x = this._DecodeFieldName(response.value, x)
            return this._getStorageMatrix(x)
              .then((response) => {
                  return this._updateFolderSize(response, element.Id, SPNumbers);
          });
        })
       
        });
        
        //Issue here
        Promise.all(promises3)
        .then(results => {
          this._renderList(SPNumbers);
        })
        .catch(e => {
         console.error(e);
        });

    }
  }

  //#region Update Array with API Info

  //Set site info to an array for Output
  private _storeSPInfo(items: ISPList[], SPN: Array<SPNumber>) 
  {
    for (let item of items)
    {
      
      let _item : SPNumber = {Title: item.Title, Id: item.Id,Created: item.Created, EntityTypeName: item.EntityTypeName, DecodedUrl: item.DecodedUrl, 
         NumOfContentType: '', NumOfListItem: '', NumOfField: '', TotalFileCount: 'N/A',TotalFileStreamSize: 'N/A', TotalSize: 'N/A'}
      SPN.push(_item);
        
    }
  }
  
  private _storeSubSiteInfo(items: Subsite[]) 
  {
    let html = `<table class="${styles.list}">`;
    
    for (let item of items)
    {
     
      if (item != null)
      {
      
        html += `<tr><td class="ms-BrandIcon--icon96 ms-BrandIcon--sharepoint"></td><td>
        <ul>
            <li><span class="ms-font-l"><b>Site Name: </b>${item.Title}</span></li>
            <li><span class="ms-font-l"><b>Site URL: </b>${item.ServerRelativeUrl}</span></li>
            <li><span class="ms-font-l"><b>URL: </b><a href='${item.Url}/_layouts/15/workbench.aspx'>${item.Title}</a></span></li>
      </ul></td></tr>`;
      }
      
    }
    html += `</table><div/><div id="spListContainer" />`;
    const listContainer2: Element = this.domElement.querySelector('#subsiteList');
    listContainer2.innerHTML = html;

  }
  

  private _updateContentTypeCount(items: ISPList[], Id: string, SPN: Array<SPNumber>) 
  {
    
      let objIndex:number = 0;
      let i:number = 0;

      
      for (let item of SPN)
      { 
        if (item != null)
        {
          if (item.Id == Id)
          {
            objIndex = i;
            SPN[i].NumOfContentType = items.length.toString();
            return;
          }
            
        }
        i++;
      }
     
  }

  private _updateItemCount(items: ISPList[], Id: string, SPN: Array<SPNumber>) 
  {
    
      let objIndex:number = 0;
      let i:number = 0;

      try{
      for (let item of SPN)
      { 
        if (item != null)
        {
          if (item.Id == Id)
          {
            objIndex = i;
            if (item != null)
              SPN[i].NumOfListItem = items.toString();
            else
              SPN[i].NumOfListItem = '0';
            return;
          }
            
        }
        i++;
      }
    }
    catch(Error)
    {}     
  }

  private _updateFieldCount(items: ISPList[], Id: string, SPN: Array<SPNumber>) 
  {
    
      let objIndex:number = 0;
      let i:number = 0;

      try{

      for (let item of SPN)
      { 
        if (item != null)
        {
          if (item.Id == Id)
          {
            objIndex = i;
            if (item != null)
              SPN[i].NumOfField = items.length.toString();
            else
              SPN[i].NumOfField = '0';
            return;
          }
            
        }
        i++;
      }
      }
      catch(Error)
      {//Do something
      }
  }

  //Storage Matrix Calcation
  private _updateFolderSize(item: StorageMetric, Id: string, SPN: Array<SPNumber>) 
  { 
      let objIndex:number = 0;
      let i:number = 0;

      try{

      if (item != null)
      {
        for (let _item of SPN)
        { 
          if (_item != null)
          {
            if (_item.Id == Id)
            {
              SPN[i].TotalFileCount = item.StorageMetrics.TotalFileCount;
              SPN[i].TotalFileStreamSize = (parseInt(item.StorageMetrics.TotalFileStreamSize)*1.25E-7).toPrecision(4).toString() + "MB";
              SPN[i].TotalSize = (parseInt(item.StorageMetrics.TotalSize)*1.25E-7).toPrecision(4).toString() + "MB";
            }
          }
          i++;
        }
      }     
    }
    catch(Error)
    {
      //Do something
    } 
  }

  //#endregion

  private _renderList(SPN: Array<SPNumber>) { 
  
    let html = `<table>`;
    let i = 0;
    for (let item of SPN)
    {      
      if (item != null)
      {
        if (i%2 == 0)
        {
          html += `<tr>`;
        }
        html += `<td>
        <ul class="${styles.list}">
            <li><div class="ms-BrandIcon--icon96 ms-BrandIcon--dotx"></div><span class="ms-font-s"><b>List Name: </b>${item.Title}</span></li>
            <li><span class="ms-font-s"><b>Content Type: </b>${item.NumOfContentType}</span></li>
            <li><span class="ms-font-s"><b>List Items: </b>${item.NumOfListItem}</span></li>
            <li><span class="ms-font-s"><b>Field Items: </b>${item.NumOfField}</span></li>
            <li><span class="ms-font-s"><b>Time Created: </b>${item.Created}</span></li>
            <li><span class="ms-font-s"><b>Total File Count: </b>${item.TotalFileCount}</span></li>
            <li><span class="ms-font-s"><b>Total File Stream Size: </b>${item.TotalFileStreamSize}</span></li>
            <li><span class="ms-font-s"><b>Total Size: </b>${item.TotalSize}</span></li>
            <li></li><br/>
            </li>
      </ul></td>`;
        if (i%2 == 1)
        {
          html += `</tr>`;
        }
        i++;
      }
      
    }
    html += `</table>`;
    const listContainer: Element = this.domElement.querySelector('#spListContainer');
    listContainer.innerHTML = html;

  }

  //#region Functions

    //Decoding SP Internal Site Name
    private  _DecodeFieldName (items: InternalKey[], input: string): string{
    
      let x = input;
  
      for (let item of items)
      {
        while (x.indexOf(item.value)!=-1)
          x = x.replace(item.value, item.key);
      }
      return x;
    }

    private _getMockListData(): Promise<ISPLists> {
      return MockHttpClient.get()
        .then((data: ISPList[]) => {
          var listData: ISPLists = { value: data };
          return listData;
        }) as Promise<ISPLists>;
    }
  
    private _getInternalFieldMapping(): Promise<InternalKeys> {
      return InternalValue.get()
        .then((data: InternalKey[]) => {
          var listData: InternalKeys = { value: data };
          return listData;
        }) as Promise<InternalKeys>;
    }
  //#endregion
}

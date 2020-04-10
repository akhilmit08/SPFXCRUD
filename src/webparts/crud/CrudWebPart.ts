import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './CrudWebPart.module.scss';
import * as strings from 'CrudWebPartStrings';
import { ISPHttpClientOptions, SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { ISPListItem } from "./ISPListItem";
import * as jQuery from 'jquery';
require('jquery-ui');
import{SPComponentLoader}from'@microsoft/sp-loader';
import { MSGraphClient } from '@microsoft/sp-http';

export interface ICrudWebPartProps {
  description: string;
}

export default class CrudWebPart extends BaseClientSideWebPart<ICrudWebPartProps> {

  protected onInit():Promise<void> {
    SPComponentLoader.loadCss('https://code.jquery.com/ui/1.12.1/themes/base/jquery-ui.min.css');
    return Promise.resolve();
  }
  public render(): void {
    this.context.msGraphClientFactory
      .getClient()
      .then((client: MSGraphClient): void => {        
        client
          .api('/me')
          .get((error, response: any, rawResponse?: any) => {
    this.domElement.innerHTML = `
      <div class="${ styles.crud }">
        <div class="${ styles.container }">
          <div class="${ styles.row }">
            <div class="${ styles.column }">
              <span class="${ styles.title }">Welcome to SharePoint Framework Grap API Demo!</span>
            
              <p class="${ styles.description }"> Display Name: ${response.displayName}</p>
              <p class="${ styles.description }">Email ID: ${response.mail}</p>

            </div>
          </div>
        </div>
      </div>`;
          });
        });
      }

  private _setButtonEventHandlers(): void {
    this.readAllItems();
    this.domElement.querySelector('#btnSubmit').addEventListener('click', () => { this.createListItem(); });
   // this.domElement.querySelector('#btnFetchDetails').addEventListener('click', () => { this.fetchItemByID(); });
    this.domElement.querySelector('#btnUpdate').addEventListener('click', () => { this.updateListItem(); });
    this.domElement.querySelector('#btnDelete').addEventListener('click', () => { this._deleteListItemByID(); });
    jQuery("#btnFetchDetails").click(() => {this.fetchItemByID();});
  }
  private fetchItemByID(): void {
    let id: string= document.getElementById("txtItemID")["value"];
    this._getListItemByID(id).then(listItem => {

    document.getElementById("txtTitleUpdate")["value"] = listItem.Title;
    document.getElementById("ddlVendorUpdate")["value"] = listItem.Vendor;
    document.getElementById("txtProductDescriptionUpdate")["value"] = listItem.ProductDescription;
    document.getElementById("txtCustomerNameUpdate")["value"] = listItem.CustomerName;
    document.getElementById("txtCustomerEmailUpdate")["value"] = listItem.CustomerEmail;
    document.getElementById("txtCustomerPhoneUpdate")["value"] = listItem.CustomerPhone;
    document.getElementById("txtCustomerAddressUpdate")["value"] = listItem.CustomerAddress;
    
    })
    .catch(error => {
      let message: Element = this.domElement.querySelector('#spListCreateItemUpdate');    
      message.innerHTML = "Read: Operation failed. "+error.message;
    });
    } 
    
  private _getListItemByID(id: string): Promise<ISPListItem> {
    const url: string = this.context.pageContext.site.absoluteUrl+"/_api/web/lists/getbytitle('ProductSales')/items?$filter=Id eq "+id;
    return this.context.spHttpClient.get(url, SPHttpClient.configurations.v1)
    .then((response: SPHttpClientResponse) => {
    return response.json();
    })
    .then( (listItems: any) => {
    const untypedItem: any = listItems.value[0];
    const listItem: ISPListItem = untypedItem as ISPListItem;
    return listItem;
    }) as Promise <ISPListItem>;
    }
  private readAllItems(): void {
    this._getListItems().then(listItems => {
      let html: string = '<table border=1 width=100% style="border-collapse: collapse;">';
      html += '<th>Title</th> <th>Vendor</th><th>ProductDescription</th><th>CustomerName</th><th>CustomerEmail</th><th>CustomerPhone</th>';
    listItems.forEach(listItem => {
      html += `<tr>            
      <td>${listItem.Title}</td>
      <td>${listItem.Vendor}</td>
      <td>${listItem.ProductDescription}</td>
      <td>${listItem.CustomerName}</td>
      <td>${listItem.CustomerEmail}</td>
      <td>${listItem.CustomerPhone}</td>        
      </tr>`;
    });
    html += '</table>';
    const listContainer: Element = this.domElement.querySelector('#spListData');
    listContainer.innerHTML = html;
    });

    }

    private _getListItems(): Promise<ISPListItem[]> {
      const url: string = this.context.pageContext.site.absoluteUrl+"/_api/web/lists/getbytitle('ProductSales')/items";
      return this.context.spHttpClient.get(url,SPHttpClient.configurations.v1)
      .then(response => {
      return response.json();
      })
      .then(json => {
      return json.value;
      }) as Promise<ISPListItem[]>;
      }

      private _deleteListItemByID(): void {
        let id: string = document.getElementById("txtItemIDToDelete")["value"];
        const url: string = this.context.pageContext.site.absoluteUrl + "/_api/web/lists/getbytitle('ProductSales')/items(" + id + ")";          
        const headers: any = { "X-HTTP-Method": "DELETE", "IF-MATCH": "*" };
        const spHttpClientOptions: ISPHttpClientOptions = {
          "headers": headers
        };
        this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, spHttpClientOptions)
          .then((response: SPHttpClientResponse) => {
            if (response.status === 204) {
              let message: Element = this.domElement.querySelector('#spListItemDeleteStatus');
              message.innerHTML = "Delete: List Item deleted successfully.";
              this.readAllItems();
            } else {
              let message: Element = this.domElement.querySelector('#spListItemDeleteStatus');
              message.innerHTML = "List Item delete failed." + response.status + " - " + response.statusText;
            }
          });
      }

      private updateListItem(): void {

        var title = document.getElementById("txtTitleUpdate")["value"];
        var vendor = document.getElementById("ddlVendorUpdate")["value"];
        var productDescription = document.getElementById("txtProductDescriptionUpdate")["value"];
        var customerName = document.getElementById("txtCustomerNameUpdate")["value"];
        var customerEmail = document.getElementById("txtCustomerEmailUpdate")["value"];
        var customerPhone = document.getElementById("txtCustomerPhoneUpdate")["value"];
        var customerAddress = document.getElementById("txtCustomerAddressUpdate")["value"];
    
        let id: string = document.getElementById("txtItemID")["value"];
    
        const url: string = this.context.pageContext.site.absoluteUrl + "/_api/web/lists/getbytitle('ProductSales')/items(" + id + ")";
        const itemDefinition: any = {
          "Title": title,
          "Vendor": vendor,
          "ProductDescription": productDescription,
          "CustomerName": customerName,
          "CustomerEmail": customerEmail,
          "CustomerPhone": customerPhone,
          "CustomerAddress": customerAddress
        };
        const headers: any = {
          "X-HTTP-Method": "MERGE",
          "IF-MATCH": "*",
        };
    
        const spHttpClientOptions: ISPHttpClientOptions = {
          "headers": headers,
          "body": JSON.stringify(itemDefinition)
        };
    
        this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, spHttpClientOptions)
          .then((response: SPHttpClientResponse) => {
            if (response.status === 204) {
              let message: Element = this.domElement.querySelector('#spListCreateItemUpdate');
              message.innerHTML = "List Item updated successfully.";
              this.readAllItems();       
            } else {
              let message: Element = this.domElement.querySelector('#spListCreateItemUpdate');
              message.innerHTML = "List Item update failed. " + response.status + " - " + response.statusText;
            }
          });
    
        }

  private createListItem(): void {
    var title = document.getElementById("txtTitle")["value"];
    var vendor = document.getElementById("ddlVendor")["value"];
    var productDescription = document.getElementById("txtProductDescription")["value"];
    var customerName = document.getElementById("txtCustomerName")["value"];
    var customerEmail = document.getElementById("txtCustomerEmail")["value"];
    var customerPhone = document.getElementById("txtCustomerPhone")["value"];
    var customerAddress = document.getElementById("txtCustomerAddress")["value"];

    const url: string = this.context.pageContext.site.absoluteUrl + "/_api/web/lists/getbytitle('ProductSales')/items";

    const itemDefinition: any = {
      "Title": title,
      "Vendor": vendor,
      "ProductDescription": productDescription,
      "CustomerName": customerName,
      "CustomerEmail": customerEmail,
      "CustomerPhone": customerPhone,
      "CustomerAddress": customerAddress
      
    };
    const spHttpClientOptions: ISPHttpClientOptions = {
      "body": JSON.stringify(itemDefinition)
    };
    this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, spHttpClientOptions)
      .then((response: SPHttpClientResponse) => {
        if (response.status === 201) {
          let message: Element = this.domElement.querySelector('#spListCreateItem');
          message.innerHTML = "Create: List Item created successfully.";
          this.clear();
          this.readAllItems();
          //this._operationResults.innerHTML = "Create: List Item created successfully.";
          //this.readAllItemsTabular();
        } else {
          let message: Element = this.domElement.querySelector('#spListCreateItem');
          message.innerHTML = "Create: List Item creation failed. " + response.status + " - " + response.statusText;
        }
      });
  }
  private clear(): void {
    document.getElementById("txtTitle")["value"] = '';
    document.getElementById("ddlVendor")["value"] = 'DELL';
    document.getElementById("txtProductDescription")["value"] = '';
    document.getElementById("txtCustomerName")["value"] = '';
    document.getElementById("txtCustomerEmail")["value"] = '';
    document.getElementById("txtCustomerPhone")["value"] = '';
    document.getElementById("txtCustomerAddress")["value"] = '';
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
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}

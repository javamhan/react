import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version, UrlQueryParameterCollection } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'DocumentViewerWebPartStrings';
import DocumentViewer from './components/DocumentViewer';
import { IDocumentViewerProps } from './components/IDocumentViewerProps';
import CommonService from '../../service/Commonservice';

export interface IDocumentViewerWebPartProps {
  description: string;
  documentUrl: string;
  documents: any[];
  documentsLength: number;
  siteUrl: string;
 }

export default class DocumentViewerWebPart extends BaseClientSideWebPart<IDocumentViewerWebPartProps> {

  private itemId: number;
  private commonService: CommonService;

  public render(): void {
    const queryParameters = new UrlQueryParameterCollection(window.location.href);

    if(this.itemId == null){
      if (queryParameters.getValue("ID")) {
        this.itemId = parseInt(queryParameters.getValue("ID"));
      }    
    }

    console.log("itemId "+this.itemId);
    this.getProjectDocuments(this.itemId);
    const element: React.ReactElement<IDocumentViewerProps > = React.createElement(
      DocumentViewer,
      {
        description: this.properties.description,
        documentUrl: this.properties.documentUrl,
        documents: this.properties.documents,
        documentsLength: this.properties.documentsLength,
        siteUrl: this.properties.siteUrl
      }
    );

    ReactDom.render(element, this.domElement);
  }

  private getProjectDocuments(itemId): void {
    this.commonService = new CommonService({
      spHttpClient: this.context.spHttpClient,
      siteUrl: this.context.pageContext.web.absoluteUrl,
      listName: `Approval%20Requests`,
      user: this.context.pageContext.user.loginName
    });

    this.properties.siteUrl = this.context.pageContext.web.absoluteUrl;
    this.commonService.getDocumentsList(itemId)
      .then((response) => {
        console.log("************************* response ************************* ");
        console.log(response);
        this.properties.documents = response['value'];
        this.properties.documentsLength = response.value.length;
        console.log("Documents Length: " + this.properties.documents.length);
        console.log("************************* response ************************* ");
    });
  }
  
  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
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
                  label: strings.DescriptionFieldLabel,
                })
              ]
            }
          ]
        }
      ]
    };
  }
}

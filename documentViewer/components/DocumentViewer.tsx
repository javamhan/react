import * as React from 'react';
import styles from './DocumentViewer.module.scss';
import { IDocumentViewerProps } from './IDocumentViewerProps';
//import { escape } from '@microsoft/sp-lodash-subset';

export default class DocumentViewer extends React.Component<IDocumentViewerProps, {}> {
  private static readonly DISPLAY_TYPE_WOPIIFRAME = "WopiFrame.aspx";
  private static readonly DISPLAY_TYPE_DOCX = "Doc.aspx";
  private static readonly ACTION_TYPE_DEFAULT = "&action=default";
  private static readonly ACTION_TYPE_INTERACTIVE = "&action=interactivepreview";
  private docSource : string;


  public render(): React.ReactElement<IDocumentViewerProps> {
    this.docSource = this.createSrc(this.props.documents, 0); // Get the first document as default view
    return (
      <div className={ styles.documentViewer }>
        {/** Just load the first document*/}       
        <iframe width="700" height="500" src={this.docSource} title="Embedded document" role="document"></iframe>        

        <table>
          <tr>
            {this.createTabs(this.props.documents)}
          </tr>
        </table>
      </div>
    );
  }

  private createSrc(documents: any[], i: number) {    
    this.docSource = this.props.siteUrl+"/teams/app_operations/_layouts/15/"
      +DocumentViewer.DISPLAY_TYPE_WOPIIFRAME
      +"?sourcedoc="
      +documents[i].ServerRelativeUrl
      +((documents[i].FileName.indexOf(".pdf") > -1) ? DocumentViewer.ACTION_TYPE_INTERACTIVE : DocumentViewer.ACTION_TYPE_DEFAULT);
      
    return this.docSource;
  }

  /** Dynamically create document tabs */
  //createTabs = () => {
  private createTabs(documents: any[]) {
    let tabs = [];
    let len = documents.length;
    let docDisplayname: string;
    for (let i = 0; i < len; i++) {
      docDisplayname = documents[i].FileName.length > 8 ? documents[i].FileName.substring(0, 8) : documents[i].FileName;
      tabs.push(<td> 
        {
          <a target= '_blank' href={documents[i].ServerRelativeUrl} className={ styles.button }>
            <span className={ styles.label }>{docDisplayname}...</span>
          </a>
        }
      </td>);
    }
    return tabs;
  }
  
}

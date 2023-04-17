import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './TestEventListnerWebPart.module.scss';
import * as strings from 'TestEventListnerWebPartStrings';

export interface ITestEventListnerWebPartProps {
  description: string;
  divisionName: string;
  teamName: string;
  docLibrary: string;
  dataResults: any[];
  asmResults: any[];
  cenResults: any[];
  divisions:string[];
  siteArray: any;
  siteName: string;
  siteTitle: string;
  isDCPowerUser:boolean;
  folderArray:any[];  
}

export default class TestEventListnerWebPart extends BaseClientSideWebPart<ITestEventListnerWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
 
  private getFolders() : void {
    let folderContainer : Element;
    let folderHTML : string = "";

    this.properties.folderArray=["Folder A","Folder B","Folder C","Folder D"];
    for(let x=0;x<this.properties.folderArray.length;x++){
      let folderID=this.properties.folderArray[x].replace(/\s+/g, "");

      folderHTML += `<div class="row">
                      <h2 class="accordion-header">
                        <button class="btn btn-primary" id="${folderID}" type="button" data-bs-toggle="collapse" aria-expanded="true"> 
                          <i class="bi bi-folder2"></i><a href="#" class="text-white ms-1">${this.properties.folderArray[x]}</a>                    
                        </button>
                      </h2>
                    </div>`;
    }
    folderContainer = this.domElement.querySelector('#folders');
    folderContainer.innerHTML = folderHTML;
    this._setButtonEventHandlers(); 
  }

  private _setButtonEventHandlers(): void {  
    let tabBtns=this.domElement.querySelectorAll('#btnCustomTab')
    tabBtns.forEach(function(tabBtns) {
      tabBtns.addEventListener('click', () => {  
        var headtext =  tabBtns.innerHTML;  
        alert(headtext);  
      });
    })
    
    for(let x=0;x<this.properties.folderArray.length;x++){        
      //if(x %2 === 0 ){
        let folderID=this.properties.folderArray[x].replace(/\s+/g, "");
        let folderName=this.properties.folderArray[x];
        //folderItemArray = this.properties.folderArray[x+1];
        console.log(folderID+" "+folderName);
      //}
      //folderNametemp=document.getElementById(folderIDtemp).innerHTML
      document.getElementById(folderID).addEventListener("click",(_e:Event) => alert(folderName));
    }     
  } 

  public render(): void {

    this.domElement.innerHTML = `
    <section class="${styles.testEventListner} ${!!this.context.sdks.microsoftTeams ? styles.teams : ''}">
      <div class="${styles.welcome}">
        <img alt="" src="${this._isDarkTheme ? require('./assets/welcome-dark.png') : require('./assets/welcome-light.png')}" class="${styles.welcomeImage}" />
        <h2>Well done, ${escape(this.context.pageContext.user.displayName)}!</h2>
        <div>${this._environmentMessage}</div>
        <div>Web part property value: <strong>${escape(this.properties.description)}</strong></div>
      </div>
      <div>
        <h3>Welcome to SharePoint Framework!</h3>
        <p>
        The SharePoint Framework (SPFx) is a extensibility model for Microsoft Viva, Microsoft Teams and SharePoint. It's the easiest way to extend Microsoft 365 with automatic Single Sign On, automatic hosting and industry standard tooling.
        </p>
        <h4>Learn more about SPFx development:</h4>
          <ul class="${styles.links}">
            <li><a href="https://aka.ms/spfx" target="_blank">SharePoint Framework Overview</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-graph" target="_blank">Use Microsoft Graph in your solution</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-teams" target="_blank">Build for Microsoft Teams using SharePoint Framework</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-viva" target="_blank">Build for Microsoft Viva Connections using SharePoint Framework</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-store" target="_blank">Publish SharePoint Framework applications to the marketplace</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-api" target="_blank">SharePoint Framework API reference</a></li>
            <li><a href="https://aka.ms/m365pnp" target="_blank">Microsoft 365 Developer Community</a></li>
          </ul>
      </div>
      <button type="button" value="tab1" id="btnCustomTab">Tab 1</button>
      <button type="button" value="tab2" id="btnCustomTab">Tab 2</button>
      <button type="button" value="tab3" id="btnCustomTab">Tab 3</button>
      <div id="folders"></div>
    </section>`;
    this.getFolders();
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }

  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              throw new Error('Unknown host');
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

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

import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle,
  PropertyPaneSlider,
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'PeopleSlickWebPartStrings';
import PeopleSlick from './components/PeopleSlick';
import { IPeopleSlickProps } from './components/IPeopleSlickProps';

export interface IPeopleSlickWebPartProps {
  description: string;
  listName: string;
  webpartName: string;
  UseRootSite: boolean;
   slickMode: string;
   minHeight:number;
   minHeightCarousell:number;
   photoWidth:number;
  showDots: boolean;
  autoplaySpeed: number;
  speed: number;
  slidesToShow: number;
  slidesToScroll: number;
  recordToReturn: number;
  enableAutoplay: boolean;
   rows: number;
slidesPerRow: number;
 
centerPadding:number;
centerMode: boolean;
infinite:boolean;
  customFilter: boolean;
  customFilterValue: string;
  enableRedirectURL: boolean;
  displayJobTitle:boolean;
 displayOffice:boolean;
 borderRadius:number;
 enableReactions:boolean;
 enableTeams:boolean;
}
export default class PeopleSlickWebPart extends BaseClientSideWebPart<IPeopleSlickWebPartProps> {
  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = "";


  public render(): void {
    const element: React.ReactElement<IPeopleSlickProps> = React.createElement(
      PeopleSlick,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        rootSiteURL: this.context.pageContext.site.absoluteUrl,
        context: this.context,
        listName: this.properties.listName,
        webpartName: this.properties.webpartName,
        UseRootSite: this.properties.UseRootSite,
        recordToReturn: this.properties.recordToReturn,
        slickMode: this.properties.slickMode,
        showDots: this.properties.showDots,
        minHeight: this.properties.minHeight,
        minHeightCarousell: this.properties.minHeightCarousell,
        photoWidth:this.properties.photoWidth,
        autoplaySpeed: this.properties.autoplaySpeed,
        speed: this.properties.speed,
        slidesToShow: this.properties.slidesToShow,
        slidesToScroll: this.properties.slidesToScroll,
      
        enableAutoplay: this.properties.enableAutoplay,
         rows: this.properties.rows,
        slidesPerRow: this.properties.slidesPerRow,
         
        centerPadding: this.properties.centerPadding,
        centerMode: this.properties.centerMode,
        infinite: this.properties.infinite,
        customFilter: this.properties.customFilter,
        customFilterValue: this.properties.customFilterValue,
        enableRedirectURL: this.properties.enableRedirectURL,
          displayJobTitle:this.properties.displayJobTitle,
        displayOffice:this.properties.displayOffice,
        borderRadius:this.properties.borderRadius,
        enableReactions:this.properties.enableReactions,
        enableTeams:this.properties.enableTeams
      }
    );

    ReactDom.render(element, this.domElement);
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
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
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

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        { displayGroupsAsAccordion:true,
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: "Data Source Configuration",

              groupFields: [
                PropertyPaneTextField('webpartName', {
                  label: "Webpart Name"
                }),

                PropertyPaneTextField('listName', {
                  label: "List Name"
                }),

                PropertyPaneToggle("UseRootSite", {
                  label: "Use Root Site?",
                  offText: "No",
                  onText: "Yes",
                }),

                   PropertyPaneSlider("recordToReturn", {
                  label: "Record to return",
                  min: 1,
                  max: 50,
                }),
                   ]
            },
            {
              groupName:"Slick Configuration",
                isCollapsed:true,
                groupFields:[
                
              
                PropertyPaneToggle("showDots", {
                  label: "Show navigation (Dots)",
                  offText: "No",
                  onText: "Yes",
                }),

                PropertyPaneSlider("minHeight", {
                  label: "minimum height (100-600)",
                  min: 100,
                  max: 600,
                 
                }),
                  PropertyPaneSlider("minHeightCarousell", {
                  label: "minimum height Carousell (100-600)",
                  min: 100,
                  max: 600,
                 
                }),

                   PropertyPaneSlider("photoWidth", {
                  label: "Photo Width (50-350)",
                  min: 50,
                  max: 350,
                 
                }),
                  PropertyPaneSlider("borderRadius", {
                  label: "Border Radius (0-300)",
                  min: 0,
                  max: 300,
                 
                }),


                PropertyPaneSlider("slidesToShow", {
                  label: "Slides to show",
                  min: 1,
                  max: 15,
                }),

                PropertyPaneSlider("slidesToScroll", {
                  label: "Slides to scroll",
                  min: 1,
                  max: 15,
                }),

                PropertyPaneSlider("rows", {
                  label: "Row to show",
                  min: 1,
                  max: 5,
                }),

                PropertyPaneSlider("slidesPerRow", {
                  label: "Slides Per Row",
                  min: 1,
                  max: 5,
                }),

           

                  PropertyPaneSlider('centerPadding', {
                    label: "Center Padding (1px - 60px)",
                     min: 1,
                     max: 60,
                }),
                
                PropertyPaneToggle("centerMode", {
                  label: "center Mode?",
                  offText: "No",
                  onText: "Yes",
                }),
                
                PropertyPaneToggle("infinite", {
                  label: "Enable infinite?",
                  offText: "No",
                  onText: "Yes",
                }),
                PropertyPaneSlider("speed", {
                  label: "Speed, Default : 5",
                  min: 1, 
                  max: 20,
                
                }),
                PropertyPaneToggle("enableAutoplay", {
                  label: "Enable autoplay?",
                  offText: "No",
                  onText: "Yes",
                }),
            
                 
                PropertyPaneSlider("autoplaySpeed", {
                  label: "Autoplay speed, default : 20",
                  min: 1, 
                  max: 50,
                  disabled: !this.properties.enableAutoplay,
                }),



              ]
            },
            {
              groupName: "Advanced configuration",

              groupFields: [
                PropertyPaneToggle("customFilter", {
                  label: "Use Custom Filter?",
                  offText: "No",
                  onText: "Yes",
                }),

                PropertyPaneTextField('customFilterValue', {
                  label: "Custom Filter Query, Eg: Tags eq 'APAC'",
                  disabled: !this.properties.customFilter,
                }),

                PropertyPaneToggle("enableRedirectURL", {
                  label: "enable RedirectURL?",
                  offText: "No",
                  onText: "Yes",
                }),
                  PropertyPaneToggle("displayJobTitle", {
                  label: "display Job Title & Department?",
                  offText: "No",
                  onText: "Yes",
                  }),
                
                  PropertyPaneToggle("displayOffice", {
                  label: "display Office?",
                  offText: "No",
                  onText: "Yes",
                  }),

                  PropertyPaneToggle("enableReactions", {
                  label: "Enable Reactions (👋 👏 🥳)?",
                  offText: "No",
                  onText: "Yes",
                  }),

                  PropertyPaneToggle("enableTeams", {
                  label: "Enable Teams Integration?",
                  offText: "No",
                  onText: "Yes",
                  }),

              ]

            },

          ]
        }
      ]
    };
  }
}

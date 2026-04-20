import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IPeopleSlickProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  rootSiteURL: string;
  listName: string;
  webpartName: string;
  context: WebPartContext;
  UseRootSite: boolean;
   slickMode: string;
   minHeight:number;
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
}

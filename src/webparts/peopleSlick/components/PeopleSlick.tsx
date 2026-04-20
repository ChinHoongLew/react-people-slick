import * as React from 'react';
import styles from './PeopleSlick.module.scss';
import type { IPeopleSlickProps } from './IPeopleSlickProps';
import { IconButton } from "office-ui-fabric-react";
// PnP Js
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/sites";
import { SPFx, spfi } from "@pnp/sp";
import { Web } from "@pnp/sp/webs";

import Slider from "react-slick";
import "slick-carousel/slick/slick.css";
import "slick-carousel/slick/slick-theme.css";

// Interface of list columns. Name must match with Sharepoint column internal name
interface CarousalItem {
  Id: number;

  Email: {
    Title: string;
    JobTitle: string;
    EMail: string;
    Department: string;
    Office: string;
  }

  RedirectURL: {
    Url: string;
  }

}


interface IState {
  listItems: CarousalItem[];

  loading: boolean;
}

export default class PeopleSlick extends React.Component<IPeopleSlickProps, IState> {
  constructor(props: IPeopleSlickProps) {
    super(props);
    this.state = {
      loading: true,

      listItems: [],
    };
  }


  public async componentDidMount(): Promise<undefined> {
    await this.getDataFromList();
    return;
  }

  private async getDataFromList(): Promise<undefined> {
    console.log("Getting data from list");
    try {

      const sp = await spfi().using(SPFx(this.props.context));
      let filterText = ""

      if (this.props.customFilter) {
        filterText = this.props.customFilterValue;
      }


      if (this.props.UseRootSite) {
        const originWeb = window.location.origin;
        const web1 = Web(originWeb).using(SPFx(this.props.context));
        const items1 = await web1.lists
          .getByTitle(this.props.listName)
          .items.expand("Email")
          .select("Published,RedirectURL,Email/Title,Email/JobTitle,Email/EMail,Email/Department,Email/Office")
          .top(this.props.recordToReturn)
          .filter(filterText)

          .orderBy("Published", false)();



        this.setState({
          listItems: items1,
          loading: false
        });

      } else {

        const items = await sp.web.lists
          .getByTitle(this.props.listName)
          .items.expand("Email")
          .select("Published,RedirectURL,Email/Title,Email/JobTitle,Email/EMail,Email/Department,Email/Office")
          .top(this.props.recordToReturn)
          .filter(filterText)

          .orderBy("Published", false)();

        this.setState({
          listItems: items,
          loading: false
        });
      }

    } catch (error) { console.log((error as Error).message); }






    return;
  }



  public render(): React.ReactElement<IPeopleSlickProps> {
    const settings = {
      dots: this.props.showDots,
      infinite: this.props.infinite,
      speed: this.props.speed * 100,
      slidesToShow: this.props.slidesToShow,
      slidesToScroll: this.props.slidesToScroll,
      autoplay: this.props.enableAutoplay,
      autoplaySpeed: this.props.autoplaySpeed * 1000,
      rows : this.props.rows,
      slidesPerRow : this.props.slidesPerRow,
      adaptiveHeight: true,
      className: "",
      cssEase: "linear",
      responsive: [
        {
          breakpoint: 1024,
          settings: {
            slidesToShow: 3,
          },
        },
        {
          breakpoint: 600,
          settings: {
            slidesToShow: 2,
          },
        },
        {
          breakpoint: 480,
          settings: {
            slidesToShow: 1,
          },
        },
      ],
    };

      const styleBlock = { "--minHeight": this.props.minHeight + "px"} as React.CSSProperties;
      const styleBlock2 = { "--borderRadius": this.props.borderRadius + "px"} as React.CSSProperties;
    return (
      <section className={`${styles.peopleSlick} `} style={styleBlock}>
        {this.state.loading && <p>Loading...</p>}
        <div className={styles.mainContainer}><p className={styles.webpartName}>{this.props.webpartName}</p>
          <Slider {...settings}>
            {this.state.listItems.map((item: CarousalItem) => {
              return (
                <div className={styles.carousalItem} key={item.Id}>
                  <p className={styles.profile}> <img style={styleBlock2} width={`${this.props.photoWidth}`} src={`${this.props.rootSiteURL}/_layouts/15/userphoto.aspx?size=L&accountname=${item.Email.EMail}`}  title={item.Email.Title} />
                
                  </p>
                  <p className={styles.title}>{item.Email.Title}  <IconButton
                    iconProps={{ iconName: "TeamsLogo" }}
                    title="Teams"
                    onClick={(event) => {
                      event.stopPropagation();
                      window.open(`https://teams.microsoft.com/l/chat/0/0?users=${item.Email.EMail}&topicName=Hello&message=Hello%20${item.Email.Title}`);
                    
                    }}
                  /></p>
                  {this.props.displayJobTitle &&(<p className={styles.description}>{item.Email.JobTitle}, {item.Email.Department}</p>)}
                  {this.props.displayOffice &&(<p className={styles.office}>{item.Email.Office}</p>)}
                  {this.props.enableRedirectURL && item.RedirectURL && (
                    <p className={styles.viewMoreP}>
                      <button
                        className={styles.viewMore}
                        onClick={() => {
                          window.open(item.RedirectURL.Url, "_blank");
                        }}
                      >
                        Read more
                      </button>
                    </p>
                  )}
                </div>
              );
            })}
          </Slider>
        </div>
      </section>
    );
  }
}

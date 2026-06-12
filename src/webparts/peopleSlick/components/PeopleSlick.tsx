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

const REACTION_EMOJIS = ["👋", "👏", "🥳"] as const;
type ReactionEmoji = typeof REACTION_EMOJIS[number];

interface ReactionEntry {
  email: string;
  emoji: ReactionEmoji;
}

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

  reaction: string;
}


interface IState {
  listItems: CarousalItem[];
  loading: boolean;
  reactionMap: { [itemId: number]: ReactionEntry[] };
  updatingReaction: { [itemId: number]: boolean };
}

export default class PeopleSlick extends React.Component<IPeopleSlickProps, IState> {
  constructor(props: IPeopleSlickProps) {
    super(props);
    this.state = {
      loading: true,
      listItems: [],
      reactionMap: {},
      updatingReaction: {},
    };
  }


  public async componentDidMount(): Promise<undefined> {
    await this.getDataFromList();
    return;
  }

  private parseReactions(reactionStr: string): ReactionEntry[] {
    if (!reactionStr) return [];
    try {
      return JSON.parse(reactionStr) as ReactionEntry[];
    } catch (error) {
      console.error("Error parsing reactions:", error);
      return [];
    }
  }

  private findReaction(reactions: ReactionEntry[], email: string): ReactionEntry | undefined {
    const found = reactions.filter((r: ReactionEntry) => r.email === email);
    return found.length > 0 ? found[0] : undefined;
  }

  private async getDataFromList(): Promise<undefined> {
    console.log("Getting data from list");
    try {

      const sp = await spfi().using(SPFx(this.props.context));
      let filterText = ""

      if (this.props.customFilter) {
        filterText = this.props.customFilterValue;
      }

      const selectFields = "Id,Published,RedirectURL,reaction,Email/Title,Email/JobTitle,Email/EMail,Email/Department,Email/Office";

      const buildReactionMap = (items: CarousalItem[]): { [id: number]: ReactionEntry[] } => {
        const map: { [id: number]: ReactionEntry[] } = {};
        items.forEach(item => { map[item.Id] = this.parseReactions(item.reaction); });
        return map;
      };

      if (this.props.UseRootSite) {
        const originWeb = window.location.origin;
        const web1 = Web(originWeb).using(SPFx(this.props.context));
        const items1 = await web1.lists
          .getByTitle(this.props.listName)
          .items.expand("Email")
          .select(selectFields)
          .top(this.props.recordToReturn)
          .filter(filterText)
          .orderBy("Published", false)();

        this.setState({
          listItems: items1,
          loading: false,
          reactionMap: buildReactionMap(items1),
        });

      } else {

        const items = await sp.web.lists
          .getByTitle(this.props.listName)
          .items.expand("Email")
          .select(selectFields)
          .top(this.props.recordToReturn)
          .filter(filterText)
          .orderBy("Published", false)();

        this.setState({
          listItems: items,
          loading: false,
          reactionMap: buildReactionMap(items),
        });
      }

    } catch (error) { console.log((error as Error).message); }

    return;
  }

  private async handleReaction(itemId: number, emoji: ReactionEmoji): Promise<void> {
    const currentUserEmail = this.props.context.pageContext.user.email;
    const currentReactions = this.state.reactionMap[itemId] || [];
    const userExisting = this.findReaction(currentReactions, currentUserEmail);

    let updatedReactions: ReactionEntry[];
    if (userExisting && userExisting.emoji === emoji) {
      // Toggle off
      updatedReactions = currentReactions.filter(r => r.email !== currentUserEmail);
    } else {
      // Add or replace
      updatedReactions = currentReactions.filter(r => r.email !== currentUserEmail);
      updatedReactions.push({ email: currentUserEmail, emoji });
    }

    // Optimistic update
    this.setState(prev => ({
      reactionMap: { ...prev.reactionMap, [itemId]: updatedReactions },
      updatingReaction: { ...prev.updatingReaction, [itemId]: true },
    }));

    try {
      const reactionJson = JSON.stringify(updatedReactions);
      if (this.props.UseRootSite) {
        const originWeb = window.location.origin;
        const web1 = Web(originWeb).using(SPFx(this.props.context));
        await web1.lists.getByTitle(this.props.listName).items.getById(itemId).update({ reaction: reactionJson });
      } else {
        const sp = spfi().using(SPFx(this.props.context));
        await sp.web.lists.getByTitle(this.props.listName).items.getById(itemId).update({ reaction: reactionJson });
      }
    } catch (error) {
      console.log((error as Error).message);
      // Revert on failure
      this.setState(prev => ({
        reactionMap: { ...prev.reactionMap, [itemId]: currentReactions },
      }));
    } finally {
      this.setState(prev => ({
        updatingReaction: { ...prev.updatingReaction, [itemId]: false },
      }));
    }
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

      const styleBlock = { "--minHeight": this.props.minHeight + "px", "--minHeightCarousell": this.props.minHeightCarousell + "px" } as React.CSSProperties;
      //const styleBlock1 = { "--minHeightCarousell": this.props.minHeightCarousell + "px"} as React.CSSProperties;
      const styleBlock2 = { "--borderRadius": this.props.borderRadius + "px"} as React.CSSProperties;
    return (
      <section className={`${styles.peopleSlick} `} style={styleBlock}>
        {this.state.loading && <p>Loading...</p>}
        <div className={styles.mainContainer}><p className={styles.webpartName}>{this.props.webpartName}</p>
          <Slider {...settings}>
            {this.state.listItems.map((item: CarousalItem) => {
              const reactions = this.state.reactionMap[item.Id] || [];
              const currentUserEmail = this.props.context.pageContext.user.email;
              const userReaction = this.findReaction(reactions, currentUserEmail);
              const isUpdating = !!this.state.updatingReaction[item.Id];

              return (
                <div className={styles.carousalItem} key={item.Id} style={styleBlock}>
                  <p className={styles.profile}> <img style={styleBlock2} width={`${this.props.photoWidth}`} src={`${this.props.rootSiteURL}/_layouts/15/userphoto.aspx?size=L&accountname=${item.Email.EMail}`}  title={item.Email.Title} />

                  </p>
                  <p className={styles.title}>{item.Email.Title}  
                    
                    {this.props.enableTeams && (
                      <IconButton
                        iconProps={{ iconName: "TeamsLogo" }}
                        title="Teams"
                        onClick={(event) => {
                          event.stopPropagation();
                          window.open(`https://teams.microsoft.com/l/chat/0/0?users=${item.Email.EMail}&topicName=Hello&message=Hello%20${item.Email.Title}`);
                        }}
                      />)}

                  </p>
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
                  {this.props.enableReactions && (
                    <p className={styles.reactionContainer}>
                      {REACTION_EMOJIS.map(emoji => {
                        const count = reactions.filter(r => r.emoji === emoji).length;
                        const isActive = userReaction?.emoji === emoji;
                        return (
                          <button
                            key={emoji}
                            className={`${styles.reactionBtn} ${isActive ? styles.reactionBtnActive : ""}`}
                            onClick={() => { this.handleReaction(item.Id, emoji).catch(() => undefined); }}
                            disabled={isUpdating}
                            title={isActive ? "Remove reaction" : "React"}
                          >
                            <span>{emoji}</span>
                            {count > 0 && <span className={styles.reactionCount}>{count}</span>}
                          </button>
                        );
                      })}
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

import * as React from "react";
import { useState, useEffect } from "react";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { Swiper, SwiperSlide } from "swiper/react";
import { Navigation,Autoplay } from "swiper";
import "swiper/css";
import "swiper/css/navigation";
import { IDescriptionCardProps, ICardItem } from "./IDescriptionCardProps";
import styles from "./DescriptionCard.module.scss";

const DescriptionCard: React.FunctionComponent<IDescriptionCardProps> = (
  props,
) => {
  const {
    listName,
    siteUrl,
    spHttpClient,
    showTitle,
    iconSize,
    itemHeight,
    itemBgColor,
    titleFontSize,
    titleColor,
    descFontSize,
    descColor,
    enableSlider,
    desktopItems,
    tabletItems,
    mobileItems,
  } = props;

  const [items, setItems] = useState<ICardItem[]>([]);
  const [loading, setLoading] = useState<boolean>(true);
  const [error, setError] = useState<string>("");

  // Property pane values -> CSS variables
  const wrapperStyle = {
    "--card-min-height": `${itemHeight}px`,
    "--card-bg": itemBgColor,
    "--icon-size": `${iconSize}`,
    "--title-size": `${titleFontSize}px`,
    "--title-color": titleColor,
    "--desc-size": `${descFontSize}px`,
    "--desc-color": descColor,
  } as React.CSSProperties;

  useEffect(() => {
    let isMounted = true;

    const loadItems = async (): Promise<void> => {
      if (!listName) {
        if (isMounted) {
          setLoading(false);
          setError("Please enter a list name in the web part properties.");
        }
        return;
      }

      setLoading(true);
      setError("");

      try {
        const itemsUrl = `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listName)}')/items?$select=Id,Title,Description,ItemLink,OrderNumber&$orderby=OrderNumber asc`;
        const itemsResponse: SPHttpClientResponse = await spHttpClient.get(
          itemsUrl,
          SPHttpClient.configurations.v1,
        );

        if (!itemsResponse.ok) {
          throw new Error(
            `Failed to load list items (${itemsResponse.status} ${itemsResponse.statusText})`,
          );
        }

        const itemsJson = await itemsResponse.json();
        const rawItems: any[] = itemsJson.value || [];

        const cardItems: ICardItem[] = await Promise.all(
          rawItems.map(async (item) => {
            let iconUrl = "";
            try {
              const attachUrl = `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listName)}')/items(${item.Id})/AttachmentFiles`;
              const attachResponse: SPHttpClientResponse =
                await spHttpClient.get(
                  attachUrl,
                  SPHttpClient.configurations.v1,
                );

              if (attachResponse.ok) {
                const attachJson = await attachResponse.json();
                const attachments: any[] = attachJson.value || [];
                if (attachments.length > 0) {
                  iconUrl = attachments[0].ServerRelativeUrl;
                }
              }
            } catch {
              // no icon for this item
            }

const hyperlink = item.ItemLink; // Hyperlink column => { Url, Description }

return {
  id: item.Id,
  title: item.Title || "",
  description: item.Description || "",
  link: hyperlink?.Url || "",
  linkText: hyperlink?.Description || "",
  iconUrl,
} as ICardItem;
          }),
        );

        if (isMounted) {
          setItems(cardItems);
          setLoading(false);
        }
      } catch (err: any) {
        if (isMounted) {
          setError(
            err.message || "Something went wrong while loading the list.",
          );
          setLoading(false);
        }
      }
    };

    loadItems();

    return () => {
      isMounted = false;
    };
  }, [listName, siteUrl]);

  if (loading) {
    return <div>Loading...</div>;
  }

  if (error) {
    return <div className={styles.error}>{error}</div>;
  }

  if (items.length === 0) {
    return <div>No items found in this list.</div>;
  }

  // Single card: <a> when a link exists, otherwise <div>
  const renderCard = (item: ICardItem): JSX.Element => (
    // <div className={styles.cardDesc} tabIndex={0}>
    //   {React.createElement(
    //     item.link ? "a" : "div",
    //     {
    //       className: styles.cardInner,
    //       ...(item.link && {
    //         href: item.link,
    //         target: "_blank",
    //         rel: "noopener noreferrer",
    //         "data-interception": "off",
    //       }),
    //     },
    //     <>
    //       {/* Front */}
    //       <div className={`${styles.cardFace} ${styles.cardFront}`}>
    //         <div className={styles.iconDiv}>
    //           {item.iconUrl && (
    //             <img src={item.iconUrl} alt={`${item.title} Icon`} />
    //           )}
    //           {showTitle && <p className={styles.title}>{item.title}</p>}
    //         </div>
    //       </div>

    //       {/* Back */}
    //       <div className={`${styles.cardFace} ${styles.cardBack}`}>
    //         <p className={styles.desc}>{item.description}</p>
    //       </div>
    //     </>,
    //   )}
    // </div>
    <div className={styles.cardDesc} tabIndex={0}>
  <div className={styles.cardInner}>
    {/* Front */}
    <div className={`${styles.cardFace} ${styles.cardFront}`}>
      <div className={styles.iconDiv}>
        {item.iconUrl && (
          <img src={item.iconUrl} alt={`${item.title} Icon`} />
        )}
        {showTitle && <p className={styles.title}>{item.title}</p>}
      </div>
    </div>

    {/* Back */}
    <div className={`${styles.cardFace} ${styles.cardBack}`}>
      <p className={styles.desc}>{item.description}</p>
        {item.link && (
  <a
    href={item.link}
    target="_blank"
    rel="noopener noreferrer"
    data-interception="off"
  >
    {item.linkText}
  </a>
)}
    </div>
  </div>
</div> 
  );

  // Slider is active only when the toggle is on AND items exceed the desktop count
  const isSlider = enableSlider && items.length > desktopItems;

  return (
    <div className={styles.cardSliderWrap} style={wrapperStyle}>
      {isSlider ?
        <Swiper
          key={`${desktopItems}-${tabletItems}-${mobileItems}-${items.length}`}
          modules={[Navigation,Autoplay]}
          navigation={true}
          loop={true}
          spaceBetween={14}
          slidesPerView={mobileItems}
          autoplay={{ delay: 5000, disableOnInteraction: false, pauseOnMouseEnter: true, }}
          breakpoints={{
            640: { slidesPerView: tabletItems },
            1024: { slidesPerView: desktopItems },
          }}
          className={styles.swiper}
        >
          {items.map((item) => (
            <SwiperSlide key={item.id} className={styles.swiperSlide}>
              {renderCard(item)}
            </SwiperSlide>
          ))}
        </Swiper>
      : <div className={styles.cardDescWrapper}>
          {items.map((item) => (
            <React.Fragment key={item.id}>{renderCard(item)}</React.Fragment>
          ))}
        </div>
      }
    </div>
  );
};

export default DescriptionCard;

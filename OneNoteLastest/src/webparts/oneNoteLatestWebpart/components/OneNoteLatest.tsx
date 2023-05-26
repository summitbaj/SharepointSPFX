import * as React from 'react';
import styles from './OneNoteLatestWebpart.module.scss';
import { IOneNoteLatestWebpartProps } from './IOneNoteLatestProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

interface IOneNoteLatestWebpartState {
  latestFileUrl: string;
  linkingUrl: string;
}

export default class OneNoteLatestWebpart extends React.Component<IOneNoteLatestWebpartProps, IOneNoteLatestWebpartState> {

  constructor(props: IOneNoteLatestWebpartProps) {
    super(props);

    this.state = {
      latestFileUrl: '',
      linkingUrl: ''
    };
  }

  public async componentDidMount() {
    try {
      const { siteURL, folderRelativeURL } = this.props;
      
      // Retrieve the last created folder in the specified folderRelativeURL
      const folderApiUrl = `${siteURL}/_api/web/getFolderByServerRelativeUrl('${folderRelativeURL}')/folders?$orderby=TimeCreated desc&$top=1&$select=ServerRelativeUrl`;
      const folderResponse: SPHttpClientResponse = await this.props.spHttpClient.get(folderApiUrl, SPHttpClient.configurations.v1);

      const folderData = await folderResponse.json();
      const lastCreatedFolderUrl = folderData.value[0].ServerRelativeUrl;
      console.log(lastCreatedFolderUrl);
      // Retrieve the last created file inside the last created folder
      const fileApiUrl = `${siteURL}/_api/web/getFolderByServerRelativeUrl('${lastCreatedFolderUrl}')/files?$orderby=TimeCreated desc&$top=1&$select=ServerRelativeUrl`;
      const fileResponse: SPHttpClientResponse = await this.props.spHttpClient.get(fileApiUrl, SPHttpClient.configurations.v1);

      const fileData = await fileResponse.json();
      const latestFileUrl = fileData.value[0].ServerRelativeUrl;

      
      this.setState({ latestFileUrl });

      // Generate the linking URL for the latest file
      const fileLinkApiUrl = `${siteURL}/_api/web/GetFileByServerRelativeUrl('${latestFileUrl}')/LinkingUrl`;
      const fileLinkResponse: SPHttpClientResponse = await this.props.spHttpClient.post(fileLinkApiUrl, SPHttpClient.configurations.v1, {
        body: JSON.stringify({
          "request": {
            "createLink": true,
            "settings": {
              "expiration": "never",
              "password": "",
              "blockDownload": false
            }
          }
        })
      });
  
      const fileLinkData = await fileLinkResponse.json();
      const linkingUrl = fileLinkData.value;
      this.setState({ linkingUrl });

    } catch (error) {
      console.log(`Error getting latest file: ${error}`);
    }
  }
  

  public render(): React.ReactElement<IOneNoteLatestWebpartProps> {
    const {
      description,
      imgSrc,
    } = this.props;


    

    return (
      <div className={styles.imageSection}>
        <a href={this.state.linkingUrl} >
          <img src={imgSrc ? imgSrc : require('../assets/welcome-light.png')} alt={escape(description)} />
          <div className={styles.overlay}></div>
          <div className={styles.textOverlay}>{escape(description)}</div>
      </a>
    </div>
  );
  }
}

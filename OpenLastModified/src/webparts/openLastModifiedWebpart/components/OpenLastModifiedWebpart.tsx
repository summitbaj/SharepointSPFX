import * as React from 'react';
import styles from './OpenLastModifiedWebpart.module.scss';
import { IOpenLastModifiedWebpartProps } from './IOpenLastModifiedWebpartProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import "@pnp/sp/webs";
import "@pnp/sp/files";

interface IOpenLastModifiedWebpartState {
  latestFileUrl: string;
  linkingUrl: string;
}

export default class OpenLastModifiedWebpart extends React.Component<IOpenLastModifiedWebpartProps, IOpenLastModifiedWebpartState> {
  constructor(props: IOpenLastModifiedWebpartProps) {
    super(props);

    this.state = {
      latestFileUrl: '',
      linkingUrl: ''
    };
  }
  

  public async componentDidMount() {
    try {
      const { siteURL } = this.props;
      const { folderRelativeURL } = this.props
      const apiUrl = `${siteURL}/_api/web/getFolderByServerRelativeUrl('${folderRelativeURL}')/files?$orderby=TimeCreated desc&$top=1&$select=ServerRelativeUrl`;
      const response: SPHttpClientResponse = await this.props.spHttpClient.get(apiUrl, SPHttpClient.configurations.v1);
      const data = await response.json();
      const latestFileUrl = data.value[0].ServerRelativeUrl;
      this.setState({ latestFileUrl });
  
      const fileApiUrl = `${siteURL}/_api/web/GetFileByServerRelativeUrl('${latestFileUrl}')/LinkingUrl`;
      const fileResponse: SPHttpClientResponse = await this.props.spHttpClient.post(fileApiUrl, SPHttpClient.configurations.v1, {
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
  
      const fileData = await fileResponse.json();
      console.log(fileData);
      const linkingUrl = fileData.value;
      this.setState({ linkingUrl });
    } catch (error) {
      console.log(`Error getting latest file: ${error}`);
    }
  }

  public render(): React.ReactElement<IOpenLastModifiedWebpartProps> {
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

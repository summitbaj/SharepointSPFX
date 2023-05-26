import * as React from 'react';
import axios from 'axios';

export interface ILocationData {
  address: string;
  town: string;
  state: string;
  listPrice: string;
  listingAgent: string;
  listingService: string;
  contact: string;
  longitude: string;
  latitude: string;
}

export interface IMapProps {
  apiKey: string;
  listUrl: string;
  listName: string;
}

export interface IMapState {
  map?: google.maps.Map;
  infoWindow: google.maps.InfoWindow | null;
}

export default class Map extends React.Component<IMapProps, IMapState> {
  private mapContainerRef: React.RefObject<HTMLDivElement>;

  constructor(props: IMapProps) {
    super(props);
    this.mapContainerRef = React.createRef<HTMLDivElement>();
    this.state = {
      infoWindow: null,
    };
  }

  componentDidMount() {
    const script = document.createElement('script');
    script.src = `https://maps.googleapis.com/maps/api/js?key=${this.props.apiKey}&libraries=places`;
    script.onload = this.initializeMap;
    document.body.appendChild(script);
  }

  initializeMap = () => {
    const mapOptions: google.maps.MapOptions = {
      zoom: 10,
    };

    const map = new google.maps.Map(this.mapContainerRef.current!, mapOptions);
    const infoWindow = new google.maps.InfoWindow();

    this.setState({ map, infoWindow }, this.fetchDataAndPlotMarkers);
  };

  fetchDataAndPlotMarkers = () => {
    const { listUrl, listName } = this.props;

    axios
      .get(`${listUrl}/_api/web/lists/getbytitle('${listName}')/items`, {
        headers: {
          Accept: 'application/json;odata=nometadata',
        },
      })
      .then((response) => {
        const data = response.data.value;

        if (data.length > 0) {
          const firstLocation: ILocationData = {
            address: data[0].Title,
            town: data[0].Town,
            state: data[0].State,
            listPrice: data[0]['ListPrice'],
            listingAgent: data[0]['ListingAgent'],
            listingService: data[0]['ListingService'],
            contact: data[0].Contact,
            longitude: data[0].Longitude,
            latitude: data[0].Latitude,
          };

          this.plotMarker(firstLocation, true);

          data.slice(1).forEach((item: any) => {
            const location: ILocationData = {
              address: item.Title,
              town: item.Town,
              state: item.State,
              listPrice: item['List_x0020_Price'],
              listingAgent: item['Listing_x0020_Agent'],
              listingService: item['Listing_x0020_Service'],
              contact: item.Contact,
              longitude: item.Longitude,
              latitude: item.Latitude,
            };

            this.plotMarker(location);
          });
        }
      })
      .catch((error) => {
        console.error('Error fetching data from SharePoint:', error);
      });
  };

  plotMarker = (location: ILocationData, isCenter: boolean = false) => {
    const { map, infoWindow } = this.state;

    const latitude = parseFloat(location.latitude);
    const longitude = parseFloat(location.longitude);

    if (map && infoWindow) {
 
      const marker = new google.maps.Marker({
        position: { lat: latitude, lng: longitude },
        map,
      });

      marker.addListener('click', () => {
        this.showMarkerDetails(marker, location);
      });

      if (isCenter) {
        map.setCenter(marker.getPosition()!);
      }
    }
  };

  showMarkerDetails = (marker: google.maps.Marker, location: ILocationData) => {
    const { infoWindow } = this.state;

    if (infoWindow) {
      const infoWindowContent = `
        <div style="color: black; font-size:10px;">
          <h4>${location.address}</h4>
          <p>Town: ${location.town}</p>
          <p>State: ${location.state}</p>
          <p>List Price: ${location.listPrice}</p>
          <p>Listing Agent: ${location.listingAgent}</p>
          <p>Listing Service: ${location.listingService}</p>
          <p>Contact: ${location.contact}</p>
        </div>
      `;

      infoWindow.setContent(infoWindowContent);
      infoWindow.open(this.state.map!, marker);
    }
  };

  render() {
    return <div ref={this.mapContainerRef} style={{ width: '100%', height: '400px' }} />;
  }
}

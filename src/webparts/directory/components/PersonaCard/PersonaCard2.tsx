import * as React from 'react';
import styles from './PersonaCard.module.scss';
import { IPersonaCardProps } from './IPersonaCardProps';
import { IPersonaCardState } from './IPersonaCardState';
import {Log} from '@microsoft/sp-core-library';
import Rating from '@material-ui/lab/Rating';
import Grid from '@material-ui/core/Grid';
import { SPComponentLoader } from '@microsoft/sp-loader';
import Card from 'react-bootstrap/Card';
import 'bootstrap/dist/css/bootstrap.min.css';
import Divider from '@material-ui/core/Divider';
import RoomIcon from '@material-ui/icons/Room';
import GoogleMapReact from 'google-map-react';
import { mergeStyleSets } from 'office-ui-fabric-react/lib/Styling';
import {Facepile, IFacepilePersona, OverflowButtonType} from 'office-ui-fabric-react/lib/Facepile';
import Axios from 'axios';
import {Image,} from 'office-ui-fabric-react';

const hotelpic:string=require('../avataar/Hotel.jpg');
const space = ' ';
const EXP_SOURCE: string = 'SPFxDirectory';
const LIVE_PERSONA_COMPONENT_ID: string ='914330ee-2df2-4f6e-a858-30c23a812408';

export class PersonaCard extends React.Component<IPersonaCardProps,IPersonaCardState>
 {
  constructor(props: IPersonaCardProps) {
    super(props);
    this.state = {
      livePersonaCard: undefined,
      pictureUrl: undefined,
      title: undefined,
      location: undefined,
      phone: undefined,
      name: undefined,
      op: undefined,
      status: undefined,
      lat: undefined,
      lng: undefined,
      hotels: undefined,
    };
  }

  /**
   *
   *
   * @memberof PersonaCard
   */

  public async componentDidMount() {
    const sharedLibrary = await this._loadSPComponentById(
      LIVE_PERSONA_COMPONENT_ID
    );
    const livePersonaCard: any = sharedLibrary.LivePersonaCard;
    this.setState({ livePersonaCard: livePersonaCard });
    this.Add();
  }
  /**
   *
   *
   * @param {IPersonaCardProps} prevProps
   * @param {IPersonaCardState} prevState
   * @memberof PersonaCard
   */

  public componentDidUpdate(
    prevProps: IPersonaCardProps,
    prevState: IPersonaCardState
  ): void {}

  /**
   *
   *
   * @private
   * @returns
   * @memberof PersonaCard
   */

  private _LivePersonaCard() {
    return React.createElement(
      this.state.livePersonaCard,
      {
        serviceScope: this.props.context.serviceScope,
        upn: this.props.profileProperties.Email,
        onCardOpen: () => {
          console.log('LivePersonaCard Open');
        },
        onCardClose: () => {
          console.log('LivePersonaCard Close');
        },
      },
      this._PersonaCard()
    );
  }

  /**
   *
   *
   * @private
   * @returns {JSX.Element}
   * @memberof PersonaCard
   */

  useStyles = mergeStyleSets({
    root: {
      height: 330,
      width: 200,
    },
    media: {
      height: 120,
    },
  });

  public getCity(): Promise<string> {
    return new Promise((resolve, reject) =>
      Axios.get(
        `https://cors-anywhere.herokuapp.com/https://maps.googleapis.com/maps/api/place/findplacefromtext/json?input=${this.props.profileProperties.DisplayName}&inputtype=textquery&fields=geometry&key=AIzaSyD7uFU7U6uumrFjm2wH7TZ8iFUxWwMcqpI`
      ).then((resp) => {resolve(resp.data.candidates[0]);})
    );
  }

  public GetAddress(): Promise<string> {
    return new Promise((resolve, reject) =>
      Axios.get(
        `https://cors-anywhere.herokuapp.com/https://maps.googleapis.com/maps/api/place/findplacefromtext/json?input=Cyclotron ${this.props.profileProperties.DisplayName}&inputtype=textquery&fields=formatted_address,business_status,opening_hours,geometry&key=AIzaSyD7uFU7U6uumrFjm2wH7TZ8iFUxWwMcqpI`
      )
        .then((resp) => {resolve(resp.data.candidates[0]);})
        .catch(() => {
          this.setState({
            op: 'Status Not Available',
            status: 'Closed Temporarily',
            location: 'Address Not Available',
          });
        })
    );
  }

  public async Add() {
    let r: any = await this.GetAddress();

    if (r != null) {

      //Address
      if (r.formatted_address != null) {this.setState({ location: r.formatted_address });} 
      else {this.setState({ location: 'Address Not Available' });}

      //Open or Not
      if (r.opening_hours != null) {
        if (r.opening_hours.open_now == true) {this.setState({ op: 'Open Now' });}
        else if (r.opening_hours.open_now == false) {this.setState({ op: 'Closed Now' });}
      } 
      else {this.setState({ op: 'Status Not Available' });}

      //Latitude and Longtitude
      if (r.geometry != null) {
        this.setState({ lat: r.geometry.location.lat });
        this.setState({ lng: r.geometry.location.lng });
      }

      //Open or closed
      if (r.business_status != null) {
        if (r.business_status == 'OPERATIONAL') {this.setState({ status: 'Operational' });} 
        else {this.setState({ status: 'Closed Temporarily' });}
      }
      else {this.setState({ status: 'Status Not Available' });}

    } 
    else {
      let s: any = await this.getCity();
      this.setState({
        op: 'Status Not Available',
        status: 'Status Not Available',
        location: 'Address Not Available',
        lat: s.geometry.location.lat,
        lng: s.geometry.location.lng,
      });
    }

    Axios.get(
      `https://cors-anywhere.herokuapp.com/https://maps.googleapis.com/maps/api/place/nearbysearch/json?location=${this.state.lat},${this.state.lng}&radius=2000&type=restaurant&key=AIzaSyD7uFU7U6uumrFjm2wH7TZ8iFUxWwMcqpI`
    ).then((res) => {
      this.hotels(res.data.results);
    });
  }

  private _PersonaCard(): JSX.Element {
    let AnyReactComponent = (lat, lng) => <RoomIcon fontSize={'large'} />;

    Axios.get(
      `https://cors-anywhere.herokuapp.com/https://maps.googleapis.com/maps/api/place/findplacefromtext/json?input=${this.props.profileProperties.DisplayName}&inputtype=textquery&fields=photos&key=AIzaSyD7uFU7U6uumrFjm2wH7TZ8iFUxWwMcqpI`
    )
      .then((response) => {
        this.getPhoto(response.data.candidates[0].photos[0].photo_reference);
      })
      .catch((error) => {
        console.log(error);
      });

    const getPersonaProps = (persona: IFacepilePersona) => ({
      hidePersonaDetails: true,
    });
    const overflowButtonProps = {
      ariaLabel: 'More users',
    };
    
    return (
      <div className={styles.hmm}>
        <Grid container spacing={2} direction={'column'}>
          <Grid item>
            <Card
              onCardOpen
              border="light"
              text="dark"
              style={{ width: '350px', height: '18rem' }}
            >
              <Card.Title>
                {this.props.profileProperties.DisplayName}
              </Card.Title>
              <Card.Img
                height={200}
                variant="top"
                src={this.state.pictureUrl}
              />
            </Card>
          </Grid>

          <h6 className="pl-3">Location</h6>
          <Grid item>
            <Grid
              className={styles.grid1}
              container
              spacing={2}
              direction={'row'}
            >
              <Grid item>
                <div style={{ height: '180px', width: '170px' }}>
                  <GoogleMapReact
                    bootstrapURLKeys={{
                      key: 'AIzaSyD7uFU7U6uumrFjm2wH7TZ8iFUxWwMcqpI',
                    }}
                    center={{ lat: this.state.lat, lng: this.state.lng }}
                    defaultZoom={18}
                  >
                    <AnyReactComponent
                      lat={
                        this.state.lat != undefined
                          ? this.state.lat
                          : -8.783194999999999
                      }
                      lng={
                        this.state.lng != undefined
                          ? this.state.lng
                          : -124.508523
                      }
                    />
                  </GoogleMapReact>
                </div>
              </Grid>

              <Grid item xs={4}>
                <div
                  style={{ height: '180px', width: '190px' }}
                  className="text-muted"
                >
                  <div className={styles.text1}>
                  {this.state.location}
                  </div>
                  <br />
                  <div className={styles.text2}>
                  {this.state.op == 'Open Now' ? (
                    <div className={styles.dot1} />
                  ) : (
                    <div className={styles.dot2} />
                  )}
                  {space}
                  {this.state.op}
                  <br/><br/>
                  {this.state.status}
                  </div>

                </div>
              </Grid>
            </Grid>
          </Grid>

          <Grid item>
            <h6 className="pt-2 ">Members</h6>
            <Facepile
             maxDisplayablePersonas={5}
             overflowButtonProps={overflowButtonProps}
             overflowButtonType={OverflowButtonType.descriptive}
             getPersonaProps={getPersonaProps}
             personas={this.faces()}
            ></Facepile>
          </Grid>

          <Grid item>
          <h6 className="pt-2 ">Nearby Hotels</h6>
            <Grid spacing={1} container direction="row">

             <Grid item>
               <Image height={80} width={100} src={this.state.hotels!=undefined ?this.state.hotels[0].picture :hotelpic}></Image>
             </Grid>

             <Grid style={{maxHeight:80, height:80, width:240}} className="text-muted" item >
             {this.state.hotels!=undefined ?this.state.hotels[0].name :"Information unavailable"}<br/>
             <div className={styles.text}> {this.state.hotels!=undefined ?this.state.hotels[0].address :"Information unavailable"}<br/>

             {this.state.hotels!=undefined ? this.state.hotels[0].op == 'Open Now' ? (<div className={styles.dot1}/>) : (<div className={styles.dot2} />) :(<div className={styles.dot2} />)} {space}
             {this.state.hotels!=undefined ?this.state.hotels[0].op :"Information unavailable"}</div>  
             </Grid>

             <Grid item>
             {this.state.hotels!=undefined ?<Rating name="read-only" value={this.state.hotels[0].rating} readOnly/> :""}
             </Grid>

            </Grid>
          </Grid>

          <Grid item>
            <Grid spacing={1} container direction="row">

             <Grid item>
               <Image height={80} width={100} src={this.state.hotels!=undefined ?this.state.hotels[1].picture :hotelpic}></Image>
             </Grid>

             <Grid style={{height:80, width:240}} className="text-muted" item >
             {this.state.hotels!=undefined ?this.state.hotels[1].name :"Information unavailable"}<br/>
             <div className={styles.text}> {this.state.hotels!=undefined ?this.state.hotels[1].address :"Information unavailable"}<br/>

             {this.state.hotels!=undefined ? this.state.hotels[1].op == 'Open Now' ? (<div className={styles.dot1}/>) : (<div className={styles.dot2} />) :(<div className={styles.dot2}/>)}{space}
             {this.state.hotels!=undefined ?this.state.hotels[1].op :"Information unavailable"}</div>
             </Grid>

             <Grid item>
             {this.state.hotels!=undefined ?<Rating name="read-only" value={this.state.hotels[1].rating} readOnly/> :""}
             </Grid>

            </Grid>
          </Grid>

        </Grid>
        <br />
        <Divider orientation={'horizontal'}></Divider>
      </div>
    );
  }

  
  public faces() {
    let le = this.props.profileProperties.Department.length;
    let pile = [];
    for (let i = 0; i < le; i++) {
      let temp: any = this.props.profileProperties.Department[i];
      let img = `/_layouts/15/userphoto.aspx?size=S&accountname=${temp.WorkEmail}`;
      let name = temp.PreferredName;
      

      pile.push({
        imageUrl: img,
        personaName: name,
        hidePersonaDetails: false,
      });
    }
    return pile;
  }

  public async hotels(hotel){
    
    let temp=[]
    let op=undefined
    let name=undefined
    let address=undefined
    let picture=undefined
    let rating=undefined

    if (hotel!=null){   
        for(let i=0;i<=1;i++){    
          
          //Hotel Name
          if(hotel[i].name!=null || hotel[i].name!=undefined){name=hotel[i].name}
          else{name="Not Available"}
          
          //Hotel Open or Not
          if(hotel[i].opening_hours!=null || hotel[i].opening_hours!=undefined)
             {if(hotel[i].opening_hours.open_now==true){op="Open Now"}
             else{op="Closed Now"}
            }
          else{op="Not Available"}

          //Hotel Address
          if(hotel[i].vicinity!=undefined || hotel[i].vicinity!=null){address=hotel[i].vicinity}
          else{address="Not Available"}

          //Hotel Picture
          if(hotel[i].photos!=undefined || hotel[i].photos!=null){
            let temp = await Axios.get( `https://cors-anywhere.herokuapp.com/https://maps.googleapis.com/maps/api/place/photo?maxwidth=400&photoreference=${hotel[i].photos[0].photo_reference}&key=AIzaSyD7uFU7U6uumrFjm2wH7TZ8iFUxWwMcqpI`,
            {responseType: 'blob', method: 'GET', headers: {'Access-Control-Allow-Origin': '*','Content-Type': 'image/jpeg'}})
            picture= window.URL.createObjectURL(temp.data)
          }
          else{picture=hotelpic}

          //Hotel Rating
          if(hotel[i].rating!=undefined || hotel[i].rating!=null){rating=hotel[i].rating}

        temp.push({"name":name,"op":op,"address":address,"picture":picture,"rating":rating}) 
      }
      this.setState({hotels:temp})
      
    }
  }

  getPhoto = (refrence) => {
    Axios.get(
      `https://cors-anywhere.herokuapp.com/https://maps.googleapis.com/maps/api/place/photo?maxwidth=400&photoreference=${refrence}&key=AIzaSyD7uFU7U6uumrFjm2wH7TZ8iFUxWwMcqpI`,
      {responseType: 'blob', method: 'GET', headers: {'Access-Control-Allow-Origin': '*','Content-Type': 'image/jpeg'}}
    )
      .then((response) => {this.setState({pictureUrl: window.URL.createObjectURL(response.data)});})
      .catch((error) => {console.log(error);});
  };



  /**
   * Load SPFx component by id, SPComponentLoader is used to load the SPFx components
   * @param componentId - componentId, guid of the component library
   */

  private async _loadSPComponentById(componentId: string): Promise<any> {
    try {
      const component: any = await SPComponentLoader.loadComponentById(
        componentId
      );
      return component;
    } catch (error) {
      Promise.reject(error);
      Log.error(EXP_SOURCE, error, this.props.context.serviceScope);
    }
  }

  /**
   *
   *
   * @returns {React.ReactElement<IPersonaCardProps>}
   * @memberof PersonaCard
   */

  public render(): React.ReactElement<IPersonaCardProps> {
    return (
      <div className={styles.personaContainer}>
        {this.state.livePersonaCard
          ? this._LivePersonaCard()
          : this._PersonaCard()}
      </div>
    );
  }
}



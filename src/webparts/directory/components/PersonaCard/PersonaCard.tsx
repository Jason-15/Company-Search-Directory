import * as React from 'react';
import styles from './PersonaCard.module.scss';
import { IPersonaCardProps } from './IPersonaCardProps';
import { IPersonaCardState } from './IPersonaCardState';
import { PersonaPresence} from 'office-ui-fabric-react/lib/Persona';
import {Log} from '@microsoft/sp-core-library';
import {SPComponentLoader } from '@microsoft/sp-loader';
import 'bootstrap/dist/css/bootstrap.min.css';
import Divider from '@material-ui/core/Divider';
import Grid from '@material-ui/core/Grid';
import EmailOutlinedIcon from '@material-ui/icons/EmailOutlined';
import PhoneOutlinedIcon from '@material-ui/icons/PhoneOutlined';
import LocationOnOutlinedIcon from '@material-ui/icons/LocationOnOutlined';
import {Persona,PersonaSize} from 'office-ui-fabric-react';

const avatar1:string=require('../avataar/avataar1.png');
const avatar2:string=require('../avataar/avataar2.png');
const avatar3:string=require('../avataar/avataar3.png');
const avatar4:string=require('../avataar/avataar7.png');
const avatar5:string=require('../avataar/avataar8.png');
const EXP_SOURCE: string = 'SPFxDirectory';
const LIVE_PERSONA_COMPONENT_ID: string ='914330ee-2df2-4f6e-a858-30c23a812408';

export class PersonaCard extends React.Component<IPersonaCardProps,IPersonaCardState>
 {
  constructor(props: IPersonaCardProps) {
    super(props);

    this.state = {
      livePersonaCard: undefined,
      pictureUrl: null,
      title: undefined,
      location:undefined,
      phone:undefined,
      name:undefined,
      status,
      op:undefined,
      lat:undefined,
      lng:undefined,
      hotels:undefined
    };
  }

  /**
   *
   *
   * @memberof PersonaCard
   */

  public async componentDidMount() {
    const sharedLibrary = await this._loadSPComponentById(LIVE_PERSONA_COMPONENT_ID);
    const livePersonaCard: any = sharedLibrary.LivePersonaCard;
    this.setState({ livePersonaCard: livePersonaCard });
  }

  /**
   *
   *
   * @param {IPersonaCardProps} prevProps
   * @param {IPersonaCardState} prevState
   * @memberof PersonaCard
   */
  public componentDidUpdate(){}
   
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

  private _PersonaCard(): JSX.Element {
  let space=" "
    return (
      <div >
      <Grid  container  direction="row">

        <Grid item >
        <Persona
          className={styles.child}
          presence={this.presence()}
          imageUrl={this.avatar()}
          text={this.props.profileProperties.DisplayName}
          secondaryText={this.title()}
          size={PersonaSize.size100}
          tertiaryText={this.props.profileProperties.Department}/>
        </Grid>

          <Grid item alignContent={"flex-start"} justify={"flex-end"}>   
         <div className="mt-4 text-muted">
           <div className="pt-2" >
    <EmailOutlinedIcon  titleAccess={this.props.profileProperties.Email} fontSize={'small'}/>{space}{this.props.profileProperties.Email!=null?(this.props.profileProperties.Email):"Not Available"}
          </div>
          
          <div className="pt-2">
    <PhoneOutlinedIcon  titleAccess={this.props.profileProperties.WorkPhone} fontSize={'small'}/>{space}{this.props.profileProperties.MobilePhone!=null?(this.props.profileProperties.MobilePhone):(this.props.profileProperties.WorkPhone!=null?(this.props.profileProperties.WorkPhone):"Not Available")}
         </div>
          
          <div className="pt-2">
    <LocationOnOutlinedIcon  titleAccess={this.props.profileProperties.Location}  fontSize={'small'}/>{space}{this.props.profileProperties.Location!=null?(this.props.profileProperties.Location):"Not Available"}
          </div>
         </div>
         </Grid>

        </Grid>
        <Divider />
      </div>
    );
  }

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

  public avatar() {
    let number = Math.floor(Math.random() * 10 + 1);
    let def2:string=
      "data:image/png;base64,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"
    let def: string =
      'data:image/png;base64,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';
    if (this.props.profileProperties.PictureUrl == null || this.props.profileProperties.PictureUrl == def || this.props.profileProperties.PictureUrl == def2 ) {
      if (number <= 2) {
        return avatar1
      } else if (number <= 4) {
        return avatar2
      } else if (number <= 6) {
        return avatar3
      } else if (number <= 8) {
        return avatar4
      } else if (number <= 10) {
        return avatar5
      }
    } 
    else {
     return `/_layouts/15/userphoto.aspx?size=L&accountname=${this.props.profileProperties.Email}`;
    }
  }

  public title() {
    if (this.props.profileProperties.Title == null) {
      let number = Math.floor(Math.random() * 10 + 1);
      if (number <= 2) {
        return 'Striving for change' ;
      } else if (number <= 4) {
        return 'Pushing boundaries' ;
      } else if (number <= 6) {
       return 'Changing the world';
      } else if (number <= 8) {
        return 'Making a difference' ;
      } else if (number <= 10) {
        return 'Touching lives';
      }
    } else {
      return this.props.profileProperties.Title ;
    }
  }

public presence(){
  if(this.props.profileProperties.Department!=null){
    let temp=this.props.profileProperties.Department.substring(0).toUpperCase()
    if (this.props.profileProperties.Department.substring(0).toUpperCase()=="CONTRACTOR".substring(0)){console.log("First if" + this.props.profileProperties.DisplayName);return PersonaPresence.offline}
    else if(this.props.profileProperties.DisplayName.substring(0).search("Shared")!=-1){ console.log("Second if" + this.props.profileProperties.DisplayName);return PersonaPresence.offline}
    else{return PersonaPresence.none}
  }
  else if(this.props.profileProperties.DisplayName.substring(0).search("Shared")!=-1){ console.log("third if"+ this.props.profileProperties.DisplayName); return PersonaPresence.offline}
  else{return PersonaPresence.none}
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






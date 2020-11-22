import * as React from "react";
import styles from "./Directory.module.scss";
import { IDirectoryProps } from "./IDirectoryProps";
import { PersonaCard } from "./PersonaCard/PersonaCard2";
import { spservices2 } from "../../../SPServices/spservices2";
import { spservices } from "../../../SPServices/spservices";
import { IDirectoryState } from "./IDirectoryState";
import * as strings from "DirectoryWebPartStrings"
import {Spinner, SpinnerSize, MessageBar, MessageBarType, SearchBox, Icon, Label, PivotItem,} from "office-ui-fabric-react";

let OfficeEmp={}
let office = [];
const az: string[] = [
  'A',
  'B',
  'C',
  'D',
  'E',
  'F',
  'G',
  'H',
  'I',
  'J',
  'K',
  'L',
  'M',
  'N',
  'O',
  'P',
  'Q',
  'R',
  'S',
  'T',
  'U',
  'V',
  'W',
  'X',
  'Y',
  'Z',
];

export default class Directory extends React.Component<IDirectoryProps,IDirectoryState>
 {
  private _services2: spservices2 = null;
  private _services: spservices = null;
  
  constructor(props: IDirectoryProps) {
    super(props);

    this.state = {
      users: [],
      isLoading: true,
      errorMessage: "",
      hasError: false,
      indexSelectedKey: undefined,
      searchString: "Location",
      val:0,
      open1:false,
      open2:false
    };

    this._services2 = new spservices2(this.props.context);
    this._services=new spservices(this.props.context)
  
    // Register event handlers
    this._searchUsers = this._searchUsers.bind(this);
    this.getOffices(); 
     
  }

  /**
   *
   *
   * @memberof Directory
   */

  public async componentDidMount() {
    this.setBro();
  }

  /**
   * Gets image base64
   * @param pictureUrl
   * @returns
   */

  public  getOffices() {
     OfficeEmp={}
     office = [];
    for (let i = 0; i < 26; i++) {
      this._services.searchUsers(az[i], true).then((res) => {
        let results = res.PrimarySearchResults.length;

        for (let j = 0; j < results; j++) {
         let user: any = res.PrimarySearchResults[j];

         if(user.BaseOfficeLocation!=null){
           if (this.includes1(user.BaseOfficeLocation)) {
            office.push(user.BaseOfficeLocation);
            OfficeEmp[user.BaseOfficeLocation]=[user]
           }
        else{
          let temp=OfficeEmp[user.BaseOfficeLocation]
          temp.push(user)
          OfficeEmp[user.Department]=temp
        }
      }
      }
      this.setState({users:office,isLoading:false})
      }

        );
    }
    
  }

  public includes1(off) {
    for (let i = 0; i < office.length; i++) {
      if (office[i] == off || off == null) {
        return false;
      }
    }
    return true;
  }

  private async _searchUsers(searchText: string) {
    if(searchText==""){
      this.setState({users:office})
      return 
    }
    searchText = searchText.trim().length > 0 ? searchText : "A";
    this.setState({
      isLoading: true,
      indexSelectedKey: searchText.substring(0, 1).toLocaleUpperCase(),
      searchString: "Location"
    });
    try {
      const users = await this._services2.searchUsers(
        searchText,
        false
      );
      
      let arr:any=[]
      let userprops=[]

      if (users && users.PrimarySearchResults.length > 0){
        for (let index = 0; index < users.PrimarySearchResults.length; index++) {
          let user:any = users.PrimarySearchResults[index]  ;
          if (this.includes(arr,user.BaseOfficeLocation))
          {}
          else{ 
            arr.push(user.BaseOfficeLocation)
            userprops.push(user)
          }
        }
       }
       
      this.setState({
        users:
          users && users.PrimarySearchResults
            ? arr
            : null,
        isLoading: false,
        errorMessage: "",
        hasError: false
      });
    } catch (error) {
      this.setState({ errorMessage: error.message, hasError: true });
    }
  }
  public includes(arr,location){
  
    for (let index = 0; index < arr.length; index++) {
      if(arr[index]==location){
        return true
      }
  }
  return false
}

  /**
   *
   *
   * @param {IDirectoryProps} prevProps
   * @param {IDirectoryState} prevState
   * @memberof Directory
   */

  public async componentDidUpdate(
    prevProps: IDirectoryProps,
    prevState: IDirectoryState
  ) {
    if (
      this.props.title != prevProps.title ||
      this.props.searchFirstName != prevProps.searchFirstName
    ) {
    }
  }

  /**
   *
   *
   * @private
   * @param {string} sortField
   * @memberof Directory
   */

  /**
   *
   *
   * @private
   * @param {PivotItem} [item]
   * @param {React.MouseEvent<HTMLElement>} [ev]
   * @memberof Directory
   */

  private _selectedIndex(item?: PivotItem, ev?: React.MouseEvent<HTMLElement>) {
    this._searchUsers(item.props.itemKey);
  }

  /**
   *
   *
   * @returns {React.ReactElement<IDirectoryProps>}
   * @memberof Directory
   */
   

public setBro(){
  this.setState({users:office, isLoading:false})
}

  public render(): React.ReactElement<IDirectoryProps> {
    const color = this.props.context.microsoftTeams ? "white" : "";
    const diretoryGrid =
      this.state.users && this.state.users.length > 0
        ? this.state.users.map((user: any) => {
            return (
              <PersonaCard
                context={this.props.context}
                profileProperties={{
                  DisplayName: user,
                  Title:null,
                  PictureUrl: null,
                  Email: null,
                  Department:  OfficeEmp[user],
                  WorkPhone: null,
                  Location: null,
                }}
              />
            );
          })
        : [];

    return (
      <div className={styles.directory}>
        <div className={styles.searchBox}> 
         <br/>     
          <SearchBox
            placeholder="Search for Office Location"
            styles={{
              root: {
                minWidth: 180,
                maxWidth: 300,
                marginLeft: "auto",
                marginRight: "auto",
                marginBottom: 25
              }
            }}
            onSearch={this._searchUsers}
            onClear={()=>{console.log(office);this.setState({users:office})}}
            onChange={this._searchUsers}
          />

        </div>
        {!this.state.users || this.state.users.length == 0 ? (
          //if no users
          <div className={styles.noUsers}>
            <Icon
              iconName={"ProfileSearch"}
              style={{ fontSize: "54px", color: color }}
            />
            <Label>
              <span style={{ marginLeft: 5, fontSize: "26px", color: color }}>
                {strings.DirectoryMessage}
              </span>
            </Label>
          </div>
        ) : this.state.isLoading ? (
          <Spinner size={SpinnerSize.large} label={"searching ..."} />
        ) : this.state.hasError ? (
          <MessageBar messageBarType={MessageBarType.error}>
            {this.state.errorMessage}
          </MessageBar>
        ) : (
          //if users               
            <div>{diretoryGrid}</div>
        )}
      </div>
    );
  }

}


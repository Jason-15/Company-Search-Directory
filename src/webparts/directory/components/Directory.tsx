import * as React from 'react';
import styles from './Directory.module.scss';
import { IDirectoryProps } from './IDirectoryProps';
import { PersonaCard } from './PersonaCard/PersonaCard';
import { spservices } from '../../../SPServices/spservices';
import { IDirectoryState } from './IDirectoryState';
import * as strings from 'DirectoryWebPartStrings';
import List from '@material-ui/core/List';
import ListItem from '@material-ui/core/ListItem';
import ListItemText from '@material-ui/core/ListItemText';
import ExpandLess from '@material-ui/icons/ExpandLess';
import ExpandMore from '@material-ui/icons/ExpandMore';
import Collapse from '@material-ui/core/Collapse';
import Grid from '@material-ui/core/Grid';
import Slider from '@material-ui/core/Slider';
import { withStyles,} from '@material-ui/core/styles';
import DehazeTwoToneIcon from '@material-ui/icons/DehazeTwoTone';
import { Spinner, SpinnerSize, MessageBar, MessageBarType, SearchBox, Icon, Label,} from 'office-ui-fabric-react';


let DeptEmp:object={}
let OfficeEmp={}
let dept = [];
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

const marks = [
  {
    value: 0,
    label: 'A',
  },
  {
    value: 4,
    label: 'B',
  },
  {
    value: 8,
    label: 'C',
  },
  {
    value: 12,
    label: 'D',
  },
  {
    value: 16,
    label: 'E',
  },
  {
    value: 20,
    label: 'F',
  },
  {
    value: 24,
    label: 'G',
  },
  {
    value: 28,
    label: 'H',
  },
  {
    value: 32,
    label: 'I',
  },
  {
    value: 36,
    label: 'J',
  },
  {
    value: 40,
    label: 'K',
  },
  {
    value: 44,
    label: 'L',
  },
  {
    value: 48,
    label: 'M',
  },
  {
    value: 52,
    label: 'N',
  },
  {
    value: 56,
    label: 'O',
  },
  {
    value: 60,
    label: 'P',
  },
  {
    value: 64,
    label: 'Q',
  },
  {
    value: 68,
    label: 'R',
  },
  {
    value: 72,
    label: 'S',
  },
  {
    value: 76,
    label: 'T',
  },
  {
    value: 80,
    label: 'U',
  },
  {
    value: 84,
    label: 'V',
  },
  {
    value: 88,
    label: 'W',
  },
  {
    value: 92,
    label: 'X',
  },
  {
    value: 96,
    label: 'Y',
  },
  {
    value: 100,
    label: 'Z',
  },
];

export default class Directory extends React.Component< IDirectoryProps, IDirectoryState> 
{
  private _services: spservices = null;

  constructor(props: IDirectoryProps) {
    super(props);

    this.state = {
      users: [],
      isLoading: true,
      errorMessage: '',
      hasError: false,
      indexSelectedKey: 'A',
      searchString: 'LastName',
      val: 0,
      open1: false,
      open2: false,
    };

    this._services = new spservices(this.props.context);
    // Register event handlers
    this._searchUsers = this._searchUsers.bind(this);
    this.getDept();
  }

  /**
   *
   *
   * @memberof Directory
   */
  public async componentDidMount() {
    await this._searchUsers('A');
  }

  /**
   * Gets image base64
   * @param pictureUrl
   * @returns
   */
  private  getImageBase64(pictureUrl: string): Promise<string> {
    return new Promise((resolve, reject) => {
      let image = new Image();
      image.addEventListener('load', () => {
        let tempCanvas = document.createElement('canvas');
         (tempCanvas.width = image.width),
          (tempCanvas.height = image.height),
           (tempCanvas.className=styles.wrapper)
          tempCanvas.getContext('2d').drawImage(image, 0, 0);
         let base64Str;
        try {
          base64Str = tempCanvas.toDataURL('image/png');
        } catch (e) {
          return '';
        }
        resolve(base64Str);
      });
      image.src = pictureUrl;
    });
  }

  private async getDept() {
      DeptEmp={}
      OfficeEmp={}
      dept = [];
      office = [];
     for (let i = 0; i < 26; i++) {
      this._services.searchUsers(az[i], true).then(async res => {
      let results = res.PrimarySearchResults.length;

       for (let j = 0; j < results; j++) {
          let user: any = res.PrimarySearchResults[j];
          user = {...user,PictureURL: await this.getImageBase64(`/_layouts/15/userphoto.aspx?size=L&accountname=${user.WorkEmail}`)}

          if(user["PreferredName"]!=null && user["PreferredName"]!=undefined){
            if(user["PreferredName"].substring(0).search("Shared")!=-1 || user["PreferredName"].substring(0).search("Admin")!=-1)
            {continue}
          }
          if(user["FirstName"]!=null && user["FirstName"]!=undefined){
            if( user["FirstName"].substring(0).search("Account")!=-1 || user["FirstName"].substring(0).search("Admin")!=-1 ||  user["FirstName"].substring(0).search("Test")!=-1 || user["FirstName"].substring(0).search("Shared")!=-1){
              continue
            }
          }
          if(user["LastName"]!=null && user["LastName"]!=undefined){
            if(user["LastName"].substring(0).search("Account")!=-1 || user["LastName"].substring(0).search("Admin")!=-1 ||  user["LastName"].substring(0).search("Test")!=-1 || user["LastName"].substring(0).search("Shared")!=-1)
           {
            continue
           }
          }

        if(user.Department!=null){
          if (this.includes(user.Department)) {
            dept.push(user.Department);
            DeptEmp[user.Department]=[user]
          }
          else{
            let temp=DeptEmp[user.Department]
            temp.push(user)
            DeptEmp[user.Department]=temp
          }
        }

        if(user.BaseOfficeLocation!=null){
          if (this.includes1(user.BaseOfficeLocation)){
            office.push(user.BaseOfficeLocation);
            OfficeEmp[user.BaseOfficeLocation]=[user]
           }

           else{
           let temp=OfficeEmp[user.BaseOfficeLocation]
           temp.push(user)
           OfficeEmp[user.BaseOfficeLocation]=temp
           }
          }
        else if(user.OfficeNumber!=null){
          if (this.includes1(user.OfficeNumber)){
            office.push(user.OfficeNumber);
            OfficeEmp[user.OfficeNumber]=[user]
           }

           else{
           let temp=OfficeEmp[user.OfficeNumber]
           temp.push(user)
           OfficeEmp[user.OfficeNumber]=temp
           }

          }

        
        }
      }
      );
    }
  }

  public includes(dep) {
    for (let i = 0; i < dept.length; i++) {
      if (dept[i] == dep ) {
        return false;
      }
    }
    return true;
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
    let temp=[]
    searchText = searchText.trim().length > 0 ? searchText : 'A';
    this.setState({
      indexSelectedKey: searchText.substring(0, 1).toLocaleUpperCase(),
      searchString: 'FirstName',
      isLoading: true,
    });

    try {
      const users = await this._services.searchUsers(
        searchText,
        this.props.searchFirstName
      );
     for(let s=0; s<users.PrimarySearchResults.length;s++){

       if(users.PrimarySearchResults[s]["PreferredName"]!=null && users.PrimarySearchResults[s]["PreferredName"]!=undefined){
        if(users.PrimarySearchResults[s]["PreferredName"].substring(0).search("Shared")==-1 && users.PrimarySearchResults[s]["PreferredName"].substring(0).search("Admin")==-1 &&  users.PrimarySearchResults[s]["PreferredName"].substring(0).search("Test")==-1 &&  users.PrimarySearchResults[s]["PreferredName"].substring(0).search("Account")==-1){
          temp.push(users.PrimarySearchResults[s])
        }
       }

       else if(users.PrimarySearchResults[s]["FirstName"]!=null && users.PrimarySearchResults[s]["FirstName"]!=undefined){
         if(users.PrimarySearchResults[s]["FirstName"].substring(0).search("Account")==-1 && users.PrimarySearchResults[s]["FirstName"].substring(0).search("Admin")==-1 &&   users.PrimarySearchResults[s]["FirstName"].substring(0).search("Test")==-1 &&  users.PrimarySearchResults[s]["FirstName"].substring(0).search("Shared")==-1 ){
          temp.push(users.PrimarySearchResults[s])
         }
       }

       else if(users.PrimarySearchResults[s]["LastName"]!=null && users.PrimarySearchResults[s]["LastName"]!=undefined){
        if(
          users.PrimarySearchResults[s]["LastName"].substring(0).search("Account")==-1 &&  users.PrimarySearchResults[s]["LastName"].substring(0).search("Admin")==-1 &&   users.PrimarySearchResults[s]["LastName"].substring(0).search("Test")==-1 &&  users.PrimarySearchResults[s]["LastName"].substring(0).search("Shared")==-1
          ){
           temp.push(users.PrimarySearchResults[s])
          }
       }
    
     }

      if (users && temp.length > 0) {
        for (
          let index = 0;
          index < temp.length;
          index++
        ) {
          let user: any = temp[index];
          if (user.PictureURL) {
            user = {
              ...user,
              PictureURL: await this.getImageBase64(`/_layouts/15/userphoto.aspx?size=L&accountname=${user.WorkEmail}`),
            };
            temp[index] = user;
          }
        }
      }

      this.setState({
        users:
           temp.length!=0
            ? temp
            : null,
        isLoading: false,
        errorMessage: '',
        hasError: false,
      });
    } catch (error) {
      this.setState({ errorMessage: error.message, hasError: true });
    }
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
      await this._searchUsers('A');
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

  /**
   *
   *
   * @returns {React.ReactElement<IDirectoryProps>}
   * @memberof Directory
   */

  public render(): React.ReactElement<IDirectoryProps> {
    const iOSBoxShadow =
      '0 3px 1px rgba(0,0,0,0.1),0 4px 8px rgba(0,0,0,0.13),0 0 0 1px rgba(0,0,0,0.02)';

    const IOSSlider = withStyles({
      root: {
        color: '#3880ff',
        height: 2,
        padding: '15px 0',
      },
      thumb: {
        height: 28,
        width: 28,
        backgroundColor: '#fff',
        boxShadow: iOSBoxShadow,
        marginTop: -14,
        marginLeft: -14,
        '&:focus, &:hover, &$active': {
          boxShadow:
            '0 3px 1px rgba(0,0,0,0.1),0 4px 8px rgba(0,0,0,0.3),0 0 0 1px rgba(0,0,0,0.02)',
          // Reset on touch devices, it doesn't add specificity
          '@media (hover: none)': {
            boxShadow: iOSBoxShadow,
          },
        },
      },
      active: {},
      valueLabel: {
        left: 'calc(-50% + 12px)',
        top: -22,
        '& *': {
          background: 'transparent',
          color: '#000',
        },
      },
      track: {
        height: 2,
      },
      rail: {
        height: 2,
        opacity: 0.5,
        backgroundColor: '#bfbfbf',
      },
      mark: {
        backgroundColor: '#bfbfbf',
        height: 8,
        width: 1,
        marginTop: -3,
      },
      markActive: {
        opacity: 1,
        backgroundColor: 'currentColor',
      },
    })(Slider);

    const color = this.props.context.microsoftTeams ? 'white' : '';
    const diretoryGrid =
      this.state.users && this.state.users.length > 0
        ? this.state.users.map((user: any) => {
            return (
              <PersonaCard
                context={this.props.context}
                profileProperties={{
                  DisplayName: user.PreferredName,
                  Title: user.JobTitle,
                  PictureUrl: user.PictureURL,
                  Email: user.WorkEmail,
                  Department: user.Department,
                  WorkPhone: user.WorkPhone,
                  Location: user.OfficeNumber
                    ? user.OfficeNumber
                    : user.BaseOfficeLocation,
                }}
              />
            );
          })
        : [];
        
        const Filter= 
        <List className="pt-5" aria-label="Filter">
        <Grid container direction="row" spacing={2}>
         <Grid className="pb-2" item><DehazeTwoToneIcon fontSize="small"/></Grid>
         <Grid item><h4>Filter</h4></Grid>
        </Grid>

        {/*Department*/}
               <ListItem
                 button
                   onClick={() => {
                 if(this.state.open1==true){
                  this.setState({ open1: false});
                  }
                  else{
                    this.setState({ open1:true});
                  }
                    }}
                  >
                    <ListItemText primary="Department" />
                    {this.state.open1 ? <ExpandLess /> : <ExpandMore />}
                  </ListItem>
                  <Collapse in={this.state.open1} timeout="auto" unmountOnExit>
                    <List dense>
                      {dept.map((index: string) => {
                        return (
                          <ListItem divider button onClick={()=>{ this.setState({users:DeptEmp[index],
                            isLoading: false,
                            errorMessage: '',
                            hasError: false,})}}>
                            <ListItemText  className="text-muted" primary={index} />
                          </ListItem>
                        );
                      })}
                    </List>
                  </Collapse>

        {/*Office*/}
          <ListItem button
            onClick={() => {
              if(this.state.open2==true){
                this.setState({ open2: false});
                }
                else{
                  this.setState({ open2:true});
                }
              }}
              >
            <ListItemText primary="Offices" />
            {this.state.open2 ? <ExpandLess /> : <ExpandMore />}
          </ListItem>
          <Collapse in={this.state.open2}  unmountOnExit>
            <List dense>
              {office.map((index: string) => {
                return (
                  <ListItem divider button onClick={()=>{this.setState({users:OfficeEmp[index]})}}>
                    <ListItemText className="text-muted" primary={index} />
                  </ListItem>
                );
              })}
            </List>
          </Collapse>

      </List>
        

    return (
      <div className={styles.entire}>
        
         <Grid container direction="row" spacing={3} >

           <Grid alignContent="flex-start" item xs={3}>
            <SearchBox
             className="mt-3"
             placeholder={strings.SearchPlaceHolder}
             onSearch={this._searchUsers}  onClear={() => {this._searchUsers('A')}} onChange={this._searchUsers}/>
           </Grid>
            
            <Grid item xs={8}>
             <IOSSlider
              className="mt-2"
              onChangeCommitted={this.handleChange}
              step={4}
              aria-label="ios slider"
              defaultValue={this.state.val}
              marks={marks}
              valueLabelDisplay="off" />
            </Grid>

          </Grid>
        

        {!this.state.users || this.state.users.length == 0 ? (
          <div className={styles.noUsers}>
            <Icon
              iconName={'ProfileSearch'}
              style={{marginLeft:'30px', fontSize: '54px', color: color }}
            />
            <Label>
              <span style={{ marginLeft: 5, fontSize: '26px', color: color }}>
                {strings.DirectoryMessage}
              </span>
            </Label>
          </div>
        ) : this.state.isLoading ? (
          <Spinner size={SpinnerSize.large} label={'searching ...'} />
        ) : this.state.hasError ? (
          <MessageBar messageBarType={MessageBarType.error}>
            {this.state.errorMessage}
          </MessageBar>
        ) :
        
        (
            <Grid container justify="flex-start" direction={"row"} spacing={2}>
              <Grid className="mb-5" item justify="flex-start"  xs={3}>
                {Filter}
              </Grid>
              <Grid className="bt-5" item xs={9}>
                {diretoryGrid}
              </Grid>
            </Grid>
           
        )}
      </div>
    );
  }

  public handleChange = (event: any, value: number) => {
    this.setState({ val: value });
    if (value == 0) {
      this._searchUsers('A');
    } else if (value == 4) {
      this._searchUsers('B');
    } else if (value == 8) {
      this._searchUsers('C');
    } else if (value == 12) {
      this._searchUsers('D');
    } else if (value == 16) {
      this._searchUsers('E');
    } else if (value == 20) {
      this._searchUsers('F');
    } else if (value == 24) {
      this._searchUsers('G');
    } else if (value == 28) {
      this._searchUsers('H');
    } else if (value == 32) {
      this._searchUsers('I');
    } else if (value == 36) {
      this._searchUsers('J');
    } else if (value == 40) {
      this._searchUsers('K');
    } else if (value == 44) {
      this._searchUsers('L');
    } else if (value == 48) {
      this._searchUsers('M');
    } else if (value == 52) {
      this._searchUsers('N');
    } else if (value == 56) {
      this._searchUsers('O');
    } else if (value == 60) {
      this._searchUsers('P');
    } else if (value == 64) {
      this._searchUsers('Q');
    } else if (value == 68) {
      this._searchUsers('R');
    } else if (value == 72) {
      this._searchUsers('S');
    } else if (value == 76) {
      this._searchUsers('T');
    } else if (value == 80) {
      this._searchUsers('U');
    } else if (value == 84) {
      this._searchUsers('V');
    } else if (value == 88) {
      this._searchUsers('W');
    } else if (value == 92) {
      this._searchUsers('X');
    } else if (value == 96) {
      this._searchUsers('Y');
    } else if (value == 100) {
      this._searchUsers('Z');
    }
  };
}


{/*            <Pivot
              styles={{
                root: {
                  whiteSpace: "normal",
                  textAlign: "center"
                }
              }}
              linkFormat={PivotLinkFormat.links}
              selectedKey={this.state.indexSelectedKey}
              onLinkClick={this._selectedIndex}
              linkSize={PivotLinkSize.large}
            >
              {az.map((index: string) => {
                return (
                  <PivotItem headerText={index} itemKey={index} key={index} />
                );
              })}
            </Pivot>

                   {dept.map((index: string) => {
                return (
                  <ListItem button>
                   <ListItemText primary={index}/>
                  </ListItem>
                );
              })}
          */
}

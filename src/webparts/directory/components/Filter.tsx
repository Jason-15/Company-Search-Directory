import * as React from "react"

import List from '@material-ui/core/List';
import ListItem, { ListItemProps } from '@material-ui/core/ListItem';
import ListItemIcon from '@material-ui/core/ListItemIcon';
import ListItemText from '@material-ui/core/ListItemText';
import ExpandLess from '@material-ui/icons/ExpandLess';
import ExpandMore from '@material-ui/icons/ExpandMore';
import Collapse from '@material-ui/core/Collapse';
import {IFilterProps} from "./IFilterprops"

export default class Main extends React.Component< IFilterProps, {open1:boolean}>{
    constructor(props: any) {
        super(props);
        this.state = {
          open1:false,
        };
      }
    public render(){
        return(
            <List  aria-label="Filter">

            {/*Department*/}
          <ListItem button
            onClick={() => {
              console.log("sdsdd")
              this.setState({open1:true})
            }}>
            <ListItemText primary="Department" />
            {this.state.open1 ? <ExpandLess /> : <ExpandMore />}
          </ListItem>
          <Collapse in={true} timeout="auto" unmountOnExit>
            <List disablePadding>
            {this.props.dept.map((index: string) => {
                return (
                  <ListItem divider button>
                   <ListItemText primary={index}/>
                  </ListItem>
                );
              })}
            </List>
          </Collapse>


          {/*Office*/}
          <ListItem button
            onClick={() => {
              console.log("sdsdd")
              this.setState({open1:true})
            }}>
            <ListItemText primary="Offices" />
            {this.state.open1 ? <ExpandLess /> : <ExpandMore />}
          </ListItem>
          <Collapse in={false} timeout="auto" unmountOnExit>
            <List disablePadding>
            {this.props.office.map((index: string) => {
                return (
                  <ListItem divider button>
                   <ListItemText primary={index}/>
                  </ListItem>
                );
              })}     
            </List>
          </Collapse>
        </List>
        )
    }

}
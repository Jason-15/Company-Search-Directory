import * as React from "react"

import NavBar from "react-bootstrap/Navbar"
import Nav from "react-bootstrap/Nav"
import Button from "react-bootstrap/Button"
import "bootstrap/dist/css/bootstrap.min.css";

import { spservices } from "../../../SPServices/spservices";
import Directory from "./Directory"
import Location from "./Location"
import {IDirectoryProps} from "./IDirectoryProps"
import {IDirectoryState} from "./IDirectoryState"
import styles from "./Directory.module.scss";

export default class Main extends React.Component<
IDirectoryProps,
{index:number}
> {
  
  private _services: spservices = null;
  constructor(props: any) {
    super(props);
    this.state = {
      index: 1,
    };
  }

  public render(): React.ReactElement<IDirectoryProps> {
    return (
      <div>
        <NavBar>
        <div className={styles.divider}/>
        <NavBar.Brand>Company Directory</NavBar.Brand>

        <Nav className="mr-auto">
          <Nav.Link  onClick={() => {this.setState({ index: 1 });}}
            variant="outline-primary">People</Nav.Link>

          <Nav.Link  onClick={() => {this.setState({ index: 2 });}}
            variant="outline-primary">Offices</Nav.Link>

          </Nav>
        </NavBar>
        <br/>
        {this.state.index == 1 ? (
          <Directory
           title={this.props.title}
           displayMode={this.props.displayMode}
           context={this.props.context}
           searchFirstName={true}
           updateProperty={this.props.updateProperty}
          />
        ) : null}

        {this.state.index == 2 ? (
          <Location
           title={this.props.title}
           displayMode={this.props.displayMode}
           context={this.props.context}
           searchFirstName={false}
           updateProperty={this.props.updateProperty}
          />
        ) : null}
    
        </div>
    )
        } 
    }

    {/*

                 <Nav.Link  onClick={() => {this.setState({ index: 3 });}}
            variant="outline-primary">Teams</Nav.Link>

          <Nav.Link  onClick={() => {this.setState({ index: 4 });}}
            variant="outline-primary">Projects</Nav.Link>
            
              <Button 
          className={styles.headerbutton}
            onClick={() => {this.setState({ index: 1 });}}
            variant="outline-primary">People
          </Button>

          <Button 
            className={styles.headerbutton} 
            onClick={() => {this.setState({ index: 2 });}}
            variant="outline-primary">Location
          </Button>

          <Button
            className={styles.headerbutton}   
            onClick={() => {this.setState({ index: 3 });}}
            variant="outline-primary">Teams
          </Button>

          <Button
            className={styles.headerbutton}  
            onClick={() => {this.setState({ index: 4 });}}
            variant="outline-primary">Project
    </Button>*/}
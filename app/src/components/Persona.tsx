import * as React from 'react';
import {
  Persona,
  PersonaSize,
  PersonaPresence
} from 'office-ui-fabric-react/lib/Persona';
import { Checkbox } from 'office-ui-fabric-react/lib/Checkbox';
import { Label } from 'office-ui-fabric-react/lib/Label';

const examplePersona = {
  imageUrl: '',
  imageInitials: 'AL',
  primaryText: 'Annie Lindqvist',
  secondaryText: 'Software Engineer',
  tertiaryText: 'In a meeting',
  optionalText: 'Available at 4:00pm'
};

import * as pnp from 'sp-pnp-js';

export interface IUserProfileWebPartState {
 
  firstName?: string;
  lastname?: string;
  userProfileProperties?: any[];
  isFirstName?: boolean;
  isLastName?: boolean;
  email?: string;
  isWorkPhone?: boolean;
  isDepartment?: boolean;
  displayName?: string;
  pictureUrl?: string;
  workPhone?: string;
  department?: string;
  isPictureUrl?: boolean;
  title?: string;
  office?: string;
  isOffice?: boolean;
}

export interface IUserProfileProps {
 
  firstName?: string;
  lastname?: string;
  userProfileProperties?: any[];
  isFirstName?: boolean;
  isLastName?: boolean;
  email?: string;
  isWorkPhone?: boolean;
  isDepartment?: boolean;
  displayName?: string;
  pictureUrl?: string;
  workPhone?: string;
  department?: string;
  isPictureUrl?: boolean;
  title?: string;
  office?: string;
  isOffice?: boolean;
}


export class PersonaComponent extends React.Component<IUserProfileProps, IUserProfileWebPartState> {
  constructor(props: IUserProfileProps) {
      debugger;
    super(props);
    this.state = {
      firstName: "",
      lastname: "",
      userProfileProperties: [],
      isFirstName: false,
      isLastName: false,
      email: "",
      workPhone: "",
      department: "",
      pictureUrl: "",
      isPictureUrl: false,
      title: "",
      office: "",
      isOffice: false
    };
  }

  public render() {
    
    return (
      <div>
       <Persona
          { ...examplePersona }
          size={ PersonaSize.extraLarge }
          presence={ PersonaPresence.blocked }
        />
      </div>
    );
  }

  public componentDidMount(): void {
    let self = this;
    this._getProperties().then(function(response: any) {
            debugger;
            self.setState({ userProfileProperties: response.UserProfileProperties.results });
            self.setState({ email: response.Email });
            self.setState({ displayName: response.DisplayName });
            self.setState({ title: response.Title });
            
            for (let i: number = 0; i < self.state.userProfileProperties.length; i++) {
 
                if (self.state.isFirstName == false || self.state.isLastName == false || self.state.isDepartment == false || self.state.isWorkPhone == false || self.state.isPictureUrl == false || self.state.isOffice == false) {
        
                if (self.state.userProfileProperties[i].Key == "FirstName") {
                    self.setState({ isFirstName:true, firstName: self.state.userProfileProperties[i].Value });
                }
                if (self.state.userProfileProperties[i].Key == "LastName") {
                    self.setState({ isLastName: true, lastname: self.state.userProfileProperties[i].Value });
                }
                if (self.state.userProfileProperties[i].Key == "WorkPhone") {
                    self.setState({ isWorkPhone: true, workPhone: self.state.userProfileProperties[i].Value });
                }
                if (self.state.userProfileProperties[i].Key == "Department") {
                    self.setState({ isDepartment:true, department: self.state.userProfileProperties[i].Value });
                }
                if (self.state.userProfileProperties[i].Key == "Office") {
                    self.setState({ isOffice: true, office: self.state.userProfileProperties[i].Value });
                }
                if (self.state.userProfileProperties[i].Key == "PictureURL") {
                    self.setState({ isPictureUrl: true, pictureUrl: self.state.userProfileProperties[i].Value });
                }
        
            }
            }
            
        }).catch(function(err) {
            console.log("Error: " + err);
        });
    
  }

  private _getProperties(): Promise<any> {
      let accountName = 'i:0#.f|membership|'+this.props.email;
      return pnp.sp.profiles.getPropertiesFor(accountName);
 
  }
}
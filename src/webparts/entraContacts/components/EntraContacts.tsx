import * as React from 'react';
//import styles from './EntraContacts.module.scss';
import type { IEntraContactsProps } from './IEntraContactsProps';
import { TextField } from '@fluentui/react/lib/TextField';
import { Stack } from '@fluentui/react/lib/Stack';
import { Persona, PersonaSize } from '@fluentui/react/lib/Persona';


export interface IUser {
  displayName: string;
  userPrincipalName: string;
  mail: string;
  jobTitle: string;
}


export interface IGroup {
  id: string;
  displayName: string;
}


export interface IEntraContactsState {
  allUsers: IUser[];
  allGroups: IGroup[];
  filteredUsers: IUser[];
  filterGroup: string;
  searchTerm: string;
}


export default class EntraContacts extends React.Component<IEntraContactsProps, IEntraContactsState> {
  constructor(props: IEntraContactsProps) {
    super(props);
    this.state = {
      allUsers: [],
      allGroups: [],
      filteredUsers: [],
      filterGroup: '',
      searchTerm: ''
    };
  }
  

  public componentDidMount(): void {
      this._getUsers().catch(error => {
        console.error(error);
      });
      this._getGroups().catch(error => {
        console.error(error);
      });
  }


  private _getUsers = async (): Promise<void> => {
    try {
      // Use await to get the response
      const response: any = await this.props.graphClient
        .api('/users')
        .select('displayName,userPrincipalName,mail,jobTitle')
        .get();

      const users: IUser[] = response.value;
      this.setState({ allUsers: users, filteredUsers: users });

    } catch (error) {
      // The catch block handles any errors
      console.error(error);
    }
  }

  private _getGroups = async (): Promise<void> => {
    try {
      const response: any = await this.props.graphClient
        .api('/groups')
        .select('id,displayName')
        .get();
      const allGroups: IGroup[] = response.value;
      const searchGroups: IGroup[] = allGroups.filter(g => g.displayName.includes('grp_')); //Only groups with a prefix will be used
      this.setState({ allGroups: searchGroups });

        console.log("gotten groups! ", searchGroups)
    } catch (error) {
      console.error(error);
    }
  }


  private _getUsersByGroup = async (): Promise<void> => {
    try {
      const response: any = await this.props.graphClient
        .api(`/groups/${this.state.filterGroup}/members/microsoft.graph.user`)
        .select('displayName,userPrincipalName,mail,jobTitle')
        .get();

        
        const users: IUser[] = response.value;
        console.log(users);
        this.setState({ filteredUsers: users });
    } catch ( error ) {
      console.error(error);
    } 
  }


  private _onSearchChange = (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string): void => {
    const searchTerm = newValue || '';
    this.setState({ searchTerm });

    const filteredUsers = this.state.allUsers.filter(user =>
      user.displayName.toLowerCase().includes(searchTerm.toLowerCase())
    );
    this.setState({ filteredUsers });
  }


  private _onSelectedGroupChange = (event: React.ChangeEvent<HTMLSelectElement>): void => {
    const selectedGroup = event.target.value
    this.setState({ filterGroup: selectedGroup });
  }


  private _onFilterPressed = (event: React.MouseEvent<HTMLButtonElement, MouseEvent>): void => {
    //for now only group
    if (this.state.filterGroup === '') {
      void this._getUsers();
    } else {
      void this._getUsersByGroup();
    }
  }

  
  public render(): React.ReactElement<IEntraContactsProps> {
    return (
      <div>
        <select onChange={this._onSelectedGroupChange}>
          <option key='' value=''>All groups</option>
          {this.state.allGroups.map(group => (
            <option key={group.id} value={group.id}>{group.displayName.replace('grp_', '').replace('_', ' ')}</option>
          ))}
        </select>
        <p>{this.state.filterGroup}</p>
        <button onClick={this._onFilterPressed}>Filter</button>

        <Stack tokens={{ childrenGap: 10 }}>
          <TextField
            label="Filter by name:"
            value={this.state.searchTerm}
            onChange={this._onSearchChange}
            placeholder="Enter a name to filter..."
          />
          {this.state.filteredUsers.map(user => (
            <Persona
              key={user.mail}
              text={user.displayName}
              secondaryText={user.userPrincipalName}
              tertiaryText={user.mail}
              size={PersonaSize.size48}
            />
          ))}
        </Stack>
      </div>
    );
  }
}

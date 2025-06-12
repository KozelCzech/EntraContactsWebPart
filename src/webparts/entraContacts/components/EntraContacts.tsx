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


export interface IEntraContactsState {
  allUsers: IUser[];
  filteredUsers: IUser[];
  searchTerm: string;
}


export default class EntraContacts extends React.Component<IEntraContactsProps, IEntraContactsState> {
  constructor(props: IEntraContactsProps) {
    super(props);
    this.state = {
      allUsers: [],
      filteredUsers: [],
      searchTerm: ''
    };
  }
  

  public componentDidMount(): void {
      this._getUsers().catch(error => {
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


  private _onSearchChange = (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string): void => {
    const searchTerm = newValue || '';
    this.setState({ searchTerm });

    const filteredUsers = this.state.allUsers.filter(user =>
      user.displayName.toLowerCase().includes(searchTerm.toLowerCase())
    );
    this.setState({ filteredUsers });
  }

  
  public render(): React.ReactElement<IEntraContactsProps> {
    return (
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
    );
  }
}

import * as React from 'react';
//import styles from './EntraContacts.module.scss';
import type { IEntraContactsProps } from './IEntraContactsProps';
import style from './EntraContacts.module.scss';
import UserCard from './userCardComponents/userCard';

import { ResponseType } from '@microsoft/microsoft-graph-client';


export interface IUser {
  id: string;
  displayName: string;
  userPrincipalName: string;
  mail: string;
  jobTitle: string;
  photoUrl?: string;
}


export interface IGroup {
  id: string;
  displayName: string;
}


export interface IEntraContactsState {
  isLoading: boolean;
  allUsers: IUser[];
  allGroups: IGroup[];
  filteredUsers: IUser[];
  filterGroup: string;
  nameSearch: string;
  emailSearch: string;
}


export default class EntraContacts extends React.Component<IEntraContactsProps, IEntraContactsState> {
  constructor(props: IEntraContactsProps) {
    super(props);
    this.state = {
      isLoading: true,
      allUsers: [],
      allGroups: [],
      filteredUsers: [],
      filterGroup: '',
      nameSearch: '',
      emailSearch: ''
    };
  }
  

  public componentDidMount(): void {
    this.setState({ isLoading: true });

    this._getUsers().then(initialUsers => {
      this.setState({ allUsers: initialUsers, filteredUsers: initialUsers, isLoading: false });
    }) .catch(error => {
      console.error(error);
      this.setState({ isLoading: false });
    });
    this._getGroups().catch(error => {
      console.error(error);
      this.setState({ isLoading: false });
    });

  }

//#region Getters
  private _getUsers = async (): Promise<IUser[]> => {
    try {
      // Use await to get the response
      const response: {value: IUser[]} = await this.props.graphClient
        .api('/users')
        .select('id,displayName,userPrincipalName,mail,jobTitle')
        .get();

      return this._assignUserPhotos(response.value);
    } catch (error) {
      // The catch block handles any errors
      console.error(error);
      return [];
    }
  }

  private _getGroups = async (): Promise<void> => {
    try {
      const response: {value: IGroup[]} = await this.props.graphClient
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


  private _getUsersByGroup = async (): Promise<IUser[]> => {
    try {
      const response: {value: IUser[] } = await this.props.graphClient
        .api(`/groups/${this.state.filterGroup}/members/microsoft.graph.user`)
        .select('id,displayName,userPrincipalName,mail,jobTitle')
        .get();

        return this._assignUserPhotos(response.value);
    } catch ( error ) {
      console.error(error);
      return [];
    } 
  }


  //#region UserPhoto
  private _getUserPhoto = async (userId: string): Promise<string | null> => {
    try {
      const photoBlob: Blob = await this.props.graphClient
        .api(`users/${userId}/photo/$value`)
        .responseType(ResponseType.BLOB)
        .get();

        return URL.createObjectURL(photoBlob);
    } catch (error) {
      console.error(error);
      return null;
    }
  }


  private _assignUserPhotos = async (users: IUser[]): Promise<IUser[]> => {
    const usersWithPhotos: IUser[] = await Promise.all(users.map(async (user) => {
      let photoUrl = await this._getUserPhoto(user.id);
      
      if (!photoUrl) {
        const initial = this._getInitials(user.displayName);
        photoUrl = this._generateInitialsImage(initial, user.displayName); 
      }

      return { ...user, photoUrl: photoUrl || undefined};
    }));

    return usersWithPhotos;
  }


  private _getInitials = (displayName: string): string => {
    if (!displayName) return '?';
    return displayName.trim()[0].toUpperCase();
  }


  private _generateInitialsImage = (initial: string, displayName: string): string => {
    const canvas = document.createElement('canvas');
    const size = 64;
    canvas.width = size;
    canvas.height = size;

    const context = canvas.getContext('2d');
    if (!context) return '';

    //generate background color from users name
    let hash = 0;
    for (let i = 0; i < displayName.length; i++) {
      hash = displayName.charCodeAt(i) + ((hash << 5) - hash);
      hash |= 0;
    }
    const color = `hsl(${hash % 360}, 40%, 50%)`;

    //draw the colored circle
    context.fillStyle = color;
    context.beginPath();
    context.arc(size / 2, size / 2, size / 2, 0, 2 * Math.PI);
    context.fill();

    //Draw the text
    context.fillStyle = '#ffffff';
    context.font = `bold ${size / 2}px Arial`;
    context.textAlign = 'center';
    context.textBaseline = 'middle';
    context.fillText(initial, size / 2, size / 1.9);

    return canvas.toDataURL();
  }
  //#endregion
//#endregion


//#region UIHandlers
  private _onSearchChange = (event: React.ChangeEvent<HTMLInputElement>): void => {
    const searchName = event.target.name as keyof IEntraContactsState;
    const newSearchTerm = event.target.value || '';
    this.setState({ [searchName]: newSearchTerm } as unknown as Pick<IEntraContactsState, typeof searchName>);
  }


  private _onSelectedGroupChange = (event: React.ChangeEvent<HTMLSelectElement>): void => {
    const selectedGroup = event.target.value
    this.setState({ filterGroup: selectedGroup });
  }


  private _onFilterPressed = async (event: React.MouseEvent<HTMLButtonElement, MouseEvent>): Promise<void> => {
    
    this.setState({ isLoading: true });
    let baseUserList: IUser[] = [];

    if (this.state.filterGroup === '') {
      baseUserList = await this._getUsers();
    } else {
      baseUserList = await this._getUsersByGroup();
    }

    console.log("Base users: ", baseUserList);

    const finalFilteredUsers: IUser[] = baseUserList
      .filter(user => 
        this.state.nameSearch ? (user.displayName || '').toLowerCase().includes(this.state.nameSearch.toLowerCase()) : true
      )
      .filter(user => 
        this.state.emailSearch ? (user.userPrincipalName || '').toLowerCase().includes(this.state.emailSearch.toLowerCase()) : true
      );

    console.log("Final filtered users: ", finalFilteredUsers);

    this.setState({ allUsers: baseUserList, filteredUsers: finalFilteredUsers, isLoading: false }, () => {
        console.log("Filtered users: ", this.state.filteredUsers);
        this.setState({ isLoading: false });
    });
  }
//#endregion
  
  public render(): React.ReactElement<IEntraContactsProps> {
    return (
      <div>
        <div className={style.inputs}>
          <div>
            <label>Group:</label>
            <select onChange={this._onSelectedGroupChange}>
              <option key='' value=''>All groups</option>
              {this.state.allGroups.map(group => (
                <option key={group.id} value={group.id}>{group.displayName.replace('grp_', '').replace('_', ' ')}</option>
              ))}
            </select>
          </div>
          <div>
            <label>Name:</label>
            <input type='text' name='nameSearch' value={this.state.nameSearch} onChange={this._onSearchChange} placeholder='Enter a name to filter...' />
          </div>
          <div>
            <label>Email:</label>
            <input type='text' name='emailSearch' value={this.state.emailSearch} onChange={this._onSearchChange} placeholder='Enter a name to filter...' />
          </div>
        </div>
        <p>{this.state.nameSearch}</p>
        <p>{this.state.emailSearch}</p>
        <button name="displayNameFilter" onClick={this._onFilterPressed}>Filter</button>

        <div className={style.cardHolder}>
          { this.state.filteredUsers.length > 0 ? (
            this.state.filteredUsers.map(user => (
            <UserCard key={user.id} user={user} />
          ))
          ) : (
            this.state.isLoading ? (
              <p>Loading users...</p>
            ) : (
              <p>No users found...</p>
            )
          )
          }
        </div>
      </div>
    );
  }
}

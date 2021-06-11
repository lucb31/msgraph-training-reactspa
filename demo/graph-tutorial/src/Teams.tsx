import withAuthProvider, {AuthComponentProps} from "./AuthProvider";
import React from "react";
import {config} from "./Config";
import {addMembersToGroup, createGroup, getGroups, getUserDetails, getUsers} from "./GraphService";
import {Group, User} from "microsoft-graph";

interface TeamsState {
    groupsLoaded: boolean;
    user: User | undefined;
    users: User [] | undefined;
}
class Teams extends React.Component<AuthComponentProps, TeamsState> {
    constructor(props: any) {
        super(props);

        this.state = {
            groupsLoaded: false,
            user: undefined,
            users: undefined
        };
        this.handleCreateButton = this.handleCreateButton.bind(this);
    }
    async getRandomUser(users: User[]): Promise<User> {
        const index = Math.floor(Math.random() * users.length);
        const user = users[index];
        if (this.state.user) {
            if (await this.getStateUserId(this.state.user) === user.id) {
                console.log("Trying to add state user. Recursion!");
                return await this.getRandomUser(users);
            }
        }
        return user;
    }

    async getStateUserId(user: User) {
        return user.id?.split("@")[0];
    }

    async getStateUser(user: User, users: User[]): Promise<User> {
        const userId = await this.getStateUserId(user);
        const filteredUsers = users.filter(user => user.id === userId);
        return (filteredUsers.length > 0) ? filteredUsers[0] : user;
    }

    async handleCreateButton() {
        const stateUser = (this.state.user && this.state.users) ? await this.getStateUser(this.state.user, this.state.users) : undefined;
        if (stateUser && this.state.users) {
            // Generate group of random members
            let groupMembers: User[] = [];
            const groupSize = 5;
            for (let i = 0; i < groupSize - 1; i++) {
                let randomUser = await this.getRandomUser(this.state.users);
                while (groupMembers.filter(user => user.id === randomUser.id).length > 0) randomUser = await this.getRandomUser(this.state.users);
                groupMembers.push(randomUser);
            }

            // Add current user with full details to members
            groupMembers.push(stateUser);

            const newGroup: Group = {
                displayName: "MS Graph API Group",
                mailNickname: "msgraphapigroup",
                description: "This is a group description",
                visibility: "Private",
                groupTypes: ["Unified"],
                mailEnabled: true,
                securityEnabled:false
            };
            try {
                const accessToken = await this.props.getAccessToken(config.scopes);
                const responseCreate = await createGroup(accessToken, newGroup);
                const responseAdd = await addMembersToGroup(accessToken, responseCreate, groupMembers);
                console.log("Response add: " + JSON.stringify(responseAdd));
            } catch (err) {
                this.props.setError('ERROR', JSON.stringify(err));
            }
        }
    }

    async componentDidUpdate() {
        if (this.props.user && !this.state.groupsLoaded) {
            try {
                // Get the user's access token
                const accessToken = await this.props.getAccessToken(config.scopes);

                const userDetails = await getUserDetails(accessToken);
                const userGroups = await getGroups(accessToken);
                console.log(userGroups);
                const users = await getUsers(accessToken);
                this.setState({
                    groupsLoaded: true,
                    user: userDetails as User,
                    users: users as User[]
                });
            } catch (err) {
                this.props.setError('ERROR', JSON.stringify(err));
            }
        }
    }

    render() {
        return (
            <div>
                <h1>Hello Teams</h1>
                <button onClick={this.handleCreateButton}>Create Group</button>
            </div>
        );
    };
}

export default withAuthProvider(Teams);

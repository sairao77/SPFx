export interface MyTeam{
    displayName: string;
    description: string;
    ownerormember: string;
}

export interface MyTeamsDetails{
    myTeamCount: number;
    Details: MyTeam[];
}
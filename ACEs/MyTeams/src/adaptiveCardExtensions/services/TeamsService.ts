import { GraphFI } from "@pnp/graph"
import { getGraph, getSP } from "../myTeams/pnpConfig"
import { SPFI } from "@pnp/sp"
import { MyTeam } from "../models/TeamsModels"


export class TeamsService{
    _sp: SPFI
    _graph: GraphFI

    constructor(){
        this._sp = getSP()
        this._graph = getGraph()
    }

    public async getMyTeams(): Promise<MyTeam[]>{
        let TeamDetails: MyTeam[] = [];
        const teamdataarray = await this._graph.me.joinedTeams()
        const cuupn = (await this._graph.me()).userPrincipalName
        for(const teamdata of teamdataarray){
          await this.confirmOwnerOrMember(teamdata.id, cuupn!).then((response) => {
                const tempdet: MyTeam = {
                    displayName: teamdata.displayName!,
                    description: teamdata.description!,
                    ownerormember: response,
                }
                TeamDetails.push(tempdet);
            })
            
        }
        console.log(TeamDetails);
        return TeamDetails;

    }

    public async confirmOwnerOrMember(teamid: string | undefined, upn: string): Promise<string>{
       const data = await this._graph.groups.getById(teamid!).owners();
       for(const d of data){
        if(d.userPrincipalName === upn){
            return "Owner"
        }
       }
       return "Member"

    }
}
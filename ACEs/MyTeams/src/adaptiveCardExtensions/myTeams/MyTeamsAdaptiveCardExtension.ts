import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseAdaptiveCardExtension } from '@microsoft/sp-adaptive-card-extension-base';
import { CardView } from './cardView/CardView';
import { QuickView } from './quickView/QuickView';
import { MyTeamsPropertyPane } from './MyTeamsPropertyPane';
import { getGraph, getSP } from './pnpConfig';
import { TeamsService } from '../services/TeamsService';
import { MyTeamsDetails } from '../models/TeamsModels';

export interface IMyTeamsAdaptiveCardExtensionProps {
  title: string;
}

export interface IMyTeamsAdaptiveCardExtensionState {
  myTeamDetails: MyTeamsDetails;
}

const CARD_VIEW_REGISTRY_ID: string = 'MyTeams_CARD_VIEW';
export const QUICK_VIEW_REGISTRY_ID: string = 'MyTeams_QUICK_VIEW';
const myIconUrl = require('./assets/TeamsLogoInverse.svg');

export default class MyTeamsAdaptiveCardExtension extends BaseAdaptiveCardExtension<
  IMyTeamsAdaptiveCardExtensionProps,
  IMyTeamsAdaptiveCardExtensionState
> {
  private _deferredPropertyPane: MyTeamsPropertyPane;
  TS: TeamsService;
  

  public onInit(): Promise<void> {
    this.state = {
      myTeamDetails: {
        myTeamCount: 0,
        Details: []
      }
     };
    getSP(this.context);
    getGraph(this.context);
    this.cardNavigator.register(CARD_VIEW_REGISTRY_ID, () => new CardView());
    this.quickViewNavigator.register(QUICK_VIEW_REGISTRY_ID, () => new QuickView());
    this.TS = new TeamsService();
    this.initialLoad();
    return Promise.resolve();
  }

  protected loadPropertyPaneResources(): Promise<void> {
    return import(
      /* webpackChunkName: 'MyTeams-property-pane'*/
      './MyTeamsPropertyPane'
    )
      .then(
        (component) => {
          this._deferredPropertyPane = new component.MyTeamsPropertyPane();
        }
      );
  }

  protected renderCard(): string | undefined {
    return CARD_VIEW_REGISTRY_ID;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return this._deferredPropertyPane?.getPropertyPaneConfiguration();
  }

  protected get iconProperty(): string {
    return myIconUrl;
  }

  private async initialLoad(){
    await this.TS.getMyTeams().then(Response => { console.log(Response.length); this.setState({myTeamDetails: {myTeamCount: Response.length,Details:Response}}) })
  }
}

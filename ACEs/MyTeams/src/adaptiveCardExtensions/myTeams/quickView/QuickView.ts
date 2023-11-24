import { ISPFxAdaptiveCard, BaseAdaptiveCardView } from '@microsoft/sp-adaptive-card-extension-base';
import { IMyTeamsAdaptiveCardExtensionProps, IMyTeamsAdaptiveCardExtensionState } from '../MyTeamsAdaptiveCardExtension';
import { MyTeam } from '../../models/TeamsModels';

export interface IQuickViewData {
  Details: MyTeam[];
  TeamCount: string;
}

export class QuickView extends BaseAdaptiveCardView<
  IMyTeamsAdaptiveCardExtensionProps,
  IMyTeamsAdaptiveCardExtensionState,
  IQuickViewData
> {
  public get data(): IQuickViewData {
    return {
      Details: this.state.myTeamDetails.Details,
      TeamCount: this.state.myTeamDetails.myTeamCount.toString()
    };
  }

  public get template(): ISPFxAdaptiveCard {
    return require('./template/QuickViewTemplate.json');
  }
}
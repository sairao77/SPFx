import { AdaptiveCardExtensionContext } from "@microsoft/sp-adaptive-card-extension-base";

// import pnp, pnp logging system, and any other selective imports needed
import { spfi, SPFI, SPFx as spSPFx } from "@pnp/sp";
import { graphfi, GraphFI ,SPFx as graphSPFx} from "@pnp/graph";
import { LogLevel, PnPLogging } from "@pnp/logging";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/batching";
import "@pnp/graph/users";
import "@pnp/graph/teams";
import "@pnp/graph/groups";
import "@pnp/graph/members";

var _sp: SPFI;
var _graph: GraphFI;

export const getSP = (context?: AdaptiveCardExtensionContext): SPFI => {
  if (context != null) {
    //You must add the @pnp/logging package to include the PnPLogging behavior it is no longer a peer dependency
    // The LogLevel set's at what level a message will be written to the console
    _sp = spfi().using(spSPFx(context)).using(PnPLogging(LogLevel.Warning));
  }
  return _sp;
};

export const getGraph = (context?: AdaptiveCardExtensionContext): GraphFI => {
    if (context != null) {
      //You must add the @pnp/logging package to include the PnPLogging behavior it is no longer a peer dependency
      // The LogLevel set's at what level a message will be written to the console
      _graph = graphfi().using(graphSPFx(context)).using(PnPLogging(LogLevel.Warning));
    }
    return _graph;
  };
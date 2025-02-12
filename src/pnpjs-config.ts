import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/sites";
import "@pnp/sp/files";
import "@pnp/sp/fields";
import "@pnp/sp/batching";
import "@pnp/sp/site-users/web";
import "@pnp/sp/folders";
import "@pnp/sp/items/get-all";
import "@pnp/sp/site-users"
import { spfi, SPFI, SPFx } from "@pnp/sp";

import { graphfi, SPFx as graphSPFx } from "@pnp/graph";
import { GraphFI } from "@pnp/graph/fi";
import "@pnp/graph/teams";
import "@pnp/graph/planner";
import "@pnp/graph/users";
import "@pnp/graph/contacts";
import "@pnp/graph/messages";

let _sp: SPFI | null = null;
let _graph: GraphFI | null = null;

export const getSP = (context: any): SPFI => {
  if (!_sp && context) _sp = spfi().using(SPFx(context));

  return _sp;
};

export const getGraph = (context: any): GraphFI => {
  if (!_graph && context) _graph = graphfi().using(graphSPFx(context));

  return _graph;
};

export const getSPByPath = (path: string, context: any): SPFI => {
  return spfi(path).using(SPFx(context));
};


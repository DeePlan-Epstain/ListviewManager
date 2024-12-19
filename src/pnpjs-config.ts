import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/sites";
import "@pnp/sp/files";
import "@pnp/sp/fields";
import "@pnp/sp/site-users/web";
import "@pnp/sp/items/get-all";
import { spfi, SPFI, SPFx } from "@pnp/sp";

var _sp: any = null;

const getSP = (context?: any): SPFI => {
  _sp = spfi().using(SPFx(context));
  return _sp;
};

export default getSP;

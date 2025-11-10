import { spfi, SPFx, SPFI } from "@pnp/sp";
import { Caching } from "@pnp/queryable";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

let _sp: SPFI | undefined = undefined;

export const getSP = (context: WebPartContext): SPFI => {
  if (!_sp) {
    _sp = spfi().using(SPFx(context)).using(
      Caching({
        store: "session",               // cache i sessionStorage
      })
    );
  }
  return _sp;
};

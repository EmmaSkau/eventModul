import { spfi, SPFx } from "@pnp/sp";
import { Caching } from "@pnp/queryable";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

let _sp: any = null;

export const getSP = (context: any) => {
  if (!_sp) {
    _sp = spfi().using(SPFx(context)).using(
      Caching({
        store: "session",               // cache i sessionStorage
      })
    );
  }
  return _sp;
};

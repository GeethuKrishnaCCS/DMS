// import pnp and pnp logging system
import { spfi, SPFx } from "@pnp/sp";
import { PnPLogging } from "@pnp/logging";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/batching";
import "@pnp/sp/files";
import "@pnp/sp/folders";
import "@pnp/sp/site-users/web";
import "@pnp/sp/items/get-all";
var _sp;
export var getSP = function (context) {
    if (context !== null) {
        //You must add the @pnp/logging package to include the PnPLogging behavior it is no longer a peer dependency
        // The LogLevel set's at what level a message will be written to the console
        _sp = spfi().using(SPFx(context)).using(PnPLogging(2 /* LogLevel.Warning */));
    }
    return _sp;
};
//# sourceMappingURL=pnpjsConfig.js.map
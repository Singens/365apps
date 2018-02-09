import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer, PlaceholderContent, PlaceholderName
} from '@microsoft/sp-application-base';
import { Dialog } from '@microsoft/sp-dialog';

import * as strings from 'NotifierApplicationCustomizerStrings';

const LOG_SOURCE: string = 'NotifierApplicationCustomizer';

/** A Custom Action which can be run during execution of a Client Side Application */
export default class NotifierApplicationCustomizer
  extends BaseApplicationCustomizer<{}> {
  constructor() {
    super();
  }

  @override
  public onInit(): Promise<void> {
    console.log("NotifierApplicationCustomizer.onInit ...");
    //Pass values to the Notifier script locatited in the Notifier assets library
    window["NotifierSingensWebAbsoluteUrl"] = this.context.pageContext.web.absoluteUrl;
    window["NotifierSingensWebServerRelativeUrl"] = this.context.pageContext.web.serverRelativeUrl;
    window.console.log("[1] Notifier.Extension: try get wrapper");
    //try to get the wrapper element; it may not be rendered yet;
    var items = window.document.getElementsByClassName("ms-compositeHeader-addnCommands");
    if (items.length > 0) {
      window.console.log("[1] wwrapper found; load notifier.js ");
      let parentNode: any = items[0];
      var wapperSpan = document.createElement("span");
      //set it as first element
      wapperSpan.innerHTML = "";
      wapperSpan.id = "RibbonContainer-TabRowRight";
      parentNode.appendChild(wapperSpan);
      parentNode.insertBefore(wapperSpan, parentNode.firstChild);

      var token = new Date();
      var headID = document.getElementsByTagName('head')[0];
      var newScript = document.createElement('script');
      newScript.type = 'text/javascript';
      newScript.src = this.context.pageContext.web.absoluteUrl + '/NotifierAddinSingens/notifier.js' + '?ver=' + (token.getTime() * 1);
      headID.appendChild(newScript);
      window.console.log("[1] done");
    }
    else {
      console.log("[1] extension wrapper was not found; execute again after timeout");
      setTimeout(() => {
        window.console.log("[2] Notifier.Extension: try get wrapper");
        // this is the same code as above
        var items2 = window.document.getElementsByClassName("ms-compositeHeader-addnCommands");
        if (items2.length > 0) {
          window.console.log("[2] wwrapper found; load notifier.js ");
          //use global variable; the timeout won't have correct this.context.pageContext object
          var scriptUrl2 = window["NotifierSingensWebAbsoluteUrl"] + '/NotifierAddinSingens/notifier.js' + '?ver=' + (new Date().getTime() * 1);

          let parentNode2: any = items2[0];
          var wapperSpan2 = document.createElement("span");
          wapperSpan2.innerHTML = "";
          wapperSpan2.id = "RibbonContainer-TabRowRight";
          parentNode2.appendChild(wapperSpan2);
          parentNode2.insertBefore(wapperSpan2, parentNode2.firstChild);

          var headID2 = document.getElementsByTagName('head')[0];
          var newScript2 = document.createElement('script');
          newScript2.type = 'text/javascript';
          newScript2.src = scriptUrl2;
          headID2.appendChild(newScript2);
          window.console.log("[2] done");
        }
        else {
          console.log("[2] Notifier: Extension wrapper was not found; execute again after timeout");
        }
      }, 2000);//2 secunds timeout
    }
    return Promise.resolve();
  }
}



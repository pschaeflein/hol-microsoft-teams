import {TeamsTheme} from './theme';

/**
 * Implementation of Module 2 configuration page
 */
export class module2Configure {
    tenantId?: string;
    
    constructor() {
      microsoftTeams.initialize();
    
      microsoftTeams.getContext((context: microsoftTeams.Context) => {
        TeamsTheme.fix(context);
        this.tenantId = context.tid;
        let val = <HTMLInputElement>document.getElementById("graph");
        if (context.entityId) {
          val.value = context.entityId;
        }
        this.setValidityState(val.value !== "");
      });
    
      microsoftTeams.settings.registerOnSaveHandler((saveEvent: microsoftTeams.settings.SaveEvent) => {
        let val = <HTMLInputElement>document.getElementById("graph");
    
        // Calculate host dynamically to enable local debugging
        let host = "https://" + window.location.host;
        microsoftTeams.settings.setSettings({
          contentUrl: host + "/module2Tab.html",
          suggestedDisplayName: 'Module 2',
          removeUrl: host + "/module2Remove.html",
          entityId: val.value
        });
    
        saveEvent.notifySuccess();
    
      });
    }
    public setValidityState(val: boolean) {
        microsoftTeams.settings.setValidityState(val);
    }
    
    public getAdminConsent() {
        microsoftTeams.authentication.authenticate({
          url: "/adminconsent.html?tenantId=" + this.tenantId,
          width: 800,
          height: 800,
          successCallback: () => { },
          failureCallback: (err) => { }
        });
      }    
}



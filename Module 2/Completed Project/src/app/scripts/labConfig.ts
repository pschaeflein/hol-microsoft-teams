import {TeamsTheme} from './theme';

/**
 * Implementation of Lab configuration page
 */
export class labConfigure {
  tenantId?: string;

  constructor() {
    microsoftTeams.initialize();

    microsoftTeams.getContext((context: microsoftTeams.Context) => {
      TeamsTheme.fix(context);
      this.tenantId = context.tid;

      // hack: the state should be retrieved from storage, not from Teams
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
        contentUrl: host + "/labTab.html",
        suggestedDisplayName: 'Lab Tab',
        removeUrl: host + "/labRemove.html",
        // hack: the state should be stored in external storage, not in Teams
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
      height: 600,
      successCallback: () => { },
      failureCallback: (err) => { }
    });
  }
}
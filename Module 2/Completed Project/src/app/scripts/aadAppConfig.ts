export class AADAppConfig {
  static clientID: string = "enter-your-app-id";
  static graphScopes: string[] = [
    "https://graph.microsoft.com/user.read",
    "https://graph.microsoft.com/group.read.all"
  ];
}

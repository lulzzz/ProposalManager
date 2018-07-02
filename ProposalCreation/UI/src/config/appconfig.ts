export class AppConfig {
    private static ApplicationId: string = "dd3e88ba-f7e7-423a-bfb6-e63f44286439";
    static get applicationId(): string { return AppConfig.ApplicationId; }
    static get accessTokenKey(): string { return "webapiAccessToken"; }
    static get title(): string { return "Commercial Lending"; }
}
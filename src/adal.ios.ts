/// <reference path="./platforms/ios/typings/adal-library.ios.d.ts" />

declare var interop: any;
declare var NSURL: any;

export class AdalContext {

  private authError: any;
  private authResult: ADAuthenticationResult;
  private authority: string;
  private clientId: string;
  private context: ADAuthenticationContext;
  private redirectUri: string;
  private resourceId: string;
  private userId;

  // Authority is in the form of https://login.microsoftonline.com/yourtenant.onmicrosoft.com
  constructor(authority: string, clientId: string, resourceId: string, redirectUri: string) {
    this.authError = new interop.Reference();
    this.authority = authority;
    this.clientId = clientId;
    this.resourceId = resourceId;
    this.redirectUri = redirectUri;
    ADAuthenticationSettings.sharedInstance().setDefaultKeychainGroup(null);
    this.context = ADAuthenticationContext.authenticationContextWithAuthorityError(this.authority, this.authError);
  }

  public login(fresh: boolean): Promise<string> {
    this.authError = new interop.Reference();
    var promptBehaviour = (fresh ? ADPromptBehavior.D_PROMPT_ALWAYS : ADPromptBehavior.D_PROMPT_AUTO);
    return new Promise<string>((resolve, reject) => {
      this.context.acquireTokenWithResourceClientIdRedirectUriPromptBehaviorUserIdentifierExtraQueryParametersCompletionBlock(
        this.resourceId,
        this.clientId,
        NSURL.URLWithString(this.redirectUri),
        promptBehaviour,
        this.userId,
        '',
        (result) => {
          this.authResult = result;
          if (result.error) {
            reject(result.error);
          } else {
            console.log('New token will expire on ' + result.tokenCacheItem.expiresOn);
            this.userId = result.tokenCacheItem.userInformation.userObjectId;
            resolve(result.accessToken);
          }
        });
    });
  }

  public getToken(): Promise<string> {
    return new Promise<string>((resolve, reject) => {
      this.context.acquireTokenSilentWithResourceClientIdRedirectUriCompletionBlock(
        this.resourceId,
        this.clientId,
        NSURL.URLWithString(this.redirectUri),
        (result) => {
          this.authResult = result;
          if (result.error) {
            // Failed in retrieving Access Token through refresh token. Now will prompt for log in
            console.log('Failed to retrieve token silently: ' + result.error);
            console.log('Now prompting for fresh login...');
            this.context.acquireTokenWithResourceClientIdRedirectUriPromptBehaviorUserIdentifierExtraQueryParametersCompletionBlock(
              this.resourceId,
              this.clientId,
              NSURL.URLWithString(this.redirectUri),
              ADPromptBehavior.D_PROMPT_ALWAYS,
              this.userId,
              '',
              (result) => {
                this.authResult = result;
                if (result.error) {
                  reject(result.error);
                } else {
                  this.userId = result.tokenCacheItem.userInformation.userObjectId;
                  console.log('New token will expire on ' + result.tokenCacheItem.expiresOn);
                  resolve(result.accessToken);
                }
              })
          } else {
            console.log('Suceeded in getting token silently.');
            console.log('New token will expire on ' + result.tokenCacheItem.expiresOn);
            resolve(result.accessToken);
          }
          
        }
      );
    });
  }
}
export interface ILoginInfo
{
    token?: string;
    succeeded: boolean;
    errorMessage?: string;
}

export class Login 
{
    static Signin():ILoginInfo
    {
        let loginInfo: ILoginInfo;

        loginInfo = {
            succeeded: false
        };

        if (Office.context.requirements.isSetSupported('IdentityAPI', 1.1)) 
        { 
            Office.context.auth.getAccessTokenAsync(
            result =>
            {
                if (result.status.toString() === "succeeded") {
                    // Use this token to call Web API
                    loginInfo.token = result.value;
                    loginInfo.succeeded = true;
                }
                else
                {
                    //TODO: add exception handling/redirection
                    loginInfo.errorMessage = `${result.error.name}<br/>${result.error.code}<br/>${result.error.message}<br/>`;
                }

                return loginInfo;
            });
        }

        loginInfo.errorMessage = "IdentityAPI not supported";

        return loginInfo;
    }
}
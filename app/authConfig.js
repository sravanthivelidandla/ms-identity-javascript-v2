// Config object to be passed to Msal on creation.
// For a full list of msal.js configuration parameters, 
// visit https://azuread.github.io/microsoft-authentication-library-for-js/docs/msal/modules/_authenticationparameters_.html
const msalConfig = {
    auth: {
        //clientId: "3ff8e6ba-7dc3-4e9e-ba40-ee12b60d6d48",
        authority: "https://login.windows-ppe.net/common",
        //directoryId: "ea8a4392-515e-481f-879e-6571ff2a8a36",
        //redirectUri: "https://localhost:5000/auth/callback",
        clientId:"3fba556e-5d4a-48e3-8e1a-fd57c12cb82e"
    },
    cache: {
        cacheLocation: "localStorage", // This configures where your cache will be stored
        storeAuthStateInCookie: false, // Set this to "true" if you are having issues on IE11 or Edge
    },
    system: {
        loggerOptions: {
            loggerCallback: (level, message, containsPii) => {
                if (containsPii) {	
                    return;	
                }	
                switch (level) {	
                    case msal.LogLevel.Error:	
                        console.error(message);	
                        return;	
                    case msal.LogLevel.Info:	
                        console.info(message);	
                        return;	
                    case msal.LogLevel.Verbose:	
                        console.debug(message);	
                        return;	
                    case msal.LogLevel.Warning:	
                        console.warn(message);	
                        return;	
                }
            }
        }
    }
};

// Add here the scopes that you would like the user to consent during sign-in
const loginRequest = {
    scopes: ["User.ReadWrite"]
};

// Add here the scopes to request when obtaining an access token for MS Graph API
const tokenRequest = {
    scopes: ["User.ReadWrite"],
    forceRefresh: false // Set this to "true" to skip a cached token and go to the server to get a new token
};

const graphRedirectRequest = {
    scopes: ['User.ReadWrite'], //graph request to fetch the user information   
  };

  // Add here scopes for access token to be used at MS Graph API endpoints.
  const shellRedirectRequest = {
    scopes: ['https://webshell.suite.office.com/default'] //fetch the shell token
    
  };

  const outlookRedirectRequest = {
    scopes:['https://outlook.office.com/Tasks.ReadWrite']
    
  };

  const silentOutlookRequest = {
    scopes: ['openid', 'profile', 'https://outlook.office.com/Tasks.ReadWrite']
    
  };

  const silentGraphRequest = {
    scopes: ['openid', 'profile', 'User.ReadWrite'],
    account: this.account,
    forceRefresh: false
  };

  const silentShellRequest = {
    scopes: ['openid', 'profile', 'https://webshell.suite.office.com/default'],
    account: this.account,
    forceRefresh: false
  };

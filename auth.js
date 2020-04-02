
import {UserAgentApplication} from "msal";

// AUTH

const attemptSilentSignIn = () => {
    return new Promise((resolve, reject) => {
        if (msalApp.getAccount()) {
            msalApp.acquireTokenSilent({scopes}).then((response) => {
                if (response && response.accessToken) {
                    resolve(true);
                } else {
                    resolve(false);
                }
            }, () => {
                resolve(false);
            })
        } else {
            resolve(false);
        }
    });
}


const signIn = () => {
    msalApp.loginRedirect({scopes});
}

const handleSignedIn = () => {
    microsoftTeams.initialize();
    microsoftTeams.authentication.notifySuccess();
}

const handleSignedOut = (error) => {
    microsoftTeams.initialize();
    microsoftTeams.authentication.notifyFailure(error);
}

const handleErrorReceived = (authError, accountState) => {
    console.log(authError, accountState);
    handleSignedOut({authError});
}

const handleTokenReceived = (response) => {
    console.log(response);
    handleSignedIn();
}





// STATE

const msalConfig = {
    auth: {
        clientId: "a974dfa0-9f57-49b9-95db-90f04ce2111a",
    },
    cache: {
        cacheLocation: "localStorage"
    }
};
const scopes = ['user.read'];

const msalApp = new UserAgentApplication(msalConfig);
msalApp.handleRedirectCallback((response) => handleTokenReceived(response), (error, state) => handleErrorReceived(error, state));






// MAIN


attemptSilentSignIn().then(success => {
    if (success){
        handleSignedIn();
    } else {
        signIn();
    }
});
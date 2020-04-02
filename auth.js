
import {UserAgentApplication} from "msal";
import {scopes, msalConfig} from './env.js';



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





// MAIN

const msalApp = new UserAgentApplication(msalConfig);
msalApp.handleRedirectCallback((response) => handleTokenReceived(response), (error, state) => handleErrorReceived(error, state));

attemptSilentSignIn().then(success => {
    if (success){
        handleSignedIn();
    } else {
        signIn();
    }
});
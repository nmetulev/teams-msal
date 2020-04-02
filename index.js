
import {UserAgentApplication} from "msal";

// STATE

let app = document.querySelector('.app');

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






// UI

const renderSignIn = () => {
    app.innerHTML = `<button>Sign In</button>`;
    app.querySelector('button').addEventListener('click', () => signInClicked());
}

const renderUser = (user) => {
    app.innerHTML = `<div>
        Hello ${user.displayName}
    </div>`
}


const signInClicked = () => {
    signIn();
}







// AUTH

const attemptSilentSignIn = () => {
    if (msalApp.getAccount()) {
        msalApp.acquireTokenSilent({scopes}).then((response) => {
            if (response && response.accessToken) {
                handleSignedIn(response.accessToken);
            } else {
                handleSignedOut();
            }
        }, () => {
            handleSignedOut();
        })
    } else {
        handleSignedOut();
    }

}

const signIn = () => {

    microsoftTeams.initialize();
    microsoftTeams.authentication.authenticate({
        url: window.location.origin + "/auth.html",
        successCallback: () => attemptSilentSignIn(),
        failureCallback: (error) => handleSignedOut(error)
    })

}

const handleSignedIn = (accessToken) => {
    let url = "https://graph.microsoft.com/v1.0/me";
    fetch(url, {
        headers: new Headers({
            'Authorization': accessToken
        })
    }).then((response) => {
        if (response && response.ok) {
            return response.json();
        }
    }).then((user) => {
        renderUser(user);
    }, error => {

    })
}

const handleSignedOut = (error) => {
    console.log(error)
    renderSignIn();
}

const handleErrorReceived = (authError, accountState) => {
    console.log(authError, accountState);
    handleSignedOut();
}

const handleTokenReceived = (response) => {
    console.log(response);
    attemptSilentSignIn();
}




// MAIN

attemptSilentSignIn();
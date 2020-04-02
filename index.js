
import {UserAgentApplication} from "msal";
import {scopes, msalConfig} from './env.js';

// UI

const renderSignedOutView = () => {
    app.innerHTML = `<button>Sign In</button>`;
    app.querySelector('button').addEventListener('click', () => signInClicked());
}

const renderUser = (user) => {
    app.innerHTML = `<div>
        Hello ${user.displayName}
    </div>`
}

const renderLoading = () => {
    app.innerHTML = `<div>Loading</div>`;
}

const renderError = (error) => {
    const errorDiv = document.createElement('div');
    errorDiv.innerText = JSON.stringify(error);
    app.appendChild(errorDiv);
}

const signInClicked = () => {
    signIn();
}






// AUTH

const attemptSilentSignIn = () => {
    renderLoading();
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
    renderLoading();
    microsoftTeams.initialize(() => {
        microsoftTeams.authentication.authenticate({
            url: window.location.origin + "/auth.html",
            successCallback: () => attemptSilentSignIn(),
            failureCallback: (error) => renderError(error)
        })
    });
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

const handleSignedOut = () => {
    renderSignedOutView();
}






// MAIN
let app = document.querySelector('.app');
const msalApp = new UserAgentApplication(msalConfig);

attemptSilentSignIn();
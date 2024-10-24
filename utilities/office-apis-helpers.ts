// office-apis-helpers.ts

import { AppState } from '../src/components/App';
import { AxiosResponse } from 'axios';

/*
     Interacting with the Office document
*/
export const writeFileNamesToEmail = async (
    result: AxiosResponse,
    displayError: (x: string) => void
  ) => {
    try {
      const names = result.data.value.slice(0, 3).map((item) => item.name);
      const message = names.join('\n');
      displayError(message);
    } catch (error) {
      displayError(error.toString());
    }
  };

/*
    Managing the dialogs.
*/

// Define processDialogEvent first
const processDialogEvent = (
    arg: { error: number; type: string },
    setState: (x: AppState) => void,
    displayError: (x: string) => void
) => {

    switch (arg.error) {
        case 12002:
            displayError('The dialog box has been directed to a page that it cannot find or load, or the URL syntax is invalid.');
            break;
        case 12003:
            displayError('The dialog box has been directed to a URL with the HTTP protocol. HTTPS is required.');
            break;
        case 12006:
            // 12006 means that the user closed the dialog instead of waiting for it to close.
            // It is not known if the user completed the login or logout, so assume the user is
            // logged out and revert to the app's starting state. It does no harm for a user to
            // press the login button again even if the user is logged in.
            setState({
                authStatus: 'notLoggedIn',
                currentFrame: 'default'
            });
            break;
        default:
            displayError('Unknown error in dialog box.');
            break;
    }
};

let loginDialog: Office.Dialog;
const dialogLoginUrl: string = location.protocol + '//' + location.hostname + (location.port ? ':' + location.port : '') + '/login/login.html';

export const signInO365 = (
    setState: (x: AppState) => void,
    setToken: (x: string) => void,
    setUserName: (x: string) => void,
    displayError: (x: string) => void
) => {

    setState({ authStatus: 'loginInProcess', currentFrame: 'default' });

    if (Office.context && Office.context.ui) {
        Office.context.ui.displayDialogAsync(
            dialogLoginUrl,
            { height: 40, width: 30 },
            (result) => {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    displayError(`${result.error.code} ${result.error.message}`);
                } else {
                    loginDialog = result.value;
                    loginDialog.addEventHandler(Office.EventType.DialogMessageReceived, processLoginMessage);
                    loginDialog.addEventHandler(Office.EventType.DialogEventReceived, processLoginDialogEvent);
                }
            }
        );
    } else {
        displayError("Office context is not available.");
    }

    const processLoginMessage = (arg: { message: string; origin: string }) => {
        // Confirm origin is correct.
        if (arg.origin !== window.location.origin) {
            throw new Error("Incorrect origin passed to processLoginMessage.");
        }

        let messageFromDialog = JSON.parse(arg.message);
        if (messageFromDialog.status === 'success') {

            // We now have a valid access token.
            loginDialog.close();
            setToken(messageFromDialog.token);
            setUserName(messageFromDialog.userName);
            setState({
                authStatus: 'loggedIn',
                currentFrame: 'Frame1'
            });
        } else {
            // Something went wrong with authentication or the authorization of the web application.
            loginDialog.close();
            displayError(messageFromDialog.result);
        }
    };

    const processLoginDialogEvent = (arg) => {
        processDialogEvent(arg, setState, displayError);
    };
};

let logoutDialog: Office.Dialog;
const dialogLogoutUrl: string = location.protocol + '//' + location.hostname + (location.port ? ':' + location.port : '') + '/logout/logout.html';

// From https://stackoverflow.com/questions/37764665/how-to-implement-sleep-function-in-typescript
function delay(milliSeconds: number) {
    return new Promise(resolve => setTimeout(resolve, milliSeconds));
}

export const logoutFromO365 = async (
    setState: (x: AppState) => void,
    setUserName: (x: string) => void,
    userName: string,
    displayError: (x: string) => void
) => {

    if (Office.context && Office.context.ui) {
        Office.context.ui.displayDialogAsync(dialogLogoutUrl,
            { height: 40, width: 30 },
            async (result) => {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    displayError(`${result.error.code} ${result.error.message}`);
                } else {
                    logoutDialog = result.value;
                    logoutDialog.addEventHandler(Office.EventType.DialogMessageReceived, processLogoutMessage);
                    logoutDialog.addEventHandler(Office.EventType.DialogEventReceived, processLogoutDialogEvent);
                    await delay(5000); // Wait for dialog to initialize and register handler for messaging.
                    logoutDialog.messageChild(JSON.stringify({ "userName": userName }));
                }
            }
        );
    } else {
        displayError("Office context is not available.");
    }

    const processLogoutMessage = () => {
        logoutDialog.close();
        setState({
            authStatus: 'notLoggedIn',
            currentFrame: 'default'
        });
        setUserName('');
    };

    const processLogoutDialogEvent = (arg) => {
        processDialogEvent(arg, setState, displayError);
    };
};

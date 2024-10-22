import * as React from "react";
import { Spinner, SpinnerType } from "office-ui-fabric-react";
import Header from "./Header";
import { HeroListItem } from "./HeroList";
import Progress from "./Progress";
import StartPageBody from "./StartPageBody";
import GetDataPageBody from "./GetDataPageBody";
import SuccessPageBody from "./SuccessPageBody";
import OfficeAddinMessageBar from "./OfficeAddinMessageBar";
import { getGraphData, createMailFolder } from "../../utilities/microsoft-graph-helpers";
import {
  writeFileNamesToEmail,
  logoutFromO365,
  signInO365,
} from "../../utilities/office-apis-helpers";
import Frame1 from "./Frame1";
import Frame2 from "./Frame2";
import Frame3 from "./Frame3";
export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

export interface AppState {
  authStatus?: string;
  fileFetch?: string;
  headerMessage?: string;
  errorMessage?: string;
  currentFrame: string;
  requestInput?: string; // Added this line
}

export default class App extends React.Component<AppProps, AppState> {
  constructor(props, context) {
    super(props, context);
    this.state = {
      authStatus: "notLoggedIn",
      fileFetch: "notFetched",
      headerMessage: "Welcome",
      errorMessage: "",
      currentFrame: "default",
    };

    // Bind methods to the class instance
    this.boundSetState = this.setState.bind(this);
    this.setToken = this.setToken.bind(this);
    this.setUserName = this.setUserName.bind(this);
    this.displayError = this.displayError.bind(this);
    this.login = this.login.bind(this);
    this.logout = this.logout.bind(this);
    this.getFileNames = this.getFileNames.bind(this);
    this.createTestMailFolder = this.createTestMailFolder.bind(this);
    this.switchToFrame1 = this.switchToFrame1.bind(this);
    this.switchToFrame2 = this.switchToFrame2.bind(this);
    this.switchToFrame3 = this.switchToFrame3.bind(this);
    this.deleteFamilyItem = this.deleteFamilyItem.bind(this); // P0312
    this.queryContainer = this.queryContainer.bind(this); // P7a8a
    this.checkInboxEmail = this.checkInboxEmail.bind(this); // Pe51a
    this.checkAndCreateUser = this.checkAndCreateUser.bind(this); // P1fdb
    // No need to bind createDatabase since it's an arrow function
  }

  /*
        Properties
    */

  // The access token is not part of state because React is concerned with the
  // UI and the token is not used to affect the UI in any way.
  accessToken: string;
  userName: string;

  listItems: HeroListItem[] = [
    {
      icon: "PlugConnected",
      primaryText: "Connects to OneDrive for Business.",
    },
    {
      icon: "Mail",
      primaryText:
        "Gets the names of the first three workbooks in OneDrive for Business.",
    },
    {
      icon: "Reply",
      primaryText: "Adds the names to the reply of an email.",
    },
  ];

  /*
        Methods
    */

  boundSetState: () => {};

  setToken = (accesstoken: string) => {
    this.accessToken = accesstoken;
  };

  setUserName = (userName: string) => {
    this.userName = userName;
  };

  displayError = (error: string) => {
    this.setState({ errorMessage: error, fileFetch: "notFetched" });
  };

  // Runs when the user clicks the X to close the message bar where
  // the error appears.
  errorDismissed = () => {
    this.setState({ errorMessage: "" });

    // If the error occurred during an "in process" phase (logging in or getting files),
    // the action didn't complete, so return the UI to the preceding state/view.
    this.setState((prevState) => {
      if (prevState.authStatus === "loginInProcess") {
        return { authStatus: "notLoggedIn" };
      } else if (prevState.fileFetch === "fetchInProcess") {
        return { fileFetch: "notFetched" };
      }
      return null;
    });
  };

  login = async () => {
    await signInO365(
      this.boundSetState,
      this.setToken,
      this.setUserName,
      this.displayError
    );
  };

  logout = async () => {
    await logoutFromO365(
      this.boundSetState,
      this.setUserName,
      this.userName,
      this.displayError
    );
  };

  getFileNames = async () => {
    this.setState({ fileFetch: "fetchInProcess" });
    try {
      let response = await getGraphData(
        // Get the `name` property of the first 3 Excel workbooks in the user's OneDrive.
        "https://graph.microsoft.com/v1.0/me/drive/root/microsoft.graph.search(q = '.xlsx')?$select=name&top=3",
        this.accessToken
      );
      await writeFileNamesToEmail(response, this.displayError);
      this.setState({ fileFetch: "fetched", headerMessage: "Success" });
    } catch (requestError) {
      // This error must be from the Axios request in getGraphData, 
      // not the Office.js in writeFileNamesToWorksheet.
      this.displayError(requestError);
    }
  };

  createTestMailFolder = async () => {
    if (!this.accessToken) {
      this.displayError('Access token is not set.');
      return;
    }

    try {
      const response = await createMailFolder(this.accessToken);
      console.log('Mail folder created:', response);
    } catch (error) {
      this.displayError('Failed to create mail folder.');
    }
  };

  // New method to call the backend endpoint and trigger createFamilyItem
  createFamilyItem = async () => {
    try {
      const response = await fetch('https://cosmosdbbackendplugin.azurewebsites.net/createFamilyItem');
      const text = await response.text();
      console.log(text);
      // Optionally update state or display a success message
      this.setState({ headerMessage: "Family addded successfully." });
    } catch (error) {
      console.error('Error adding Family:', error);
      this.displayError('Error adding family.');
    }
  };

  // New method to call the backend endpoint and trigger deleteFamilyItem
  deleteFamilyItem = async () => {
    try {
      const response = await fetch('https://cosmosdbbackendplugin.azurewebsites.net/deleteFamilyItem');
      const text = await response.text();
      console.log(text);
      // Optionally update state or display a success message
      this.setState({ headerMessage: "Family deleted successfully." });
    } catch (error) {
      console.error('Error deleting Family:', error);
      this.displayError('Error deleting family.');
    }
  };

  // New method to call the backend endpoint and trigger queryContainer
  queryContainer = async () => {
    try {
      const response = await fetch('https://cosmosdbbackendplugin.azurewebsites.net/queryContainer');
      const text = await response.text();
      console.log(text);
      // Optionally update state or display a success message
      this.setState({ headerMessage: "Query executed successfully." });
    } catch (error) {
      console.error('Error executing query:', error);
      this.displayError('Error executing query.');
    }
  };

  checkInboxEmail = async () => {
    try {
      const response = await getGraphData(
        "https://graph.microsoft.com/v1.0/me",
        this.accessToken
      );
      const emailAddress = response.data.mail;
      this.checkAndCreateUser(emailAddress);
    } catch (error) {
      this.displayError('Failed to get inbox email address.');
    }
  };

  checkAndCreateUser = async (emailAddress: string) => {
    try {
      const response = await fetch(`https://cosmosdbbackendplugin.azurewebsites.net/checkUser?email=${emailAddress}`);
      const result = await response.json();
      if (!result.exists) {
        await fetch('https://cosmosdbbackendplugin.azurewebsites.net/createUser', {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
          },
          body: JSON.stringify({ email: emailAddress }),
        });
      }
    } catch (error) {
      this.displayError('Failed to check or create user.');
    }
  };

  switchToFrame1 = () => {
    this.setState({ currentFrame: "Frame1" });
  };

// Modify `switchToFrame2` to accept `requestInput`
    switchToFrame2 = (requestInput: string) => {
    this.setState({ currentFrame: "Frame2", requestInput });
  };


  switchToFrame3 = () => {
    this.setState({ currentFrame: "Frame3" });
  };

  render() {
    const { title, isOfficeInitialized } = this.props;
  
    if (!isOfficeInitialized) {
      return (
        <Progress
          title={title}
          logo="assets/Onedrive_Charts_icon_80x80px.png"
          message="Please sideload your add-in to see app body."
        />
      );
    }

    // Set the body of the page based on where the user is in the workflow.
    let body;

    if (this.state.currentFrame === "Frame1") {
      body = <Frame1 switchToFrame2={this.switchToFrame2} displayError={this.displayError} accessToken={this.accessToken} />;
    } else if (this.state.currentFrame === "Frame2") {
      body = 
        <Frame2
          switchToFrame3={this.switchToFrame3}
          accessToken={this.accessToken}
          requestInput={this.state.requestInput} // Pass `requestInput` here
        />
      ;
    } else if (this.state.currentFrame === "Frame3") {
      body = 
        <Frame3
          accessToken={this.accessToken}
          requestInput={this.state.requestInput}
        />
    } else if (this.state.authStatus === "notLoggedIn") {
      body = <StartPageBody login={this.login} listItems={this.listItems} />;
    } else if (this.state.authStatus === "loginInProcess") {
      body = (
        <Spinner
          className="spinner"
          type={SpinnerType.large}
          label="Please sign-in on the pop-up window."
        />
      );
    } else {
      if (this.state.fileFetch === "notFetched") {
        body = (
          <GetDataPageBody
            getFileNames={this.getFileNames}
            logout={this.logout}
            createTestMailFolder={this.createTestMailFolder}
            createFamilyItem={this.createFamilyItem} // Pass the createDatabase method
            deleteFamilyItem={this.deleteFamilyItem} // P0312
            queryContainer={this.queryContainer} // P7a8a
            checkInboxEmail={this.checkInboxEmail} // Pc13f
          />
        );
      } else if (this.state.fileFetch === "fetchInProcess") {
        body = (
          <Spinner
            className="spinner"
            type={SpinnerType.large}
            label="We are getting the data for you."
          />
        );
      } else {
        body = (
          <SuccessPageBody
            getFileNames={this.getFileNames}
            logout={this.logout}
          />
        );
      }
    }

    return (
      <div>
        {this.state.errorMessage ? (
          <OfficeAddinMessageBar
            onDismiss={this.errorDismissed}
            message={this.state.errorMessage + " "}
          />
        ) : null}

        <div className="ms-welcome">
          <Header
            logo="assets/Onedrive_Charts_icon_80x80px.png"
            title={this.props.title}
            message={this.state.headerMessage}
          />
          {body}
          <button onClick={this.switchToFrame1} style={{ position: "fixed", bottom: 0, width: "100%" }}>
            Go to Home
          </button>
        </div>
      </div>
    );
  }
}

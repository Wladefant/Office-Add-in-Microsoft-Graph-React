import React, { useState, useEffect } from "react";
import { FluentProvider, webLightTheme, Text } from "@fluentui/react-components";
import axios from 'axios';

interface Frame3Props {
  accessToken: string;
  requestInput: string;
}

const Frame3: React.FC<Frame3Props> = ({ accessToken, requestInput }) => {
  const [objectName, setObjectName] = useState("");
  const [numberOfAcceptedEmails, setNumberOfAcceptedEmails] = useState(0);
  const [numberOfRejectedEmails, setNumberOfRejectedEmails] = useState(0);
  const [folderName, setFolderName] = useState("");
  const [savedTime, setSavedTime] = useState(0);

  const fetchObjectNameFromCosmosDB = async (outlookEmailId: string) => {
    try {
      const encodedEmailId = encodeURIComponent(outlookEmailId);
      const response = await fetch(
        `https://cosmosdbbackendplugin.azurewebsites.net/fetchName?outlookEmailId=${encodedEmailId}`
      );
      const result = await response.json();
      return result.objectname;
    } catch (error) {
      console.error("Error fetching objectname from CosmosDB:", error);
      return "Error fetching objectname.";
    }
  };

  const fetchFolderNameFromBackend = async (outlookEmailId: string) => {
    try {
      const encodedEmailId = encodeURIComponent(outlookEmailId);
      const response = await fetch(
        `https://cosmosdbbackendplugin.azurewebsites.net/fetchFolderName?outlookEmailId=${encodedEmailId}`
      );
      const result = await response.json();
      return result.folderName;
    } catch (error) {
      console.error("Error fetching folder name from backend:", error);
      return null;
    }
  };

  const fetchEmailsByFolderName = async (folderName: string) => {
    try {
      const encodedFolderName = encodeURIComponent(folderName);
      const response = await fetch(
        `https://cosmosdbbackendplugin.azurewebsites.net/fetchEmailsByFolderName?folderName=${encodedFolderName}`
      );
      const emails = await response.json();
      return emails;
    } catch (error) {
      console.error("Error fetching emails by folder name:", error);
      return [];
    }
  };

  useEffect(() => {
    const fetchData = async () => {
      if (Office.context.mailbox.item) {
        const restId = Office.context.mailbox.convertToRestId(
          Office.context.mailbox.item.itemId,
          Office.MailboxEnums.RestVersion.v2_0
        );

        // Fetch object name
        const objectName = await fetchObjectNameFromCosmosDB(restId);
        setObjectName(objectName);

        // Fetch folder name
        const folderName = await fetchFolderNameFromBackend(restId);
        setFolderName(folderName);

        if (folderName) {
          // Fetch emails in the folder
          const emails = await fetchEmailsByFolderName(folderName);

          // Parse requestInput to get the number of accepted emails
          const acceptedEmailsCount = parseInt(requestInput, 10) || 0;

          // Ensure acceptedEmailsCount does not exceed total emails
          const totalEmails = emails.length;
          const numberOfAcceptedEmails = Math.min(acceptedEmailsCount, totalEmails);

          // Calculate number of rejected emails
          const numberOfRejectedEmails = totalEmails - numberOfAcceptedEmails;

          setNumberOfAcceptedEmails(numberOfAcceptedEmails);
          setNumberOfRejectedEmails(numberOfRejectedEmails);

          // Calculate saved time (assuming 5 minutes saved per email)
          const savedTimePerEmail = 5;
          const totalSavedTime = totalEmails * savedTimePerEmail;
          setSavedTime(totalSavedTime);
        } else {
          setNumberOfAcceptedEmails(0);
          setNumberOfRejectedEmails(0);
          setSavedTime(0);
        }
      }
    };

    fetchData();

    const itemChangedHandler = () => {
      fetchData();
    };

    Office.context.mailbox.addHandlerAsync(
      Office.EventType.ItemChanged,
      itemChangedHandler
    );

    return () => {
      Office.context.mailbox.removeHandlerAsync(
        Office.EventType.ItemChanged,
        itemChangedHandler
      );
    };
  }, [requestInput]);

  return (
    <FluentProvider theme={webLightTheme}>
      <div style={{ padding: "40px 20px", maxWidth: "400px", margin: "0 auto", textAlign: "center" }}>
        {/* Congratulations Message */}
        <Text style={{ fontSize: "24px", fontWeight: "bold", marginBottom: "30px" }}>
          Glückwunsch!
        </Text>

        <p style={{ fontSize: "16px", marginBottom: "20px" }}>
          ImmoMail hat dir die {numberOfAcceptedEmails} E-Mails für die {numberOfAcceptedEmails} besten Bewerber in deinen Entwürfen unter "{folderName}" abgelegt.
        </p>

        <p style={{ fontSize: "16px", marginBottom: "20px" }}>
          Für alle abgelehnten Bewerber haben wir dir die Entwürfe in "{folderName}" abgelegt.
        </p>

        <p style={{ fontSize: "16px", marginBottom: "20px" }}>
          Du hast dir ca. {savedTime} Minuten Arbeitszeit gespart!
        </p>

        <p style={{ fontSize: "16px", marginBottom: "20px" }}>
          Objektname: {objectName}
        </p>
      </div>
    </FluentProvider>
  );
};

export default Frame3;

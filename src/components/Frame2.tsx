import React, { useState, useEffect } from "react";
import {
  FluentProvider,
  webLightTheme,
  Button,
  Input,
  Text,
  Textarea,
} from "@fluentui/react-components";
import MarkdownCard from "./MarkdownCard";
import { Configuration, OpenAIApi } from "openai";
import OPENAI_API_KEY from "../../config/openaiKey";
import axios from "axios";

interface Frame2Props {
  switchToFrame3: () => void;
  accessToken: string;
}

const Frame2: React.FC<Frame2Props> = ({ switchToFrame3, accessToken }) => {
  // State for the dynamic values
  const [propertyName, setPropertyName] = useState("Immobilie XXX");
  const [requestsInfo, setRequestsInfo] = useState(
    "XXX der XXX Anfragen treffen auf die Profilbeschreibung zu"
  );
  const [confirmationTemplate, setConfirmationTemplate] = useState("");
  const [rejectionTemplate, setRejectionTemplate] = useState("");
  const [customerProfile, setCustomerProfile] = useState("");

  const restId = Office.context.mailbox.convertToRestId(
    Office.context.mailbox.item.itemId,
    Office.MailboxEnums.RestVersion.v2_0
  );
  console.log("REST-formatted Item ID:", restId);
  const emailId =restId;
  // Fetch customer profile and property name
  useEffect(() => {
    const fetchEmailContent = async () => {
      if (Office.context.mailbox.item) {
        // Get the REST ID of the current email
        const restId = Office.context.mailbox.convertToRestId(
          Office.context.mailbox.item.itemId,
          Office.MailboxEnums.RestVersion.v2_0
        );

        // Fetch customer profile and property name
        const customerProfile = await fetchCustomerProfileFromBackend(restId);
        setCustomerProfile(customerProfile);
        const objectname = await fetchObjectNameFromCosmosDB(restId);
        setPropertyName(objectname);
      }
    };

    fetchEmailContent();

    const itemChangedHandler = () => {
      fetchEmailContent();
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
  }, []);

  // Function to fetch customer profile from backend
  const fetchCustomerProfileFromBackend = async (outlookEmailId: string) => {
    try {
      const encodedEmailId = encodeURIComponent(outlookEmailId);
      const response = await fetch(
        `https://cosmosdbbackendplugin.azurewebsites.net/fetchCustomerProfile?outlookEmailId=${encodedEmailId}`
      );
      const result = await response.json();
      return result.customerProfile;
    } catch (error) {
      console.error("Error fetching customer profile from backend:", error);
      return "Error fetching customer profile.";
    }
  };

  // Function to fetch object name from CosmosDB
  const fetchObjectNameFromCosmosDB = async (outlookEmailId: string) => {
    try {
      const encodedEmailId = encodeURIComponent(outlookEmailId);
      const response = await fetch(
        `https://cosmosdbbackendplugin.azurewebsites.net/fetchName?outlookEmailId=${encodedEmailId}`
      );
      const result = await response.json();
      return result.objectname;
    } catch (error) {
      console.error("Error fetching object name from CosmosDB:", error);
      return "Error fetching object name.";
    }
  };

  // Function to check if the "akzeptiert" folder exists, and create it if not
  const ensureAkzeptiertFolderExists = async (): Promise<string | null> => {
    try {
      // Check if the folder exists
      const response = await axios.get(
        "https://graph.microsoft.com/v1.0/me/mailFolders",
        {
          headers: {
            Authorization: `Bearer ${accessToken}`,
          },
        }
      );

      const folders = response.data.value;
      let folder = folders.find((f: any) => f.displayName === "akzeptiert");

      if (folder) {
        // Folder exists, return its ID
        return folder.id;
      } else {
        // Folder doesn't exist, create it
        const createFolderResponse = await axios.post(
          "https://graph.microsoft.com/v1.0/me/mailFolders",
          {
            displayName: "akzeptiert",
          },
          {
            headers: {
              Authorization: `Bearer ${accessToken}`,
              "Content-Type": "application/json",
            },
          }
        );
        return createFolderResponse.data.id;
      }
    } catch (error) {
      console.error("Error ensuring 'akzeptiert' folder exists:", error);
      return null;
    }
  };

  // Function to create a draft reply and move it to the "akzeptiert" folder
  const createDraftReplyAndMove = async () => {
    try {
      // Ensure the "akzeptiert" folder exists and get its ID
      const akzeptiertFolderId = await ensureAkzeptiertFolderExists();

      if (!akzeptiertFolderId) {
        console.error("Could not obtain 'akzeptiert' folder ID.");
        return;
      }

      // Create the draft reply
      const createDraftResponse = await axios.post(
        `https://graph.microsoft.com/v1.0/me/messages/${emailId}/createReply`,
        {
          comment: "akzeptiert",
        },
        {
          headers: {
            Authorization: `Bearer ${accessToken}`,
            "Content-Type": "application/json",
          },
        }
      );

      const draftMessageId = createDraftResponse.data.id;

      // Move the draft to the "akzeptiert" folder
      await axios.post(
        `https://graph.microsoft.com/v1.0/me/messages/${draftMessageId}/move`,
        {
          destinationId: akzeptiertFolderId,
        },
        {
          headers: {
            Authorization: `Bearer ${accessToken}`,
            "Content-Type": "application/json",
          },
        }
      );

      console.log("Draft reply created and moved to 'akzeptiert' folder.");

      // Switch to Frame3
      switchToFrame3();
    } catch (error) {
      console.error("Error creating draft reply and moving:", error);
    }
  };

  return (
    <FluentProvider theme={webLightTheme}>
      <div style={{ padding: "20px", maxWidth: "400px", margin: "0 auto" }}>
        {/* Property Information */}
        <MarkdownCard markdown={`**${propertyName}**`} />
        <MarkdownCard markdown={requestsInfo} />
        <MarkdownCard markdown={customerProfile} />

        {/* Templates */}
        <Textarea
          placeholder="Template für Bestätigungsemail"
          value={confirmationTemplate}
          onChange={(e) => setConfirmationTemplate(e.target.value)}
          style={{
            marginBottom: "10px",
            width: "100%",
            height: "100px",
          }}
        />
        <Textarea
          placeholder="Template für Absageemails"
          value={rejectionTemplate}
          onChange={(e) => setRejectionTemplate(e.target.value)}
          style={{
            marginBottom: "20px",
            width: "100%",
            height: "100px",
          }}
        />

        {/* Drafts Button */}
        <Button
          appearance="primary"
          style={{ width: "100%" }}
          onClick={createDraftReplyAndMove}
        >
          Drafts erstellen
        </Button>
      </div>
    </FluentProvider>
  );
};

export default Frame2;

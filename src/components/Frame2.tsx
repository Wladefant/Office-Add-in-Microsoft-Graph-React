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
  const [requestsInfo, setRequestsInfo] = useState("XXX der XXX Anfragen treffen auf die Profilbeschreibung zu");
  const [confirmationTemplate, setConfirmationTemplate] = useState("");
  const [rejectionTemplate, setRejectionTemplate] = useState("");
  const [customerProfile, setCustomerProfile] = useState("");

  const restId = Office.context.mailbox.convertToRestId(
    Office.context.mailbox.item.itemId,
    Office.MailboxEnums.RestVersion.v2_0
  );
  console.log("REST-formatted Item ID:", restId);
  const emailId =restId;
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
      // Fetch the folder name from Cosmos DB for the current email
      const folderName = await fetchFolderNameFromBackend(emailId);
  
      if (!folderName) {
        console.error("Could not obtain folder name from Cosmos DB.");
        return;
      }
  
      // Ensure the folder exists and get its ID
      const folderId = await ensureFolderExists(folderName);
  
      if (!folderId) {
        console.error(`Could not obtain folder ID for folder: ${folderName}`);
        return;
      }
  
      // Fetch all emails with the same folder name from Cosmos DB
      const emails = await fetchEmailsByFolderName(folderName);
  
      if (!emails || emails.length === 0) {
        console.log(`No emails found with folder name: ${folderName}`);
        return;
      }
  
      // For each email, create a draft reply and move it to the folder
      for (const email of emails) {
        await createDraftReplyForEmail(email.outlookEmailId, folderId);
      }
  
      console.log("Draft replies created and moved to the folder.");
  
      // Switch to Frame3
      switchToFrame3();
    } catch (error) {
      console.error("Error creating draft replies:", error);
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
  
  const ensureFolderExists = async (folderName: string): Promise<string | null> => {
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
      let folder = folders.find((f: any) => f.displayName === folderName);
  
      if (folder) {
        // Folder exists, return its ID
        return folder.id;
      } else {
        // Folder doesn't exist, create it
        const createFolderResponse = await axios.post(
          "https://graph.microsoft.com/v1.0/me/mailFolders",
          {
            displayName: folderName,
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
      console.error(`Error ensuring folder '${folderName}' exists:`, error);
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

  const createDraftReplyForEmail = async (emailId: string, folderId: string) => {
    try {
      // Create the draft reply
      const createDraftResponse = await axios.post(
        `https://graph.microsoft.com/v1.0/me/messages/${emailId}/createReply`,
        {
          comment: "Your reply here", // Customize the reply content as needed
        },
        {
          headers: {
            Authorization: `Bearer ${accessToken}`,
            "Content-Type": "application/json",
          },
        }
      );
  
      const draftMessageId = createDraftResponse.data.id;
  
      // Move the draft to the folder
      await axios.post(
        `https://graph.microsoft.com/v1.0/me/messages/${draftMessageId}/move`,
        {
          destinationId: folderId,
        },
        {
          headers: {
            Authorization: `Bearer ${accessToken}`,
            "Content-Type": "application/json",
          },
        }
      );
  
      console.log(`Draft reply for email ${emailId} created and moved.`);
    } catch (error) {
      console.error(`Error creating draft reply for email ${emailId}:`, error);
    }
  };
  
  
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

  useEffect(() => {
    let intervalId: NodeJS.Timeout;

    const pollForEmailUpdate = async () => {
      const customerProfile = await fetchCustomerProfileFromBackend(emailId);
      const objectname = await fetchObjectNameFromCosmosDB(emailId);

      if (customerProfile && objectname) {
        // Update the state with new data
        setCustomerProfile(customerProfile);
        setPropertyName(objectname);
        
        // Clear the polling interval once the email is fetched
        clearInterval(intervalId);
      }
    };

    // Start polling every 2 seconds
    intervalId = setInterval(pollForEmailUpdate, 2000);

    return () => {
      // Clean up the interval on component unmount
      clearInterval(intervalId);
    };
  }, [emailId]);

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
            width: '100%',
            height: '100px',
          }}
        />
        <Textarea
          placeholder="Template für Absageemails"
          value={rejectionTemplate}
          onChange={(e) => setRejectionTemplate(e.target.value)}
          style={{
            marginBottom: "20px",
            width: '100%',
            height: '100px',
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

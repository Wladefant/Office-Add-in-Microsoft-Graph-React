import React, { useState, useEffect } from "react";
import {
  FluentProvider,
  webLightTheme,
  Button,
  Input,
  Textarea,
} from "@fluentui/react-components";
import { Text } from "@fluentui/react";
import { Configuration, OpenAIApi } from "openai";
import axios from 'axios';
import OPENAI_API_KEY from "../../config/openaiKey";
import MarkdownCard from "./MarkdownCard";

interface Frame1Props {
  switchToFrame2: () => void;
  displayError: (error: string) => void;
  accessToken: string;
}

const Frame1: React.FC<Frame1Props> = ({ switchToFrame2, displayError, accessToken }) => {
  const [location, setLocation] = useState("bisher nicht gespeichert");
  const [name, setName] = useState("bisher nicht gespeichert"); // Added state for name
  const [requests, setRequests] = useState("bisher nicht gespeichert");
  const [perfectCustomerProfile, setPerfectCustomerProfile] = useState("");
  const [requestInput, setRequestInput] = useState("");
  const [customerProfile, setCustomerProfile] = useState("noch nicht gespeichert");


  const determineLocation = async (emailContent: string) => {
    const configuration = new Configuration({
      apiKey: OPENAI_API_KEY,
    });
    const openai = new OpenAIApi(configuration);

    try {
      const response = await openai.createChatCompletion({
        model: "gpt-4o-mini",
        messages: [
          {
            role: "system",
            content: "Du bist ein hilfreicher Assistent, der den Ort aus E-Mail-Inhalten extrahiert.",
          },
          {
            role: "user",
            content: `Bestimme den Ort aus dem folgenden E-Mail-Inhalt und gib als output nur die adresse wieder, falls nicht gefunden "nicht gefunden": ${emailContent}`,
          },
        ],
        max_tokens: 50,
      });

      if (response.data.choices && response.data.choices[0].message) {
        const determinedLocation = response.data.choices[0].message.content.trim();
        setLocation(determinedLocation);
        return determinedLocation;
      } else {
        throw new Error("Unexpected API response structure");
      }
    } catch (error) {
      console.error("Error determining location:", error);
      return "Error determining location.";
    }
  };

  const determineName = async (emailContent: string) => {
    const configuration = new Configuration({
      apiKey: OPENAI_API_KEY,
    });
    const openai = new OpenAIApi(configuration);

    try {
      const response = await openai.createChatCompletion({
        model: "gpt-4o-mini",
        messages: [
          {
            role: "system",
            content: "Du bist ein hilfreicher Assistent, der den Namen der Immobilie aus E-Mail-Inhalten extrahiert.",
          },
          {
            role: "user",
            content: `Bestimme den Namen der Immobilie aus dem folgenden E-Mail-Inhalt und gib als output nur den Namen wieder, falls nicht gefunden "nicht gefunden": ${emailContent}`,
          },
        ],
        max_tokens: 50,
      });

      if (response.data.choices && response.data.choices[0].message) {
        const determinedName = response.data.choices[0].message.content.trim();
        setName(determinedName);
        return determinedName;
      } else {
        throw new Error("Unexpected API response structure");
      }
    } catch (error) {
      console.error("Error determining name:", error);
      return "Error determining name.";
    }
  };

  const determineCustomerProfile = async (emailContent: string) => {
    const configuration = new Configuration({
      apiKey: OPENAI_API_KEY,
    });
    const openai = new OpenAIApi(configuration);

    try {
      const response = await openai.createChatCompletion({
        model: "gpt-4o-mini",
        messages: [
          {
            role: "system",
            content: "Du bist ein hilfreicher Assistent, der E-Mails zusammenfasst und Mieter anhand ihres Profils bewertet.",
          },
          {
            role: "user",
            content: `Gib eine kurze, strukturierte Beschreibung zu dem Mieter auf Deutsch und bewerte den Mieter auf einer Skala von 1 bis 10, wobei 10 der wünschenswerteste Mieter ist.  Keine Titel oder Sonstiges, strukuriert und kompakt in Stickpunkten: ${emailContent}`,
          },
        ],
        max_tokens: 200,
      });

      if (response.data.choices && response.data.choices[0].message) {
        return response.data.choices[0].message.content.trim();
      } else {
        throw new Error("Unexpected API response structure");
      }
    } catch (error) {
      console.error("Error generating summary:", error);
      return "Error generating summary.";
    }
  };


  const checkEmailExistsInCosmosDB = async (outlookEmailId: string) => {
    try {
      const response = await fetch(`https://cosmosdbbackendplugin.azurewebsites.net/checkEmail?outlookEmailId=${outlookEmailId}`);
      const result = await response.json();
      return result.exists;
    } catch (error) {
      console.error('Error checking email existence in CosmosDB:', error);
      return false;
    }
  };

  const uploadEmailToCosmosDB = async (emailData: any) => {
    try {
      const response = await fetch('https://cosmosdbbackendplugin.azurewebsites.net/uploadEmail', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify(emailData),
      });
  
      if (!response.ok) {
        throw new Error('Failed to upload email to CosmosDB');
      }
  
      const data = await response.json();
      console.log('Email uploaded successfully:', data);
    } catch (error) {
      console.error('Error uploading email to server:', error);
    }
  };

  const changeIDanduploademail = async () => {
    try {
      const restId = Office.context.mailbox.convertToRestId(
        Office.context.mailbox.item.itemId,
        Office.MailboxEnums.RestVersion.v2_0
      );
      const fetchResponse = await fetch(`https://cosmosdbbackendplugin.azurewebsites.net/fetchDocument?outlookEmailId=${restId}`);
      if (!fetchResponse.ok) {
        throw new Error('Failed to fetch document from CosmosDB');
      }
      const fetchedDocument = await fetchResponse.json();
      console.log('Fetched document:', fetchedDocument);
  
      // Delete the fetched document
      const deleteResponse = await fetch(`https://cosmosdbbackendplugin.azurewebsites.net/deleteFamilyItem?id=${fetchedDocument.id}`, {
        method: 'GET',
      });
      if (!deleteResponse.ok) {
        throw new Error('Failed to delete document from CosmosDB');
      }
      console.log('Document deleted successfully');
  
      // Create a new document with the same properties and fields but with a new id
      const newDocument = { ...fetchedDocument, id: fetchedDocument.id, location: "" };
  
      // Upload the new document to CosmosDB
      const uploadResponse = await fetch('https://cosmosdbbackendplugin.azurewebsites.net/uploadEmail', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify(newDocument),
      });
  
      if (!uploadResponse.ok) {
        throw new Error('Failed to upload email to CosmosDB');
      }
  
      const data = await uploadResponse.json();
      console.log('Email uploaded successfully:', data);
    } catch (error) {
      console.error('Error uploading email to server:', error);
    }
  };
  
  const fetchLocationFromCosmosDB = async (outlookEmailId: string) => {
    try {
      const encodedEmailId = encodeURIComponent(outlookEmailId);
      const response = await fetch(
        `https://cosmosdbbackendplugin.azurewebsites.net/fetchLocation?outlookEmailId=${encodedEmailId}`
      );
      console.log("Fetching location for Email ID:", outlookEmailId);
      const result = await response.json();
      console.log("Fetch result:", result);
      return result.location;
    } catch (error) {
      console.error("Error fetching location from CosmosDB:", error);
      return "Error fetching location.";
    }
  };

  const fetchNameFromCosmosDB = async (outlookEmailId: string) => {
    try {
      const encodedEmailId = encodeURIComponent(outlookEmailId);
      const response = await fetch(
        `https://cosmosdbbackendplugin.azurewebsites.net/fetchName?outlookEmailId=${encodedEmailId}`
      );
      console.log("Fetching name for Email ID:", outlookEmailId);
      const result = await response.json();
      console.log("Fetch result:", result);
      return result.objectname;
    } catch (error) {
      console.error("Error fetching name from CosmosDB:", error);
      return "Error fetching name.";
    }
  };

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
  const handleAnalyseClick = async () => {
    try {
      // Fetch emails from the Graph API
      const response = await axios.get(
        "https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages?$filter=from/emailAddress/address eq 'w.kirjanovs@realest-ai.com'",
        {
          headers: {
            Authorization: `Bearer ${accessToken}`,
          },
        }
      );
  
      const emails = response.data.value;
  
      // Process each email
      for (const email of emails) {
        const emailExists = await checkEmailExistsInCosmosDB(email.id);
        if (!emailExists) {
          // Extract details if the email does not exist
          const location = await determineLocation(email.body.content);
          const name = await determineName(email.body.content);
          const customerProfile = await determineCustomerProfile(email.body.content);
  
          const emailData = {
            subject: email.subject,
            userId: email.from.emailAddress.address,
            receivedAt: email.receivedDateTime,
            sent: false,
            location: location,
            objectname: name,
            customerProfile: customerProfile,
            outlookEmailId: email.id,
            emailBody: email.body.content,
          };
  
          // Upload to CosmosDB only if the email doesn't already exist
          await uploadEmailToCosmosDB(emailData);
          console.log(`Uploaded email: ${email.id}`);
        } else {
          console.log(`Email ${email.id} already exists, skipping upload.`);
        }
      }
  
      // Use the Office JavaScript API to get the current email's ID
      const item = Office.context.mailbox.item;
      if (item) {
        // Get the REST-formatted ID of the current item
        const restId = Office.context.mailbox.convertToRestId(
          item.itemId,
          Office.MailboxEnums.RestVersion.v2_0
        );
  
        // Display the success message along with the current email's ID
      } else {
        // If no item is available, display a different message
        displayError("Emails processed successfully. No email is currently open.");
      }
    } catch (error) {
      displayError(error.toString());
    }
  
    await changeIDanduploademail();
  };
  

  useEffect(() => {
    const fetchEmailContent = async () => {
      console.log("fetching email content");
      if (Office.context.mailbox.item) {
        console.log("context is on email");
        console.log("Original Item ID:", Office.context.mailbox.item.itemId);
  
        // Convert the item ID to REST format
        const restId = Office.context.mailbox.convertToRestId(
          Office.context.mailbox.item.itemId,
          Office.MailboxEnums.RestVersion.v2_0
        );
        console.log("REST-formatted Item ID:", restId);
  
        // Fetch the location using the REST-formatted ID
        const location = await fetchLocationFromCosmosDB(restId);
        console.log("Fetched Location:", location);
        setLocation(location);

        const name = await fetchNameFromCosmosDB(restId);
        console.log("Fetched Name:", name);
        setName(name);

        const customerProfile = await fetchCustomerProfileFromBackend(restId);
        setCustomerProfile(customerProfile);
      }
    };
  
    fetchEmailContent();

    const itemChangedHandler = () => {
      fetchEmailContent();
    };

    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, itemChangedHandler);

    return () => {
      Office.context.mailbox.removeHandlerAsync(Office.EventType.ItemChanged, itemChangedHandler);
    };
  }, []);
  

  return (
    <FluentProvider theme={webLightTheme}>
      <div style={{ padding: "20px", width: "calc(100% - 40px)", margin: "0 auto" }}>
        
        <MarkdownCard markdown={`**Ort:** ${location}\n\n**Name:** ${name}\n\n **${requests}** Anfragen gefunden.`} />

        <Textarea
          placeholder="Beschreiben sie die Voraussetzungen für den perfekten Kunden"
          value={perfectCustomerProfile}
          onChange={(e) => setPerfectCustomerProfile(e.target.value)}
          style={{
            marginBottom: '20px',
            width: '100%', // Ensure the input takes the full width of its container
            height: '100px', // Fixed height to allow for multiple lines
          }}
        />
        <MarkdownCard markdown={customerProfile} />

        <Text style={{ fontSize: "16px", marginBottom: "10px" }}>
          Anzahl der akzeptierten Anfragen:
        </Text>
        <Input
          placeholder="Geben sie ein Zahl ein"
          value={requestInput}
          onChange={(e) => setRequestInput(e.target.value)}
          style={{
            marginBottom: "20px",
            width: '100%', // Ensure the input takes the full width of its container
          }}
        />
       

        <Button
          appearance="primary"
          style={{ width: "100%" }}
          onClick={() => {
            handleAnalyseClick();
            switchToFrame2();
          }}
        >
          Analyse durchführen
        </Button>
      </div>
    </FluentProvider>
  );
};

export default Frame1;

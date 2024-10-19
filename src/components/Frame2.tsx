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

  const generateSummary = async (emailContent: string) => {
    const configuration = new Configuration({
      apiKey: OPENAI_API_KEY,
    });
    const openai = new OpenAIApi(configuration);

    try {
      const response = await openai.createChatCompletion({
        model: "gpt-4o", // or 'gpt-4' if you have access
        messages: [
          {
            role: "system",
            content: "Du bist ein hilfreicher Assistent, der E-Mails zusammenfasst und Mieter anhand ihres Profils bewertet.",
          },
          {
            role: "user",
            content: `Gib eine kurze, strukturierte Beschreibung zu dem Mieter auf Deutsch und bewerte den Mieter auf einer Skala von 1 bis 10, wobei 10 der wünschenswerteste Mieter ist. Output in Markdown. Keine Titel oder Sonstiges, strukuriert und kompakt in Stickpunkten: ${emailContent}`,
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

  // Function to create a draft reply
  const createDraftReply = async () => {
    try {
      // Make API call to create draft reply
      await axios.post(
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

      console.log("Draft reply created successfully.");

      // Switch to Frame3
      switchToFrame3();
    } catch (error) {
      console.error("Error creating draft reply:", error);
    }
  };
  useEffect(() => {
    const fetchEmailContent = async () => {
      if (Office.context.mailbox.item) {
        const restId = Office.context.mailbox.convertToRestId(
          Office.context.mailbox.item.itemId,
          Office.MailboxEnums.RestVersion.v2_0
        );
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

    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, itemChangedHandler);

    return () => {
      Office.context.mailbox.removeHandlerAsync(Office.EventType.ItemChanged, itemChangedHandler);
    };
  }, []);

  // Dummy data for the top results
  const topResults = [
    { name: "Name", platform: "Plattform", description: "Kurze Beschreibung" },
    { name: "Name", platform: "Plattform", description: "Kurze Beschreibung" },
    { name: "Name", platform: "Plattform", description: "Kurze Beschreibung" },
  ];

  return (
    <FluentProvider theme={webLightTheme}>
      <div style={{ padding: "20px", maxWidth: "400px", margin: "0 auto" }}>
        {/* Logo and Title */}
        

        {/* Property Information */}
        <MarkdownCard markdown={`**${propertyName}**`} />
        <MarkdownCard markdown={requestsInfo} />

        {/* Top Results */}
        {/* <Text style={{ fontSize: "16px", fontWeight: "bold", marginBottom: "10px" }}>
          Top XXX Treffer:
        </Text>
        <div style={{ marginBottom: "20px" }}>
          {topResults.map((result, index) => (
            <MarkdownCard key={index} markdown={`**${result.name}**\n\n${result.platform}\n\n${result.description}`} />
          ))}
        </div> */}
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
          onClick={createDraftReply}
        >
          Drafts erstellen
        </Button>
      </div>
    </FluentProvider>
  );
};

export default Frame2;

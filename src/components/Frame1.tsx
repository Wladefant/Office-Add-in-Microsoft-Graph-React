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
import OPENAI_API_KEY from "../../config/openaiKey";
import MarkdownCard from "./MarkdownCard";

interface Frame1Props {
  switchToFrame2: () => void;
}

const Frame1: React.FC<Frame1Props> = ({ switchToFrame2 }) => {
  const [location, setLocation] = useState("xxx");
  const [requests, setRequests] = useState("XXX");
  const [perfectCustomerProfile, setPerfectCustomerProfile] = useState("");
  const [requestInput, setRequestInput] = useState("");

  const determineLocation = async (emailContent: string) => {
    const configuration = new Configuration({
      apiKey: OPENAI_API_KEY,
    });
    const openai = new OpenAIApi(configuration);

    try {
      const response = await openai.createChatCompletion({
        model: "gpt-4o",
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
        await saveLocationToCosmosDB(determinedLocation); // Save location to CosmosDB
        return determinedLocation;
      } else {
        throw new Error("Unexpected API response structure");
      }
    } catch (error) {
      console.error("Error determining location:", error);
      return "Error determining location.";
    }
  };

  // Function to save location to CosmosDB
  const saveLocationToCosmosDB = async (location: string) => {
    if (!location || location === "nicht gefunden") return;
  
    try {
      const itemId = Office.context.mailbox.item.itemId; // Get the email ID
  
      // Send a POST request to your server
      const response = await fetch('http://localhost:5000/save-location', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({ location, emailId: itemId }),
      });
  
      if (!response.ok) {
        throw new Error('Failed to save location to CosmosDB');
      }
  
      const data = await response.json();
      console.log('Location saved successfully:', data);
    } catch (error) {
      console.error('Error saving location to server:', error);
    }
  };
  

  useEffect(() => {
    const fetchEmailContent = async () => {
      if (Office.context.mailbox.item) {
        Office.context.mailbox.item.body.getAsync("text", (result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            determineLocation(result.value).then((location) => {
              setLocation(location);
            });
          } else {
            console.error("Error fetching email content:", result.error);
          }
        });
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
        
        <MarkdownCard markdown={`**Ort:** ${location}\n\n **${requests}** Anfragen gefunden.`} />

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
          onClick={switchToFrame2}
        >
          Analyse durchführen
        </Button>
      </div>
    </FluentProvider>
  );
};

export default Frame1;
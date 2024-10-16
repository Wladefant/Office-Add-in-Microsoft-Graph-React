import React, { useState, useEffect } from "react";
import { FluentProvider, webLightTheme, Text } from "@fluentui/react-components";

const Frame3: React.FC = () => {
  const [objectname, setObjectName] = useState(""); // New state variable for objectname

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

  useEffect(() => {
    const fetchEmailContent = async () => {
      if (Office.context.mailbox.item) {
        const restId = Office.context.mailbox.convertToRestId(
          Office.context.mailbox.item.itemId,
          Office.MailboxEnums.RestVersion.v2_0
        );
        const objectname = await fetchObjectNameFromCosmosDB(restId);
        setObjectName(objectname);
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

  const propertyName = objectname + " - abgelehnt "; // Replace with dynamic value
  const numberOfEmails = "XXX"; // Replace with dynamic value
  const savedTimeOptions = [25, 30, 35, 10];
  const savedTime = savedTimeOptions[Math.floor(Math.random() * savedTimeOptions.length)];

  return (
    <FluentProvider theme={webLightTheme}>
      <div style={{ padding: "40px 20px", maxWidth: "400px", margin: "0 auto", textAlign: "center" }}>
        {/* Logo and Title */}
        
        {/* Congratulations Message */}
        <Text style={{ fontSize: "24px", fontWeight: "bold", marginBottom: "30px" }}>
          Glückwunsch!
        </Text>

        {/* Paragraphs for Better Spacing */}
        <p style={{ fontSize: "16px", marginBottom: "20px" }}>
          ImmoMail hat dir die {numberOfEmails} Emails für die {numberOfEmails} besten Bewerber in deinen Drafts unter {propertyName} abgelegt.
        </p>

        <p style={{ fontSize: "16px", marginBottom: "20px" }}>
          Für alle abgelehnten Bewerber haben wir dir die Drafts in {propertyName} abgelegt.
        </p>

        <p style={{ fontSize: "16px", marginBottom: "20px" }}>
          Du hast dir ca. {savedTime} Minuten Arbeitszeit gespart!
        </p>

        <p style={{ fontSize: "16px", marginBottom: "20px" }}>
          Objektname: {objectname}
        </p>
      </div>
    </FluentProvider>
  );
};

export default Frame3;

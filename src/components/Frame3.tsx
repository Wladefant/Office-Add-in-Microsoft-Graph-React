import * as React from "react";
import { FluentProvider, webLightTheme, Text } from "@fluentui/react-components";

const Frame3: React.FC = () => {
  const propertyName = "XXX (Immobilienname)"; // Replace with dynamic value
  const numberOfEmails = "XXX"; // Replace with dynamic value
  const savedTime = "XXX"; // Replace with dynamic value

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
      </div>
    </FluentProvider>
  );
};

export default Frame3;

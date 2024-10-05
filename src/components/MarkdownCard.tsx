import React from 'react';
import ReactMarkdown from 'react-markdown';
import { Card } from '@fluentui/react-components';

interface MarkdownCardProps {
  markdown: string;
}

const MarkdownCard: React.FC<MarkdownCardProps> = ({ markdown }) => {
  return (
    <Card
      style={{
        marginBottom: '20px',
        padding: '20px',
        overflow: 'auto',
        maxHeight: '400px', // Set max height to trigger scrollbar on overflow
      }}
    >
      <ReactMarkdown>{markdown}</ReactMarkdown>
    </Card>
  );
};

export default MarkdownCard;

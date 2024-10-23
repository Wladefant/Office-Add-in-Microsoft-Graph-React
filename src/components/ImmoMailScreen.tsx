// ImmoMailScreen.tsx
import * as React from 'react';
import { Stack, Text, PrimaryButton } from '@fluentui/react';

interface ImmoMailScreenProps {
  login: () => void;
}

const ImmoMailScreen: React.FunctionComponent<ImmoMailScreenProps> = ({ login }) => {
  return (
    <Stack
      verticalAlign="center"
      horizontalAlign="center"
      styles={{
        root: {
          height: '100vh',
          backgroundColor: '#ffffff',
        },
      }}
    >
      {/* Icon */}
      <Stack.Item>
        <img
          src="data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTMyIiBoZWlnaHQ9IjEyNSIgdmlld0JveD0iMCAwIDEzMiAxMjUiIGZpbGw9Im5vbmUiIHhtbG5zPSJodHRwOi8vd3d3LnczLm9yZy8yMDAwL3N2ZyI+CjxnIGNsaXAtcGF0aD0idXJsKCNjbGlwMF85OV84MCkiPgo8cGF0aCBkPSJNMTI0LjU5MiAxMjIuMzE1SDcuNDA4MTZDNC44MDg1NyAxMjIuMzE1IDIuNjkzODggMTIwLjIwNyAyLjY5Mzg4IDExNy42MTZWNTAuNzU5OUMyLjY5Mzg4IDQ5LjMyMzQgMy4zNTM4OCA0Ny45NTQgNC40OTg3OCA0Ny4wNjhMNTcuMTY0MSA1LjczMjQ3QzYyLjM0OTggMS42NjQ3IDY5LjY2MzcgMS42NjQ3IDc0Ljg0OTQgNS43MzI0N0wxMjcuNTE1IDQ3LjA1NDZDMTI4LjY0NiA0Ny45NDA2IDEyOS4zMiA0OS4zMSAxMjkuMzIgNTAuNzQ2NFYxMTcuNjAzQzEyOS4zMiAxMjAuMTk0IDEyNy4yMDUgMTIyLjMwMiAxMjQuNjA1IDEyMi4zMDJMMTI0LjU5MiAxMjIuMzE1WiIgZmlsbD0id2hpdGUiIHN0cm9rZT0iIzIzMUYyMCIgc3Ryb2tlLXdpZHRoPSI0IiBzdHJva2UtbWl0ZXJsaW1pdD0iMTAiLz4KPHBhdGggZD0iTTQuMDQwODIgNDkuODJMNTIuNTMwNiA4Ni4wNjc1QzU0Ljc1MzEgODcuMTE0NiA1OC40OTc2IDg4LjUyNDMgNjMuMzA2MSA4OC43NTI1QzY5LjU1NTkgODkuMDQ3OCA3NC4zNzggODcuMTk1MiA3Ni43NzU1IDg2LjA2NzVDODcuNTM3NiA3OS4wMzI4IDEwNy43NTUgNjMuMTEwOCAxMjYuNjEyIDQ4LjQ3NzUiIGZpbGw9IndoaXRlIi8+CjxwYXRoIGQ9Ik00LjA0MDgyIDQ5LjgyTDUyLjUzMDYgODYuMDY3NUM1NC43NTMxIDg3LjExNDYgNTguNDk3NiA4OC41MjQzIDYzLjMwNjEgODguNzUyNUM2OS41NTU5IDg5LjA0NzggNzQuMzc4IDg3LjE5NTIgNzYuNzc1NSA4Ni4wNjc1Qzg3LjUzNzYgNzkuMDMyOCAxMDcuNzU1IDYzLjExMDggMTI2LjYxMiA0OC40Nzc1IiBzdHJva2U9IiMyMzFGMjAiIHN0cm9rZS13aWR0aD0iNCIgc3Ryb2tlLW1pdGVybGltaXQ9IjEwIi8+CjxwYXRoIGQ9Ik01LjM4Nzc2IDExOS42M0w1Ny45MTg0IDg3LjQxMDIiIHN0cm9rZT0iIzIzMUYyMCIgc3Ryb2tlLXdpZHRoPSI0IiBzdHJva2UtbWl0ZXJsaW1pdD0iMTAiLz4KPHBhdGggZD0iTTc0LjA4MTYgODcuNDEwMkwxMjcuOTU5IDEyMC45NzMiIHN0cm9rZT0iIzIzMUYyMCIgc3Ryb2tlLXdpZHRoPSI0IiBzdHJva2UtbWl0ZXJsaW1pdD0iMTAiLz4KPHBhdGggZD0iTTIyLjg5OCA4LjIwMjY0VjMxLjAyNTFNIzIzMUYyMCIgc3Ryb2tlLXdpZHRoPSI0IiBzdHJva2UtbWl0ZXJsaW1pdD0iMTAiLz4KPHBhdGggZD0iTTM5LjA2MTIgNi44NjAxMVYyMC4yODUxIiBzdHJva2U9IiMyMzFGMjAiIHN0cm9rZS13aWR0aD0iNCIgc3Ryb2tlLW1pdGVybGltaXQ9IjEwIi8+CjxwYXRoIGQ9Ik0yMC4yMDQxIDYuODYwMTFIMzkuMDQwOCIgZmlsbD0id2hpdGUiIHN0cm9rZT0iIzIzMUYyMCIgc3Ryb2tlLXdpZHRoPSI0IiBzdHJva2UtbWl0ZXJsaW1pdD0iMTAiLz4KPHBhdGggZD0iTTQuMDQwODIgNDkuODJMMTI2LjYxMiA0OC40Nzc1IiBzdHJva2U9IiMyMzFGMjAiIHN0cm9rZS13aWR0aD0iNCIgc3Ryb2tlLW1pdGVybGltaXQ9IjEwIi8+CjwvZz4KPC9zdmc+" // Embedded SVG data
          alt="ImmoMail Icon"
          width={100}
          height={100}
        />
      </Stack.Item>

      {/* Text */}
      <Stack.Item>
        <Text variant="xLarge" styles={{ root: { fontWeight: 'bold' } }}>
          Willkommen bei <span style={{ color: '#231F20' }}>ImmoMail</span>
        </Text>
      </Stack.Item>
      <Stack.Item>
        <Text variant="medium" styles={{ root: { textAlign: 'center', padding: '10px 0' } }}>
          Wir helfen dir deine Mieteranfragen innerhalb weniger Minuten abzuarbeiten!
        </Text>
      </Stack.Item>

      {/* Button */}
      <Stack.Item>
        <PrimaryButton
          text="Mit Office 365 Verbinden"
          onClick={login}
          styles={{
            root: {
              backgroundColor: '#000000',
              borderColor: '#000000',
              color: 'white',
              fontSize: '16px',
              padding: '12px 24px',
              borderRadius: '5px',
            },
            rootHovered: {
              backgroundColor: '#1a1a1a',
              borderColor: '#1a1a1a',
            },
          }}
        />
      </Stack.Item>

      {/* Feedback button (optional) */}
      <Stack.Item
        styles={{
          root: {
            position: 'absolute',
            bottom: '10px',
            left: '10px',
            color: '#999999',
            cursor: 'pointer',
          },
        }}
      >
        <Text variant="small">Feedback & Fragen</Text>
      </Stack.Item>
    </Stack>
  );
};

export default ImmoMailScreen;

var config = {}

config.endpoint = ''
config.key = ''

config.database = {
  id: 'ToDoList'
}

config.container = {
  id: 'Items',
  users: {
    id: 'Users'
  },
  emails: {
    id: 'Emails'
  }
}

config.items = {
  Andersen: {
    id: 'Anderson.1',
    Country: 'USA',
    partitionKey: 'USA',
    lastName: 'Andersen',
    parents: [
      {
        firstName: 'Thomas'
      },
      {
        firstName: 'Mary Kay'
      }
    ],
    children: [
      {
        firstName: 'Henriette Thaulow',
        gender: 'female',
        grade: 5,
        pets: [
          {
            givenName: 'Fluffy'
          }
        ]
      }
    ],
    address: {
      state: 'WA',
      county: 'King',
      city: 'Seattle'
    }
  },
  Wakefield: {
    id: 'Wakefield.7',
    partitionKey: 'Italy',
    Country: 'Italy',
    parents: [
      {
        familyName: 'Wakefield',
        firstName: 'Robin'
      },
      {
        familyName: 'Miller',
        firstName: 'Ben'
      }
    ],
    children: [
      {
        familyName: 'Merriam',
        firstName: 'Jesse',
        gender: 'female',
        grade: 8,
        pets: [
          {
            givenName: 'Goofy'
          },
          {
            givenName: 'Shadow'
          }
        ]
      },
      {
        familyName: 'Miller',
        firstName: 'Lisa',
        gender: 'female',
        grade: 1
      }
    ],
    address: {
      state: 'NY',
      county: 'Manhattan',
      city: 'NY'
    },
    isRegistered: false
  },
  users: {
    user1: {
      userId: 'user1',
      name: 'John Doe',
      email: 'johndoe@example.com',
      settings: {
        notificationPreferences: 'email',
        defaultReplyTemplate: 'Thank you for your inquiry about [property]. I will get back to you shortly.'
      }
    },
    user2: {
      userId: 'user2',
      name: 'Jane Smith',
      email: 'janesmith@example.com',
      settings: {
        notificationPreferences: 'sms',
        defaultReplyTemplate: 'Hello, thanks for reaching out!'
      }
    }
  },
  emails: {
    email1: {
      id: 'email1',
      userId: 'user1',
      emailData: 'Dear Sir/Madam, I am interested in the apartment located at XYZ Street...',
      summary: 'Inquiry about apartment at XYZ Street.',
      location: 'XYZ Street, City',
      sent: false,
      platform: 'PlatformA',
      outlookEmailId: 'outlook-id-123',
      draftFolderPath: '/Drafts',
      receivedAt: '2023-10-01T12:34:56Z',
      processedAt: '2023-10-01T12:35:30Z'
    },
    email2: {
      id: 'email2',
      userId: 'user2',
      emailData: 'Hello, I saw your listing for the house on ABC Avenue...',
      summary: 'Inquiry about house on ABC Avenue.',
      location: 'ABC Avenue, City',
      sent: true,
      platform: 'PlatformB',
      outlookEmailId: 'outlook-id-456',
      draftFolderPath: '/Sent Items',
      receivedAt: '2023-10-02T08:15:00Z',
      processedAt: '2023-10-02T08:16:00Z'
    }
  }
}

module.exports = config

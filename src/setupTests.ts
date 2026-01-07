/**
 * Jest Setup File
 *
 * Configures testing environment and mocks for Office.js
 */

// Mock Office.js global object
(global as any).Office = {
  onReady: jest.fn((callback) => {
    callback({
      host: 'Outlook',
      platform: 'PC'
    });
    return Promise.resolve({
      host: 'Outlook',
      platform: 'PC'
    });
  }),
  context: {
    mailbox: {
      item: {
        itemType: 'message',
        subject: 'Test Email Subject',
        from: {
          emailAddress: 'sender@example.com',
          displayName: 'Test Sender'
        },
        dateTimeCreated: new Date('2026-01-06T12:00:00Z'),
        to: [
          {
            emailAddress: 'recipient@example.com',
            displayName: 'Test Recipient'
          }
        ],
        body: {
          getAsync: jest.fn((_coercionType, callback) => {
            callback({
              status: 'succeeded',
              value: 'Test email body content'
            });
          })
        }
      }
    }
  },
  HostType: {
    Outlook: 'Outlook'
  },
  MailboxEnums: {
    ItemType: {
      Message: 'message',
      Appointment: 'appointment'
    }
  },
  AsyncResultStatus: {
    Succeeded: 'succeeded',
    Failed: 'failed'
  },
  CoercionType: {
    Text: 'text',
    Html: 'html'
  }
};

// Mock console methods to reduce test output noise
global.console = {
  ...console,
  log: jest.fn(),
  warn: jest.fn(),
  error: jest.fn()
};

# Email Retrieval Script with Microsoft Graph API

This script allows users to fetch emails from specific Microsoft Exchange folders based on a subject filter, and optionally, a body filter. It employs the Microsoft Graph API, leveraging the client credentials flow with a JWT bearer assertion.

## Prerequisites

- **Node.js**: Ensure you have Node.js installed. If not, download and install it from [Node.js official website](https://nodejs.org/).

## Setup

### 1. Azure App Registration

To use this script, you need to have an application registered on Azure.

- Navigate to the [Azure Portal](https://portal.azure.com/)
- Go to **Azure Active Directory** > **App Registrations** > **New Registration**.
- Name your application and ensure the **Redirect URI** is set to `Web` and points to `http://localhost`.
- Once registered, note down the `Application (client) ID` (this will replace `clientId` in the code).
- Under **API permissions**, add `Mail.Read` permissions under Microsoft Graph.

### 2. Create a Certificate

For the JWT bearer assertion, you'll need a certificate.

- You can generate a self-signed certificate using OpenSSL:
  ```bash
  openssl req -x509 -sha256 -nodes -days 365 -newkey rsa:2048 -keyout key.pem -out cert.pem
  ```

- This command generates both the private key (`key.pem`) and the public certificate (`cert.pem`).

### 3. Upload Certificate to Azure

- In your Azure App Registration, go to **Certificates & Secrets**.
- Click on **Upload certificate** and upload `cert.pem`.
- After uploading, Azure will show a thumbprint for the certificate. Note it down (this will replace `thumbprintHex` in the code).

### 4. Configure the Script

In both `token.js` and `getEmail.js`:

- Replace placeholders like `xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx` with your actual `clientId` and `tenantId`.
- Update the path to your private key in `token.js` if necessary.
- Place the noted thumbprint from Azure in the `thumbprintHex` placeholder in `token.js`.

## Execution

Before running the script, you need to install the required Node.js modules:

```bash
npm install axios jsonwebtoken
```

To execute the script:

1. Navigate to the directory containing `getEmail.js`.
2. Run the script:
   ```bash
   node getEmail.js
   ```

## Customization

- You can update the `getEmail` function call at the end of `getEmail.js` to specify different user emails, subject filters, and body filters.
- The folders from which emails are fetched are currently hardcoded (`['Inbox', 'Deleted Items', 'Sent Items', 'Archive']`). Modify this array in `getEmail.js` if needed.

## Troubleshooting

- If you encounter `invalid_client` errors, ensure that your certificate, clientId, and tenantId are correctly set up.
- Ensure you've provided the necessary API permissions in your Azure App Registration.

## Contributing

Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change.

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.
# Email Attachment Fetcher with Node.js

## Overview

This Node.js application provides a flexible solution for automating the retrieval of email attachments from an IMAP server. The code utilizes the `imap`, `fs`, and `mailparser` libraries to connect to an email account, fetch emails, parse them, and save supported attachments locally.

## Features

- **IMAP Connectivity:** Establishes a connection to the IMAP server to interact with emails.
- **Attachment Parsing:** Parses email bodies and identifies supported attachments.
- **Local Storage:** Saves supported attachments to a specified local folder.
- **File Type Support:** Currently supports attachments with extensions such as .pdf, .xlsx, .jpg, .zip, .rar, and .docx.

## Prerequisites

Ensure you have the following installed:

- [Node.js](https://nodejs.org/)
- [npm](https://www.npmjs.com/) (Node.js package manager)

## Setup

1. Clone or download the repository to your local machine.
2. Navigate to the project directory using the terminal.

```bash
cd path/to/email-attachment-fetcher
```

3. Install the required dependencies.

```bash
npm install
```

4. Create a `.env` file in the project's root directory and provide your email credentials and IMAP server details.

```env
EMAIL=your_email@example.com
PASS=your_email_password
HOST=imap.your-email-provider.com
```

## Usage

Modify the `emailConfig` object in the code with your specific email, password, and IMAP host details. Adjust the `localFolderPath` variable if you want to save attachments to a different local folder.

Run the application:

```bash
node fetcher.js
```

## Notes

- Ensure your email provider supports IMAP access.
- Supported file extensions can be customized in the `isSupportedAttachment` method.

## Contributors

- Valentin Lutchanka

Feel free to contribute by submitting issues or pull requests.

Happy coding!
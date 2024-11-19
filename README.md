# PST to Office365 Migrator
## This Migrator uses Microsoft Graph API

### Requirements
- NodeJS 18+
- App Registered in Azure Portal

### What does it do?
- Create Missing Folders, Contact Folders and Calendars
- Import Mails, Contacts and Events
- Checks if already exists before importing

### How it works
- Register App in Azure
- Create .env File in Root Directory:
```
TENANT_ID="FROM_AZURE_APP"
CLIENT_ID="FROM_AZURE_APP"
CLIENT_SECRET="FROM_AZURE_APP"
TARGET_MAILBOX="mailbox@example.com"
PST_FILE="mailbox.pst"
```
- Set needed App Permission as Application Permission in Azure Poral:
```
Calendars.ReadWrite
Calendars.ReadShared
Contacts.ReadWrite
Group.ReadWrite.All
Mail.ReadWrite
MailboxFolder.ReadWrite.All
MailboxSettings.ReadWrite
User.ReadWrite.All
```
- Install Dependencies `npm install`
- Run Script `npm start`

!!!
```
Currently the pst-extractor library has some missing exports, to fix this open the following file after running npm install.
- Open node_modules/pst-extractor/dist/index.d.ts
- Add following to end of file:
export { PSTAttachment } from './PSTAttachment.class';
export { PSTNodeInputStream } from './PSTNodeInputStream.class';
export { PSTContact } from './PSTContact.class';
export { PSTAppointment } from './PSTAppointment.class';
```

> We might compile a binary .exe File to use without NodeJs later.
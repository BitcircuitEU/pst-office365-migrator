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

!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

> Currently the pst-extractor library has some missing exports, to fix this open the following file after running npm install.
> - Open node_modules/pst-extractor/dist/index.d.ts
> - Add following to end of file:
```
export { PSTAttachment } from './PSTAttachment.class';
export { PSTNodeInputStream } from './PSTNodeInputStream.class';
export { PSTContact } from './PSTContact.class';
export { PSTAppointment } from './PSTAppointment.class';
```

> Also replace the getNextChild Function in PSTFolder.class.js in same directory as it has a bug which may not itterate through all items
```
getNextChild() {
    this.initEmailsTable();
    if (this.emailsTable) {
        if (this.currentEmailIndex >= this.contentCount) {
            //console.log(`[DEBUG] Reached end of content. Index: ${this.currentEmailIndex}, Count: ${this.contentCount}`);
            return null;
        }
        
        // get the emails from the rows in the main email table
        const rows = this.emailsTable.getItems(this.currentEmailIndex, 1);
        //console.log(`[DEBUG] Retrieved row for index ${this.currentEmailIndex}`);
        
        if (rows.length === 0) {
            //console.log(`[DEBUG] No rows retrieved for index ${this.currentEmailIndex}`);
            this.currentEmailIndex++;
            return this.getNextChild();
        }
        
        const emailRow = rows[0].get(0x67f2);
        if (!emailRow || emailRow.itemIndex === -1) {
            //console.log(`[DEBUG] Invalid email row for index ${this.currentEmailIndex}`);
            this.currentEmailIndex++;
            return this.getNextChild();
        }
        
        //console.log(`[DEBUG] Attempting to load child for index ${this.currentEmailIndex}`);
        try {
            const childDescriptor = this.pstFile.getDescriptorIndexNode(long_1.default.fromNumber(emailRow.entryValueReference));
            const child = PSTUtil_class_1.PSTUtil.detectAndLoadPSTObject(this.pstFile, childDescriptor);
            this.currentEmailIndex++;
            return child;
        } catch (err) {
            //console.error(`[ERROR] Failed to load child for index ${this.currentEmailIndex}:`, err);
            this.currentEmailIndex++;
            return this.getNextChild();
        }
    } else if (this.fallbackEmailsTable) {
        if (this.currentEmailIndex >= this.contentCount ||
            this.currentEmailIndex >= this.fallbackEmailsTable.length) {
            // no more!
            return null;
        }
        const childDescriptor = this.fallbackEmailsTable[this.currentEmailIndex];
        const child = PSTUtil_class_1.PSTUtil.detectAndLoadPSTObject(this.pstFile, childDescriptor);
        this.currentEmailIndex++;
        return child;
    }
    return null;
}
```

> We might compile a binary .exe File to use without NodeJs later.

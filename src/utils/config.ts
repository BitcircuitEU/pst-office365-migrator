import path from 'path';
import dotenv from 'dotenv';

const envPath = path.resolve(process.cwd(), '.env');
dotenv.config({ path: envPath });


export const config = {
  tenantId: process.env.TENANT_ID || '',
  clientId: process.env.CLIENT_ID || '',
  clientSecret: process.env.CLIENT_SECRET || '',
  targetMailbox: process.env.TARGET_MAILBOX || '',
  pstFile: process.env.PST_FILE || '',
  supportedItemTypes: [
    'IPM.Note',
    'IPM.Note.Draft',
    'IPM.Note.SMIME',
    'IPM.Note.SMIME.MultipartSigned',
    'IPM.Appointment',
    'IPM.Contact',
  ],
  supportedFolderTypes: [
    'IPF.Note',
    'IPF.Appointment',
    'IPF.Contact',
  ],
  skipFolders: [
    'Search Root',
    'SPAM Search Folder',
    'SPAM Search Folder 2',
    'Deleted Items',
    'Conversation Action Settings',
    'Dateien',
    'Files',
    'Einstellungen für QuickSteps',
    'Einstellungen für Unterhaltungsaktionen',
    'ExternalContacts',
    'Journal',
    'GAL Contacts',
    'Recipient Cache',
    'Notizen',
    'Notes',
    'Postausgang',
    'Outbox',
    'RSS-Feeds',
    'Yammer-Stamm',
    'Recoverable Items',
    'Organizational Contacts',
    'PeopleCentricConversation Buddies',
    'RSS-Abonnements',
    'Synchronisierungsprobleme'
  ]
};
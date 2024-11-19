import { PSTFolder, PSTMessage, PSTContact, PSTAppointment, PSTAttachment, PSTNodeInputStream } from 'pst-extractor';
import { v4 as uuidv4 } from 'uuid';

import {config} from '../utils/config'
import {client} from '../utils/graphClient'

function encodeForGraphFilter(value: string): string {
  // Encode only specific characters that need encoding in OData queries
  return value.replace(/['"()%#[*+/|,;=:!@&]/g, (char) => {
    return encodeURIComponent(char);
  }).replace(/'/g, "''");
}

export interface O365MailFolder {
  name: string;
  id: string;
  parentFolderId?: string;
  children?: O365MailFolder[];
}

export interface O365ContactFolder {
    name: string;
    id: string;
    parentFolderId?: string;
    children?: O365ContactFolder[];
}

export interface O365CalendarFolder {
    name: string;
    id: string;
    isDefaultCalendar: boolean;
}

export async function getMailFolderIds(): Promise<O365MailFolder[]> {
  async function getAllFolders(parentId?: string): Promise<O365MailFolder[]> {
    const endpoint = parentId
      ? `https://graph.microsoft.com/v1.0/users/${config.targetMailbox}/mailFolders/${parentId}/childFolders`
      : `https://graph.microsoft.com/v1.0/users/${config.targetMailbox}/mailFolders`;
    
    let allFolders: O365MailFolder[] = [];
    let nextLink = endpoint;

    while (nextLink) {
      const response = await client.api(nextLink).get();
      
      const folders: O365MailFolder[] = response.value.map((folder: any) => ({
        name: folder.displayName,
        id: folder.id,
        parentFolderId: folder.parentFolderId || 'root'
      }));

      allFolders = allFolders.concat(folders);

      nextLink = response['@odata.nextLink'] || null;
    }

    for (const folder of allFolders) {
      const childFolders = await getAllFolders(folder.id);
      if (childFolders.length > 0) {
        folder.children = childFolders;
      }
    }

    return allFolders;
  }

  return await getAllFolders();
}
  
export function formatMailFolderStructure(folders: O365MailFolder[], depth: number = 0): string {
  let result = '';
  for (const folder of folders) {
    result += '  '.repeat(depth) + `${folder.name} (ID: ${folder.id})\n`;
    if (folder.children && folder.children.length > 0) {
      result += formatMailFolderStructure(folder.children, depth + 1);
    }
  }
  return result;
}

export async function getContactFolderIds(): Promise<O365ContactFolder[]> {
    async function getAllFolders(parentId?: string): Promise<O365ContactFolder[]> {
      const endpoint = parentId
        ? `https://graph.microsoft.com/v1.0/users/${config.targetMailbox}/contactFolders/${parentId}/childFolders`
        : `https://graph.microsoft.com/v1.0/users/${config.targetMailbox}/contactFolders`;
      
      const response = await client.api(endpoint).get();
      
      let folders: O365ContactFolder[] = response.value.map((folder: any) => ({
        name: encodeForGraphFilter(folder.displayName),
        id: folder.id,
        parentFolderId: folder.parentFolderId
      }));
  
      for (const folder of folders) {
        const childFolders = await getAllFolders(folder.id);
        folders = folders.concat(childFolders);
      }
  
      return folders;
    }
  
    return await getAllFolders();
}

export async function getCalendarIds(): Promise<O365CalendarFolder[]> {
    let endpoint = `https://graph.microsoft.com/v1.0/users/${config.targetMailbox}/calendars`;
    const response = await client.api(endpoint).get();

    let folders: O365CalendarFolder[] = response.value.map((folder: any) => ({
        name: encodeForGraphFilter(folder.name),
        id: folder.id,
        isDefaultCalendar: folder.isDefaultCalendar
    }));

    return folders;
}

export async function createMailFolder(folderName: string, parentFolderName: string): Promise<string> {
  const folders = await getMailFolderIds();
  let parentFolder = folders.find(f => f.name.toLowerCase() === parentFolderName.toLowerCase());

  let endpoint: string;

  if (!parentFolder) {
    console.log(`Parent folder "${parentFolderName}" not found. Creating in root.`);
    endpoint = `https://graph.microsoft.com/v1.0/users/${config.targetMailbox}/mailFolders`;
  } else {
    endpoint = `https://graph.microsoft.com/v1.0/users/${config.targetMailbox}/mailFolders/${parentFolder.id}/childFolders`;
  }
  
  try {
    const response = await client.api(endpoint).post({
      displayName: encodeForGraphFilter(folderName)
    });
    console.log(`Created mail folder "${folderName}" in "${parentFolderName || 'root'}"`);
    return response.id;
  } catch (error) {
    console.error(`Error creating mail folder "${folderName}":`, error);
    throw error;
  }
}

export async function createContactFolder(folderName: string, parentFolderName: string): Promise<string> {
  const folders = await getContactFolderIds();
  let parentFolder = folders.find(f => f.name.toLowerCase() === parentFolderName.toLowerCase());

  let endpoint: string;

  if (!parentFolder) {
    console.log(`Parent folder "${parentFolderName}" not found. Creating in root.`);
    endpoint = `https://graph.microsoft.com/v1.0/users/${config.targetMailbox}/contactFolders`;
  } else {
    endpoint = `https://graph.microsoft.com/v1.0/users/${config.targetMailbox}/contactFolders/${parentFolder.id}/childFolders`;
  }
  
  try {
    const response = await client.api(endpoint).post({
      displayName: encodeForGraphFilter(folderName)
    });
    console.log(`Created contact folder "${folderName}" in "${parentFolderName || 'root'}"`);
    return response.id;
  } catch (error) {
    console.error(`Error creating contact folder "${folderName}":`, error);
    throw error;
  }
}

export async function createCalendar(calendarName: string): Promise<string | null> {
  if (calendarName.toLowerCase() === "kalender" || calendarName.toLowerCase() === "calendar") {
    console.log(`Calendar "${calendarName}" is the default calendar and already exists.`);
    
    const calendars = await getCalendarIds();
    const defaultCalendar = calendars.find(c => c.isDefaultCalendar);
    
    if (defaultCalendar) {
      return defaultCalendar.id;
    } else {
      console.error("Default calendar not found. This should not happen.");
      return null;
    }
  }

  const endpoint = `https://graph.microsoft.com/v1.0/users/${config.targetMailbox}/calendars`;
  
  try {
    const response = await client.api(endpoint).post({
      name: encodeForGraphFilter(calendarName)
    });
    console.log(`Created calendar "${calendarName}"`);
    return response.id;
  } catch (error) {
    console.error(`Error creating calendar "${calendarName}":`, error);
    throw error;
  }
}

export async function createMissingFolders(pstFolders: any[]): Promise<void> {
  const o365MailFolders = await getMailFolderIds();
  const o365ContactFolders = await getContactFolderIds();
  const o365Calendars = await getCalendarIds();
  const mailFolderMap = new Map<string, O365MailFolder>();
  const contactFolderMap = new Map<string, O365ContactFolder>();
  const calendarMap = new Map<string, O365CalendarFolder>();

  // Statistics
  const stats = {
    total: { mail: 0, contact: 0, calendar: 0 },
    skipped: { mail: 0, contact: 0, calendar: 0 },
    created: { mail: 0, contact: 0, calendar: 0 },
    error: { mail: 0, contact: 0, calendar: 0 }
  };

  // Create maps of existing O365 folders and calendars
  function addFolderToMap(folder: O365MailFolder | O365ContactFolder, map: Map<string, any>) {
    map.set(folder.name.toLowerCase(), folder);
    if (folder.children) {
      folder.children.forEach(child => addFolderToMap(child, map));
    }
  }
  o365MailFolders.forEach(folder => addFolderToMap(folder, mailFolderMap));
  o365ContactFolders.forEach(folder => addFolderToMap(folder, contactFolderMap));
  o365Calendars.forEach(calendar => calendarMap.set(calendar.name.toLowerCase(), calendar));

  // Function to recursively create folders and calendars
  async function createFolder(folder: any): Promise<void> {
    if (folder.class !== 'IPF.Note' && folder.class !== 'IPF.Contact' && folder.class !== 'IPF.Appointment') return;
  
    const folderKey = folder.name.toLowerCase();
    const isContactFolder = folder.class === 'IPF.Contact';
    const isCalendar = folder.class === 'IPF.Appointment';
    const folderMap = isContactFolder ? contactFolderMap : (isCalendar ? calendarMap : mailFolderMap);
  
    // Update total count
    if (isCalendar) {
      stats.total.calendar++;
    } else if (isContactFolder) {
      stats.total.contact++;
    } else {
      stats.total.mail++;
    }

    // Skip creation for default "Contacts", "Kontakte", "Calendar", or "Kalender"
    if ((isContactFolder && (folderKey === 'contacts' || folderKey === 'kontakte')) ||
        (isCalendar && (folderKey === 'calendar' || folderKey === 'kalender'))) {
      if (isCalendar) {
        stats.skipped.calendar++;
      } else {
        stats.skipped.contact++;
      }
      return;
    }
  
    if (!folderMap.has(folderKey)) {
      try {
        if (isCalendar) {
          const newCalendar = await createCalendar(folder.name);
          if (newCalendar) {
            calendarMap.set(folderKey, {
              name: folder.name,
              id: newCalendar,
              isDefaultCalendar: false
            });
            stats.created.calendar++;
          }
        } else {
          let parentFolderId = '';
          let endpoint = isContactFolder
            ? `https://graph.microsoft.com/v1.0/users/${config.targetMailbox}/contactFolders`
            : `https://graph.microsoft.com/v1.0/users/${config.targetMailbox}/mailFolders`;
          
          if (folder.parentName) {
            const parentFolder = folderMap.get(folder.parentName.toLowerCase());
            if (parentFolder) {
              parentFolderId = parentFolder.id;
              endpoint = isContactFolder
                ? `https://graph.microsoft.com/v1.0/users/${config.targetMailbox}/contactFolders/${parentFolderId}/childFolders`
                : `https://graph.microsoft.com/v1.0/users/${config.targetMailbox}/mailFolders/${parentFolderId}/childFolders`;
            }
          }
  
          // Use $filter to check if the folder already exists
          const existingFoldersResponse = await client.api(endpoint)
            .filter(`displayName eq '${encodeForGraphFilter(folder.name)}'`)
            .get();
  
          let newFolder;
          if (existingFoldersResponse.value.length === 0) {
            // Folder doesn't exist, create it
            newFolder = await client.api(endpoint).post({
              displayName: encodeForGraphFilter(folder.name)
            });
            if (isContactFolder) {
              stats.created.contact++;
            } else {
              stats.created.mail++;
            }
          } else {
            // Folder already exists, use the existing one
            newFolder = existingFoldersResponse.value[0];
            if (isContactFolder) {
              stats.skipped.contact++;
            } else {
              stats.skipped.mail++;
            }
          }
  
          if (isContactFolder) {
            (folderMap as Map<string, O365ContactFolder>).set(folderKey, {
              name: folder.name,
              id: newFolder.id,
              parentFolderId: parentFolderId
            });
          } else {
            (folderMap as Map<string, O365MailFolder>).set(folderKey, {
              name: folder.name,
              id: newFolder.id,
              parentFolderId: parentFolderId,
              children: []
            });
          }
        }
      } catch (error) {
        if (isCalendar) {
          stats.error.calendar++;
        } else if (isContactFolder) {
          stats.error.contact++;
        } else {
          stats.error.mail++;
        }
      }
    } else {
      if (isCalendar) {
        stats.skipped.calendar++;
      } else if (isContactFolder) {
        stats.skipped.contact++;
      } else {
        stats.skipped.mail++;
      }
    }
  
    // Recursively create child folders if any (not applicable for calendars)
    if (!isCalendar && folder.children && Array.isArray(folder.children)) {
      for (const childFolder of folder.children) {
        await createFolder(childFolder);
      }
    }
  }

  // Start creating folders and calendars from the top level
  for (const folder of pstFolders) {
    await createFolder(folder);
  }

  // Log the final statistics
  console.log("Folder Creation Statistics:");
  console.log("Total Folders in PST:");
  console.log(`  Mail Folders: ${stats.total.mail}`);
  console.log(`  Contact Folders: ${stats.total.contact}`);
  console.log(`  Calendar Folders: ${stats.total.calendar}`);
  console.log("Created Folders:");
  console.log(`  Mail Folders: ${stats.created.mail}`);
  console.log(`  Contact Folders: ${stats.created.contact}`);
  console.log(`  Calendar Folders: ${stats.created.calendar}`);
  console.log("Skipped Folders (Already Exist):");
  console.log(`  Mail Folders: ${stats.skipped.mail}`);
  console.log(`  Contact Folders: ${stats.skipped.contact}`);
  console.log(`  Calendar Folders: ${stats.skipped.calendar}`);
  console.log("Skipped Folders (Due to Errors):");
  console.log(`  Mail Folders: ${stats.error.mail}`);
  console.log(`  Contact Folders: ${stats.error.contact}`);
  console.log(`  Calendar Folders: ${stats.error.calendar}`);
}

export async function checkEventExists(
  startTime: Date | null,
  endTime: Date | null,
  subject: string,
  isAllDay: boolean,
  body: string,
  location: string,
  calendarId?: string,
  attachments?: PSTAttachment[]
): Promise<boolean> {
  try {
    const endpoint = calendarId 
      ? `/users/${config.targetMailbox}/calendars/${calendarId}/events`
      : `/users/${config.targetMailbox}/events`;

    let filter = `subject eq '${encodeForGraphFilter(subject)}'`;

    if (isAllDay && startTime) {
      // For all-day events, set the time to midnight UTC
      const startDate = new Date(startTime);
      startDate.setUTCHours(0, 0, 0, 0);
      const endDate = new Date(startDate);
      endDate.setDate(endDate.getDate() + 1);

      filter += ` and start/dateTime eq '${startDate.toISOString()}' and end/dateTime eq '${endDate.toISOString()}'`;
    } else if (startTime && endTime) {
      filter += ` and start/dateTime eq '${startTime.toISOString()}' and end/dateTime eq '${endTime.toISOString()}'`;
    } else {
      console.warn(`Unable to check for event "${subject}" due to missing start or end time.`);
      return false;
    }    

    if (!startTime || !endTime) {
      console.warn(`Unable to check for event "${subject}" due to missing start or end time.`);
      return false;
    }

    const response = await client
      .api(endpoint)
      .filter(filter)
      .get();

    if (response.value.length > 0) {
      console.log("[Calendar] -> Skipped Existing | " + subject);
      return true;
    } else {
      if (isAllDay && startTime) {
        const startDate = new Date(startTime);
        startDate.setUTCHours(0, 0, 0, 0);
        const endDate = new Date(startDate);
        endDate.setDate(endDate.getDate() + 1);
        console.log("[Calendar] -> Importing | " + subject);
        await createEvent(subject, startDate, endDate, isAllDay, body, location, calendarId, attachments);
      } else {
        console.log("[Calendar] -> Importing | " + subject);
        await createEvent(subject, startTime, endTime, isAllDay, body, location, calendarId, attachments);
      }
      return true;
    }
  } catch (error) {
    console.error('Error checking/creating event:', error);
    return false;
  }
}

export async function createEvent(
  subject: string,
  startTime: Date | null,
  endTime: Date | null,
  isAllDay: boolean,
  body: string,
  location: string,
  calendarId?: string,
  attachments?: PSTAttachment[]
): Promise<string> {
  try {
    const endpoint = calendarId 
      ? `/users/${config.targetMailbox}/calendars/${calendarId}/events`
      : `/users/${config.targetMailbox}/events`;

    const eventData: any = {
      subject: subject,
      body: {
        contentType: "HTML",
        content: body
      },
      location: {
        displayName: location
      },
      isAllDay: isAllDay
    };

    if (startTime && endTime) {
      eventData.start = {
        dateTime: startTime.toISOString(),
        timeZone: "UTC"
      };
      eventData.end = {
        dateTime: endTime.toISOString(),
        timeZone: "UTC"
      };
    } else {
      throw new Error("Invalid event data: start time or end time is missing");
    }

    // If there are attachments, set hasAttachments to true
    if (attachments && attachments.length > 0) {
      eventData.hasAttachments = true;
    }

    const response = await client.api(endpoint).post(eventData);
    const createdEventId = response.id;

    // If there are attachments, add them to the created event
    if (attachments && attachments.length > 0) {
      for (const attachment of attachments) {
        await addAttachmentToEvent(createdEventId, attachment);
      }
    }

    return createdEventId;
  } catch (error) {
    console.error('Error creating event:', error);
    throw error;
  }
}

async function addAttachmentToEvent(eventId: string, attachment: PSTAttachment): Promise<void> {
  try {
    const attachmentEndpoint = `/users/${config.targetMailbox}/events/${eventId}/attachments`;
    
    // Get the attachment content as a Buffer
    const attachmentStream = attachment.fileInputStream;
    if (!attachmentStream) {
      throw new Error('Attachment stream is null');
    }
    const attachmentBuffer = await streamToBuffer(attachmentStream);

    const attachmentData = {
      "@odata.type": "#microsoft.graph.fileAttachment",
      name: attachment.filename,
      contentBytes: attachmentBuffer.toString('base64')
    };

    await client.api(attachmentEndpoint).post(attachmentData);
  } catch (error) {
    console.error('Error adding attachment to event:', error);
    throw error;
  }
}

// Helper function to convert a stream to a buffer
async function streamToBuffer(stream: PSTNodeInputStream): Promise<Buffer> {
  const chunks: number[] = [];
  let chunk: number;
  
  while ((chunk = stream.read()) !== -1) {
    chunks.push(chunk);
  }

  return Buffer.from(new Uint8Array(chunks));
}

export async function checkContactExists(
  contact: PSTContact,
  contactFolderId?: string
): Promise<boolean> {
  try {
    let endpoint: string;

    // Use the provided contactFolderId if available
    if (contactFolderId) {
      endpoint = `/users/${config.targetMailbox}/contactFolders/${contactFolderId}/contacts`;
    } else {
      endpoint = `/users/${config.targetMailbox}/contacts`;
    }

    let filter: string;

    // Check if contact exists based on email if available, otherwise use display name
    if (contact.email1EmailAddress) {
      filter = `emailAddresses/any(e:e/address eq '${encodeForGraphFilter(contact.email1EmailAddress)}')`;
    } else if (contact.displayName) {
      filter = `displayName eq '${encodeForGraphFilter(contact.displayName)}'`;
    } else {
      console.log("[Contact] -> Importing (No email or display name) | Unknown");
      await createContact(contact, contactFolderId);
      return true;
    }

    const response = await client.api(endpoint).filter(filter).get();

    if (response.value.length > 0) {
      console.log("[Contact] -> Skipped Existing | " + (contact.displayName || contact.email1EmailAddress || "Unknown"));
      return true;
    } else {
      // Create the contact
      console.log("[Contact] -> Importing | " + (contact.displayName || contact.email1EmailAddress || "Unknown"));
      await createContact(contact, contactFolderId);
      return true;
    }
  } catch (error) {
    console.error('Error checking/creating contact:', error);
    return false;
  }
}

async function createContact(contact: PSTContact, contactFolderId?: string): Promise<void> {
  try {
    const endpoint = contactFolderId
      ? `/users/${config.targetMailbox}/contactFolders/${contactFolderId}/contacts`
      : `/users/${config.targetMailbox}/contacts`;

      const contactData: any = {
        givenName: contact.givenName || '',
        surname: contact.surname || '',
        displayName: contact.displayName || '',
        emailAddresses: [],
        phones: [],
        businessHomePage: contact.businessHomePage || '',
        personalNotes: contact.body || ''
      };
  
      // Add email address if it exists
      if (contact.email1EmailAddress) {
        contactData.emailAddresses.push({
          address: contact.email1EmailAddress,
          name: contact.email1DisplayName || contact.displayName || ''
        });
      }

    // Add phone numbers if they exist
    if (contact.homeTelephoneNumber) {
      contactData.phones.push({ number: contact.homeTelephoneNumber, type: "home" });
    }
    if (contact.businessTelephoneNumber) {
      contactData.phones.push({ number: contact.businessTelephoneNumber, type: "business" });
    }
    if (contact.mobileTelephoneNumber) {
      contactData.phones.push({ number: contact.mobileTelephoneNumber, type: "mobile" });
    }

    // Remove empty fields
    Object.keys(contactData).forEach(key => 
      (contactData[key] === '' || contactData[key].length === 0) && delete contactData[key]
    );

    await client.api(endpoint).post(contactData);
  } catch (error) {
    console.error('Error creating contact:', error);
    if (error instanceof Error) {
      console.error('Error message:', error.message);
      console.error('Error stack:', error.stack);
    }
    if (typeof error === 'object' && error !== null) {
      console.error('Error details:', JSON.stringify(error, null, 2));
    }
    throw error;
  }
}

export async function checkMailExists(
  message: PSTMessage,
  mailFolderId: string
): Promise<boolean> {
  try {
    const endpoint = `/users/${config.targetMailbox}/mailFolders/${mailFolderId}/messages`;

    // Construct a more specific filter
    let filter = `subject eq '${encodeForGraphFilter(message.subject)}'`;

    if (message.senderEmailAddress) {
      filter += ` and from/emailAddress/address eq '${encodeForGraphFilter(message.senderEmailAddress)}'`;
    }

    if (message.messageDeliveryTime) {
      // Use a time range of ±1 minute to account for potential small differences
      const deliveryTime = new Date(message.messageDeliveryTime);
      const minTime = new Date(deliveryTime.getTime() - 60000); // 1 minute before
      const maxTime = new Date(deliveryTime.getTime() + 60000); // 1 minute after

      filter += ` and receivedDateTime ge ${minTime.toISOString()} and receivedDateTime le ${maxTime.toISOString()}`;
    }

    const response = await client
      .api(endpoint)
      .filter(filter)
      .select('id,subject,receivedDateTime,from')
      .top(1) // We only need to know if at least one matching email exists
      .get();

    if (response.value.length > 0) {
      console.log("[Mail] -> Skipped Existing | " + message.subject);
      return true;
    } else {
      console.log("[Mail] -> Importing | " + message.subject);
      await createMail(message, mailFolderId);
      return true;
    }
  } catch (error) {
    console.error('Error checking mail:', error);
    // If there's an error checking, attempt to create the mail anyway
    try {
      console.log("[Mail] -> Importing (after check error) | " + message.subject);
      await createMail(message, mailFolderId);
      return true;
    } catch (createError) {
      console.error('Error creating mail after check error:', createError);
      return false;
    }
  }
}

export async function createMail(
  message: PSTMessage,
  folderId: string
): Promise<string> {
  try {
    const endpoint = `/users/${config.targetMailbox}/mailFolders/${folderId}/messages`;
    const messageContent = message.bodyHTML || message.body || message.bodyRTF;
    const isDraft = message.messageClass === 'IPM.Note.Draft';

    const emailData: any = {
      internetMessageId: message.internetMessageId || `<${uuidv4()}@brede-wulf.de>`,
      subject: message.subject || "(No subject)",
      body: {
          contentType: 'HTML',
          content: messageContent
      },
      toRecipients: message.displayTo ? message.displayTo.split(';').map(recipient => ({ emailAddress: { address: recipient.trim() } })) : [],
      ccRecipients: message.displayCC ? message.displayCC.split(';').map(recipient => ({ emailAddress: { address: recipient.trim() } })) : [],
      bccRecipients: message.displayBCC ? message.displayBCC.split(';').map(recipient => ({ emailAddress: { address: recipient.trim() } })) : [],
      sender: message.senderEmailAddress ? {
          emailAddress: {
              name: message.senderName,
              address: message.senderEmailAddress
          }
      } : undefined,
      from: message.senderEmailAddress ? {
          emailAddress: {
              name: message.senderName,
              address: message.senderEmailAddress
          }
      } : undefined,
      isDraft: isDraft,
      isRead: !isDraft,
      attachments: [],  
    };

    if (isDraft) {
      emailData.singleValueExtendedProperties = [
          {
              id: "SystemTime 0x0039",
              value: message.clientSubmitTime?.toISOString() || new Date().toISOString()
          },
          {
              id: "SystemTime 0x0E06",
              value: message.messageDeliveryTime?.toISOString() || new Date().toISOString()
          },
          {
              id: "SystemTime 0x3007",
              value: message.creationTime?.toISOString() || new Date().toISOString()
          },
          {
              id: "SystemTime 0x3008",
              value: message.modificationTime?.toISOString() || new Date().toISOString()
          }
      ];
    } else {
      emailData.singleValueExtendedProperties = [
          {
              id: "Integer 0x0E07",
              value: "1"
          },
          {
              id: "SystemTime 0x0039",
              value: message.clientSubmitTime?.toISOString() || new Date().toISOString()
          },
          {
              id: "SystemTime 0x0E06",
              value: message.messageDeliveryTime?.toISOString() || new Date().toISOString()
          },
          {
              id: "SystemTime 0x3007",
              value: message.creationTime?.toISOString() || new Date().toISOString()
          },
          {
              id: "SystemTime 0x3008",
              value: message.modificationTime?.toISOString() || new Date().toISOString()
          }
      ];
    }

    // Process attachments
    if (message.numberOfAttachments > 0) {
      for (let i = 0; i < message.numberOfAttachments; i++) {
          const attachment = message.getAttachment(i);
          if (attachment && attachment.filename) {
              const attachmentStream = attachment.fileInputStream;
              if (attachmentStream) {
                  const chunks: Buffer[] = [];
                  const bufferSize = 8176;
                  const buffer = Buffer.alloc(bufferSize);
                  let bytesRead;

                  do {
                      bytesRead = attachmentStream.read(buffer);
                      if (bytesRead > 0) {
                          chunks.push(Buffer.from(buffer.slice(0, bytesRead)));
                      }
                  } while (bytesRead === bufferSize);

                  const attachmentBuffer = Buffer.concat(chunks);

                  emailData.attachments.push({
                      "@odata.type": "#microsoft.graph.fileAttachment",
                      name: attachment.longFilename || attachment.filename,
                      contentType: attachment.mimeTag,
                      contentBytes: attachmentBuffer.toString('base64'),
                  });
              } else {
                  console.warn(`Failed to get stream for attachment: ${attachment.filename}`);
              }
          }
      }
    }

    const response = await client.api(endpoint).post(emailData);
    const createdMessageId = response.id;

    return createdMessageId;
  } catch (error) {
    console.error('Error creating mail:', error);
    throw error;
  }
}
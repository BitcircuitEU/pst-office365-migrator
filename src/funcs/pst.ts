import { config } from '../utils/config';
import { getCalendarIds, checkEventExists, checkMailExists, getContactFolderIds, checkContactExists, getMailFolderIds } from './graph';
import { 
  PSTFile, 
  PSTFolder, 
  PSTMessage,
  PSTAppointment,
  PSTContact,
} from 'pst-extractor';

let pstFile: PSTFile;

function findMailFolderId(folders: any[], targetName: string): string | undefined {
  for (const folder of folders) {
    if (folder.name.toLowerCase() === targetName.toLowerCase()) {
      return folder.id;
    }
    if (folder.children && folder.children.length > 0) {
      const childResult = findMailFolderId(folder.children, targetName);
      if (childResult) {
        return childResult;
      }
    }
  }
  return undefined;
}

export async function initializePSTFile(filePath: string) {
  pstFile = new PSTFile(filePath);
}

function findTopOfPersonalFolders(folder: PSTFolder, depth: number = 0): PSTFolder | null {
  if (folder.displayName === 'Top of Personal Folders') {
    return folder;
  }

  if (depth <= 1 && folder.hasSubfolders) {
    for (let childFolder of folder.getSubFolders()) {
      const found = findTopOfPersonalFolders(childFolder, depth + 1);
      if (found) return found;
    }
  }

  return null;
}

export async function getPSTFolders(): Promise<any[]> {
  function getAllFolders(folder: PSTFolder, depth: number = 0, parentName: string = "Root"): any[] {
    const result: any[] = [];

    // Only process folders within "Top of Personal Folders"
    if (depth > 0 || folder.displayName === 'Top of Personal Folders') {
      const folderType = folder.containerClass || "IPF.Note";
      
      // Include all folders except "Top of Personal Folders" itself
      if (depth > 0) {
        result.push({
          name: folder.displayName,
          class: folderType,
          depth: depth - 1, // Adjust depth to start from 0 within "Top of Personal Folders"
          parentName: parentName,
          shouldSkip: config.skipFolders.includes(folder.displayName) || 
                      !config.supportedFolderTypes.includes(folderType) ||
                      (folder.displayName.startsWith('{') && folder.displayName.endsWith('}'))
        });
      }

      if (folder.hasSubfolders) {
        for (let childFolder of folder.getSubFolders()) {
          result.push(...getAllFolders(childFolder, depth + 1, folder.displayName));
        }
      }
    }
    
    return result;
  } 

  const rootFolder = pstFile.getRootFolder();
  const topFolder = findTopOfPersonalFolders(rootFolder);

  if (topFolder) {
    return getAllFolders(topFolder);
  } else {
    console.warn("'Top of Personal Folders' not found. Returning an empty array.");
    return [];
  }
}

export async function loopItemsToImport(): Promise<void> {
  function shouldSkipFolder(folderName: string): boolean {
    return config.skipFolders.includes(folderName) ||
           (folderName.startsWith('{') && folderName.endsWith('}'));
  }

  async function processFolder(folder: PSTFolder, depth: number = 0): Promise<void> {
    if (shouldSkipFolder(folder.displayName)) {
      return;
    }
  
    console.log(`[INFO] Processing folder: ${folder.displayName} (Depth: ${depth})`);
    //console.log(`[DEBUG] Total items in folder: ${folder.contentCount}`);
  
    let calendarId: string | undefined;
    let contactFolderId: string | undefined;
    let mailFolderId: string | undefined;
  
    switch(folder.containerClass) {
      case "IPF.Appointment":
        if (folder.displayName.toLowerCase() !== 'kalender' && folder.displayName.toLowerCase() !== 'calendar') {
          const calendarIds = await getCalendarIds();
          calendarId = calendarIds.find(c => c.name === folder.displayName)?.id;
          console.log("[INFO] Mappen Folder to -> ", calendarId);
        }
        break;
      case "IPF.Contact":
        if (folder.displayName.toLowerCase() !== 'kontakte' && folder.displayName.toLowerCase() !== 'contacts') {
          const contactIds = await getContactFolderIds();
          contactFolderId = contactIds.find(c => c.name === folder.displayName)?.id;
          console.log("[INFO] Mappen Folder to -> ", contactFolderId);
        }
        break;
      default:
        const mailFolderIds = await getMailFolderIds();
        mailFolderId = findMailFolderId(mailFolderIds, folder.displayName);
        console.log("[INFO] Mappen Folder to -> ", mailFolderId);
        break;
    }
  
    let itemCount = 0;
    let maxIterations = folder.contentCount;
  
    //console.log("[DEBUG] Starting item iteration");
  
    // Reset the cursor to the beginning of the folder
    folder.moveChildCursorTo(0);
  
    for (let i = 0; i < maxIterations; i++) {
      //console.log(`[DEBUG] Attempting to get item ${i + 1}`);
      const item = folder.getNextChild();
      if (!item) {
        //console.log(`[DEBUG] No item retrieved at index ${i}`);
        break;
      }
  
      //console.log(`[DEBUG] Processing item ${i + 1}: ${item.descriptorNodeId} - ${item.messageClass}`);
  
      if (config.supportedItemTypes.includes(item.messageClass)) {
        itemCount++;
        switch(item.messageClass) {
          case "IPM.Contact":
            const contact = item as PSTContact;
            try {
              await checkContactExists(contact, contactFolderId);              
            } catch (error) {
              console.error(`[Contact] Error processing ${contact.displayName}:`, error);
            }
            break;
          case "IPM.Appointment":
            const appointment = item as PSTAppointment;
            const isAllDay = appointment.startTime && appointment.endTime
              ? appointment.startTime.getHours() === 0 && 
                appointment.startTime.getMinutes() === 0 &&
                appointment.endTime.getHours() === 0 &&
                appointment.endTime.getMinutes() === 0
              : false;
  
            await checkEventExists(
              appointment.startTime,
              appointment.endTime,
              appointment.subject,
              isAllDay,
              appointment.body,
              appointment.location,
              calendarId
            );
            break;
            default:
              if (item instanceof PSTMessage) {
                const message = item as PSTMessage;
                try {
                  if (mailFolderId) {
                    //console.log(`[DEBUG] Attempting to process mail: ${message.subject}`);
                    await checkMailExists(message, mailFolderId);
                    //console.log(`[DEBUG] Successfully processed mail: ${message.subject}`);
                  } else {
                    console.warn(`[Mail] No matching folder found for ${message.subject}`);
                  }
                } catch (error) {
                  console.error(`[Mail] Error processing ${message.subject}:`, error);
                }
              } else {
                console.log(`[WARN] Unexpected item type: ${item.messageClass}`);
              }
          }
        } else {
          //console.log(`[DEBUG] Skipping unsupported item type: ${item.messageClass}`);
        }
    
        // Move to the next item explicitly
        folder.moveChildCursorTo(i + 1);
    
        // Safety check to prevent infinite loop
        if (i >= folder.contentCount - 1) {
          //console.log(`[DEBUG] Reached end of folder contents`);
          break;
        }
      }
    
      console.log(`[INFO] Processed ${itemCount} items in folder: ${folder.displayName}`);
    
      // Process subfolders
      if (folder.hasSubfolders) {
        for (let childFolder of folder.getSubFolders()) {
          await processFolder(childFolder, depth + 1);
        }
      }
    }

  const rootFolder = pstFile.getRootFolder();
  const topFolder = findTopOfPersonalFolders(rootFolder) || rootFolder;

  await processFolder(topFolder);
}
import { config } from './utils/config'
import { createMissingFolders, getMailFolderIds, formatMailFolderStructure } from './funcs/graph'
import { initializePSTFile, getPSTFolders, loopItemsToImport } from './funcs/pst'

async function main() {
  try {
    await initializePSTFile(config.pstFile);
    const pstFolders = await getPSTFolders();
    const mailFolders = await getMailFolderIds();

    //const formattedStructure = formatMailFolderStructure(mailFolders);
    //console.log('Mail Folder Structure:');
    //console.log(formattedStructure);
    
    await createMissingFolders(pstFolders);
    await loopItemsToImport();
    
    console.log('PST-Migration abgeschlossen.');
  } catch (error) {
    console.error('Fehler während der PST-Migration:', error);
  }
}

main();
/**
 * Email folder utilities
 */
const { callGraphAPI } = require('../utils/graph-api');

/**
 * Cache of folder information to reduce API calls
 * Format: { userId: { folderName: { id, path } } }
 */
const folderCache = {};

/**
 * Well-known folder names and their endpoints
 */
const WELL_KNOWN_FOLDERS = {
  'inbox': 'me/mailFolders/inbox/messages',
  'drafts': 'me/mailFolders/drafts/messages',
  'sent': 'me/mailFolders/sentItems/messages',
  'deleted': 'me/mailFolders/deletedItems/messages',
  'junk': 'me/mailFolders/junkemail/messages',
  'archive': 'me/mailFolders/archive/messages'
};

/**
 * Well-known folder names accepted by Graph as folder IDs (e.g. as a move
 * destination), keyed by the friendly names people actually type.
 */
const WELL_KNOWN_FOLDER_IDS = {
  'inbox': 'inbox',
  'drafts': 'drafts',
  'sent': 'sentitems',
  'sent items': 'sentitems',
  'sentitems': 'sentitems',
  'deleted': 'deleteditems',
  'deleted items': 'deleteditems',
  'deleteditems': 'deleteditems',
  'trash': 'deleteditems',
  'junk': 'junkemail',
  'junk email': 'junkemail',
  'junkemail': 'junkemail',
  'archive': 'archive'
};

const MAX_FOLDER_DEPTH = 4;

/** Escape a value for use inside a single-quoted OData string literal */
function escapeODataString(value) {
  return String(value).replace(/'/g, "''");
}

/**
 * Resolve a folder name to its messages endpoint path.
 * An empty name means the inbox. A name that cannot be found is an error:
 * silently listing the inbox instead would make callers act on the wrong mail.
 * @param {string} accessToken - Access token
 * @param {string} folderName - Folder name to resolve
 * @returns {Promise<string>} - Resolved endpoint path
 * @throws {Error} - If the folder does not exist or lookup fails
 */
async function resolveFolderPath(accessToken, folderName) {
  if (!folderName) {
    return WELL_KNOWN_FOLDERS['inbox'];
  }

  const lowerFolderName = folderName.toLowerCase();
  if (WELL_KNOWN_FOLDERS[lowerFolderName]) {
    console.error(`Using well-known folder path for "${folderName}"`);
    return WELL_KNOWN_FOLDERS[lowerFolderName];
  }

  const folderId = await getFolderIdByName(accessToken, folderName, { throwOnError: true });
  if (!folderId) {
    throw new Error(`Folder "${folderName}" not found. Use list-folders to see the exact folder names.`);
  }
  const path = `me/mailFolders/${folderId}/messages`;
  console.error(`Resolved folder "${folderName}" to path: ${path}`);
  return path;
}

/**
 * Search a list of folders (and, recursively, their children) for a
 * case-insensitive display name match.
 */
async function findFolderRecursive(accessToken, folders, lowerName, depth) {
  const direct = folders.find(f => (f.displayName || '').toLowerCase() === lowerName);
  if (direct) return direct;
  if (depth >= MAX_FOLDER_DEPTH) return null;

  for (const folder of folders.filter(f => f.childFolderCount > 0)) {
    const children = await callGraphAPI(
      accessToken,
      'GET',
      `me/mailFolders/${folder.id}/childFolders`,
      null,
      { $top: 100, $select: 'id,displayName,childFolderCount' }
    );
    const match = await findFolderRecursive(accessToken, children.value || [], lowerName, depth + 1);
    if (match) return match;
  }
  return null;
}

/**
 * Get the ID of a mail folder by its name. Matches well-known names
 * (inbox, archive, junk, ...), then top-level folders exactly, then any
 * folder case-insensitively, including subfolders.
 * @param {string} accessToken - Access token
 * @param {string} folderName - Name of the folder to find
 * @param {object} [options]
 * @param {boolean} [options.throwOnError] - Rethrow API errors instead of returning null
 * @returns {Promise<string|null>} - Folder ID or null if not found
 */
async function getFolderIdByName(accessToken, folderName, options = {}) {
  const wellKnownId = WELL_KNOWN_FOLDER_IDS[String(folderName).toLowerCase()];
  if (wellKnownId) {
    return wellKnownId;
  }

  try {
    console.error(`Looking for folder with name "${folderName}"`);
    const response = await callGraphAPI(
      accessToken,
      'GET',
      'me/mailFolders',
      null,
      { $filter: `displayName eq '${escapeODataString(folderName)}'` }
    );

    if (response.value && response.value.length > 0) {
      console.error(`Found folder "${folderName}" with ID: ${response.value[0].id}`);
      return response.value[0].id;
    }

    console.error(`No exact match found for "${folderName}", trying case-insensitive search incl. subfolders`);
    const allFoldersResponse = await callGraphAPI(
      accessToken,
      'GET',
      'me/mailFolders',
      null,
      { $top: 100, $select: 'id,displayName,childFolderCount' }
    );

    const match = await findFolderRecursive(
      accessToken, allFoldersResponse.value || [], folderName.toLowerCase(), 1
    );
    if (match) {
      console.error(`Found case-insensitive match for "${folderName}" with ID: ${match.id}`);
      return match.id;
    }

    console.error(`No folder found matching "${folderName}"`);
    return null;
  } catch (error) {
    console.error(`Error finding folder "${folderName}": ${error.message}`);
    if (options.throwOnError) throw error;
    return null;
  }
}

/**
 * Get all mail folders
 * @param {string} accessToken - Access token
 * @returns {Promise<Array>} - Array of folder objects
 */
async function getAllFolders(accessToken) {
  try {
    // Get top-level folders
    const response = await callGraphAPI(
      accessToken,
      'GET',
      'me/mailFolders',
      null,
      { 
        $top: 100,
        $select: 'id,displayName,parentFolderId,childFolderCount,totalItemCount,unreadItemCount'
      }
    );
    
    if (!response.value) {
      return [];
    }
    
    // Get child folders for folders with children
    const foldersWithChildren = response.value.filter(f => f.childFolderCount > 0);
    
    const childFolderPromises = foldersWithChildren.map(async (folder) => {
      try {
        const childResponse = await callGraphAPI(
          accessToken,
          'GET',
          `me/mailFolders/${folder.id}/childFolders`,
          null,
          {
            $top: 100,
            $select: 'id,displayName,parentFolderId,childFolderCount,totalItemCount,unreadItemCount'
          }
        );
        
        return childResponse.value || [];
      } catch (error) {
        console.error(`Error getting child folders for "${folder.displayName}": ${error.message}`);
        return [];
      }
    });
    
    const childFolders = await Promise.all(childFolderPromises);
    
    // Combine top-level folders and all child folders
    return [...response.value, ...childFolders.flat()];
  } catch (error) {
    console.error(`Error getting all folders: ${error.message}`);
    return [];
  }
}

module.exports = {
  WELL_KNOWN_FOLDERS,
  WELL_KNOWN_FOLDER_IDS,
  escapeODataString,
  resolveFolderPath,
  getFolderIdByName,
  getAllFolders
};

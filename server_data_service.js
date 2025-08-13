/**
 * @file server_data_service.js
 * @description Service for fetching, processing, and caching all data required by the application,
 * primarily for the sidebar UI.
 */
const DataService = (() => {
  'use strict';

  // =================================================================
  // PRIVATE CLASSES
  // =================================================================

  /**
   * @class _DataCacheManager
   * @description Handles all caching operations to a dedicated Google Drive file.
   */
  class _DataCacheManager {
    constructor() {
      this.cacheFileId = PropertiesService.getScriptProperties().getProperty('cacheFileId');
    }

    getCachedData() {
      if (!this.cacheFileId) return null;
      try {
        const file = DriveApp.getFileById(this.cacheFileId);
        const content = file.getBlob().getDataAsString();
        if (content && content.trim()) {
          const cache = JSON.parse(content);
          // Simple validation: check for a key property
          if (cache && cache.items) {
            console.log('💾 Data loaded from cache.');
            return cache;
          }
        }
      } catch (e) {
        console.warn(`Cache file not found or invalid: ${e.message}. A new one will be created.`);
        this.cacheFileId = null; // Reset to allow creation of a new cache file
      }
      return null;
    }

    cacheData(data) {
      const json = JSON.stringify(data, null, 2);
      try {
        if (this.cacheFileId) {
          DriveApp.getFileById(this.cacheFileId).setContent(json);
        } else {
          const file = DriveApp.createFile('sidebar_cache.json', json, 'application/json');
          this.cacheFileId = file.getId();
          PropertiesService.getScriptProperties().setProperty('cacheFileId', this.cacheFileId);
        }
        console.log('💾 Cache updated successfully.');
      } catch (e) {
        console.error(`❌ Failed to save cache: ${e.message}`);
      }
    }

    clearCache() {
      if (!this.cacheFileId) {
        return { success: true, message: 'No cache file to clear.' };
      }
      try {
        DriveApp.getFileById(this.cacheFileId).setContent('{}');
        return { success: true, message: 'Cache cleared successfully.' };
      } catch (e) {
        return { success: false, message: `Failed to clear cache: ${e.message}` };
      }
    }
  }

  /**
   * @class _StreamlinedDataFetcher
   * @description Fetches and processes data from Google Sheets and Drive.
   */
  class _StreamlinedDataFetcher {
    constructor() {
      // Configuration is fetched on-demand to ensure it's always fresh
      this.sheetId = PropertiesService.getScriptProperties().getProperty('sheetData');
      this.promptsFolderId = PropertiesService.getScriptProperties().getProperty('promptFolderId');
      this.spreadsheet = null;
      this.promptsFolder = null;
    }

    _openSpreadsheet() {
      if (!this.spreadsheet && this.sheetId) {
        this.spreadsheet = SpreadsheetApp.openById(this.sheetId);
      }
      return this.spreadsheet;
    }

    _getPromptsFolder() {
      if (!this.promptsFolder && this.promptsFolderId) {
        this.promptsFolder = DriveApp.getFolderById(this.promptsFolderId);
      }
      return this.promptsFolder;
    }

    fetchFreshData() {
      if (!this._openSpreadsheet() || !this._getPromptsFolder()) {
        throw new Error("sheetData or promptFolderId is not configured in Script Properties.");
      }
      const sheets = this._extractSheetsData();
      const prompts = this._extractPromptsData();
      const items = this._extractItemsFromSheets(sheets);

      return { sheets, items, promptTitles: prompts.titles, allPrompts: prompts.all };
    }

    _extractSheetsData() {
      return this._openSpreadsheet().getSheets().map(sheet => {
        const name = sheet.getName();
        if (name.startsWith('_')) return null;
        return {
          sheetName: name,
          sheetButton: (name.match(/\[(.*?)\]/) || [])[1] || name,
          sheetTag: (name.match(/\{(.*?)\}/) || [])[1] || name.toLowerCase(),
          sheetStyle: (name.match(/\((.*?)\)/) || [])[1] || 'primary',
        };
      }).filter(Boolean); // Filter out nulls for sheets starting with _
    }

    _extractItemsFromSheets(sheets) {
      const allItems = [];
      sheets.forEach(sheetConfig => {
        const sheet = this._openSpreadsheet().getSheetByName(sheetConfig.sheetName);
        if (!sheet) return;
        const values = sheet.getDataRange().getValues();
        for (let i = 1; i < values.length; i++) {
          const row = values[i];
          if (row[0] && row[2]) { // Title and Text must exist
            allItems.push({
              itemSheet: sheetConfig.sheetTag,
              itemTitle: row[0].toString(),
              itemTag: row[1] ? row[1].toString() : '',
              itemText: row[2].toString(),
              itemStyle: row[3] ? row[3].toString() : '',
              requires: row[4] ? row[4].toString() : '',
            });
          }
        }
      });
      return allItems;
    }

    _extractPromptsData() {
      const folder = this._getPromptsFolder();
      if (!folder) return { titles: [], all: {} };

      const files = folder.getFilesByType(MimeType.PLAIN_TEXT);
      const titles = [];
      const all = {};

      while (files.hasNext()) {
        const file = files.next();
        const fileName = file.getName();
        const parts = fileName.replace(/\.txt$/i, '').split(',');
        if (parts.length >= 3) {
          const tabName = parts[0].trim();
          if (!titles.includes(tabName)) titles.push(tabName);
          all[tabName] = {
            fromHeadline2: parts[1].trim(),
            toHeadline2: parts[2].trim(),
            content: file.getBlob().getDataAsString('UTF-8').trim(),
          };
        }
      }
      return { titles, all };
    }
  }


  // =================================================================
  // PUBLIC API
  // =================================================================
  const publicApi = {};
  const cacheManager = new _DataCacheManager();

  /**
   * Main function to get all data for the sidebar, using cache if available.
   * @returns {string} A JSON string of the sidebar data.
   */
  publicApi.getSidebarData = () => {
    try {
      const cachedData = cacheManager.getCachedData();
      if (cachedData) {
        return JSON.stringify(cachedData);
      }

      console.log('🔄 Fetching fresh data for sidebar...');
      const fetcher = new _StreamlinedDataFetcher();
      const freshData = fetcher.fetchFreshData();
      cacheManager.cacheData(freshData);

      return JSON.stringify(freshData);

    } catch (error) {
      console.error(`❌ Fatal error in getSidebarData: ${error.message}`);
      return JSON.stringify({ error: `Fatal error getting sidebar data: ${error.message}` });
    }
  };

  /**
   * Clears the data cache.
   * @returns {Object} An object indicating the result of the operation.
   */
  publicApi.clearCache = () => {
    return cacheManager.clearCache();
  };

  /**
   * Gets all prompts and filters them by the fromHeadline2 field.
   * @param {string} fromHeadline2 - The headline to filter prompts by.
   * @returns {Array<Object>} An array of matching prompt objects.
   */
  publicApi.getFilteredPrompts = (fromHeadline2) => {
    try {
      const allData = JSON.parse(publicApi.getSidebarData());
      if (allData.error || !allData.allPrompts) {
        throw new Error(allData.error || "Prompt data is not available.");
      }

      const filtered = Object.keys(allData.allPrompts)
        .map(key => {
          const prompt = allData.allPrompts[key];
          return {
            key: key,
            title: key,
            fromHeadline2: prompt.fromHeadline2,
            toHeadline2: prompt.toHeadline2,
            content: prompt.content
          };
        })
        .filter(prompt => prompt.fromHeadline2 === fromHeadline2);

      return filtered;

    } catch(e) {
      console.error(`Error filtering prompts: ${e.message}`);
      return [];
    }
  };

  return publicApi;

})();

// Expose public functions to the global scope for Google Apps Script
function getSidebarData() {
  return DataService.getSidebarData();
}

function clearDataCache() {
  return DataService.clearCache();
}

function getFilteredPromptsFromServer(fromHeadline2) {
    return DataService.getFilteredPrompts(fromHeadline2);
}

/**
 * @file server_config.js
 * @description Service for managing and validating script configuration properties.
 */
const ConfigService = (() => {
  'use strict';

  const REQUIRED_PROPS = ['sheetData', 'docData', 'promptFolderId', 'imageFolderId'];

  // =================================================================
  // PRIVATE HELPER FUNCTIONS
  // =================================================================

  /**
   * Validates that all required script properties are set and not empty.
   * @private
   * @returns {boolean} True if all properties are valid.
   */
  function _validateProperties() {
    const properties = PropertiesService.getScriptProperties();
    return REQUIRED_PROPS.every(prop => {
      const value = properties.getProperty(prop);
      return value && value.trim() !== '' && !value.includes('YOUR_');
    });
  }

  // =================================================================
  // PUBLIC API
  // =================================================================
  const publicApi = {};

  /**
   * Validates script properties, accessible globally.
   */
  publicApi.validate = () => {
    return _validateProperties();
  };

  /**
   * Sets up placeholder values for any missing required script properties.
   */
  publicApi.setup = () => {
    const properties = PropertiesService.getScriptProperties();
    REQUIRED_PROPS.forEach(prop => {
      if (!properties.getProperty(prop)) {
        properties.setProperty(prop, `YOUR_${prop.toUpperCase()}_HERE`);
      }
    });
    console.log('Script properties checked and defaults set if necessary.');
  };

  /**
   * Retrieves a specific configuration property.
   * @param {string} key - The key of the property to retrieve.
   * @returns {string|null} The property value or null if not found.
   */
  publicApi.get = (key) => {
    return PropertiesService.getScriptProperties().getProperty(key);
  };

  /**
   * Initializes the server configuration by validating and setting up properties.
   * This is the main entry point for configuration.
   * @returns {{success: boolean, error?: string}}
   */
  publicApi.initialize = () => {
    console.log('📋 Validating configuration...');
    if (!_validateProperties()) {
      console.log('⚙️ One or more properties are missing. Setting up defaults...');
      publicApi.setup();

      // Re-validate after setup
      if (!_validateProperties()) {
        const message = 'Failed to configure Script Properties automatically. Please set them manually.';
        DocumentApp.getUi().alert(`❌ ${message}`);
        return { success: false, error: message };
      }
    }
    console.log('✅ Configuration validated.');
    return { success: true };
  };

  return publicApi;

})();


// =================================================================
// GLOBAL STUBS for other server files
// =================================================================

function validateScriptProperties() {
  return ConfigService.validate();
}

function setupScriptProperties() {
  return ConfigService.setup();
}

function getSheetId() {
    return ConfigService.get('sheetData');
}
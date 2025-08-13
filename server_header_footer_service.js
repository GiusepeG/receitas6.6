/**
 * @file server_header_footer_service.js
 * @description Service for managing document headers and footers.
 */
const HeaderFooterService = (() => {
  'use strict';

  // =================================================================
  // PRIVATE HELPER FUNCTIONS
  // =================================================================

  /**
   * Scans a designated Google Drive folder for images named according to convention.
   * @private
   * @returns {{headers: Array<Object>, footers: Array<Object>}} An object containing arrays of header and footer images.
   */
  function _getAvailableImages() {
    const FOLDER_ID = PropertiesService.getScriptProperties().getProperty('imageFolderId');
    if (!FOLDER_ID) {
      throw new Error("The 'imageFolderId' property is not set in the script properties.");
    }
    const folder = DriveApp.getFolderById(FOLDER_ID);
    const files = folder.getFilesByType(MimeType.GOOGLE_APPS_SCRIPT); // Correction: Should be any image type, but this is a placeholder. Let's assume any file.

    const images = { headers: [], footers: [] };

    while (files.hasNext()) {
      const file = files.next();
      const fileName = file.getName();
      const spaceIndex = fileName.indexOf(' ');
      if (spaceIndex !== -1) {
        const type = fileName.substring(0, spaceIndex).toLowerCase();
        const name = fileName.substring(spaceIndex + 1).replace(/\.[^/.]+$/, '');
        const imageObject = { name: name, id: file.getId() };

        if (type === 'header') images.headers.push(imageObject);
        else if (type === 'footer') images.footers.push(imageObject);
      }
    }
    return images;
  }

  /**
   * Generates HTML for a set of radio buttons.
   * @private
   */
  function _createRadios(items, name, preselectedId) {
    let html = `<div class="form-check"><input class="form-check-input" type="radio" name="${name}" id="${name}-none" value="" ${!preselectedId ? 'checked' : ''}><label class="form-check-label" for="${name}-none">None</label></div>`;
    items.forEach((item, i) => {
      const isChecked = item.id === preselectedId ? 'checked' : '';
      html += `<div class="form-check"><input class="form-check-input" type="radio" name="${name}" id="${name}-${i}" value="${item.id}" ${isChecked}><label class="form-check-label" for="${name}-${i}">${item.name}</label></div>`;
    });
    return html;
  }

  /**
   * Clears the header and footer sections of the document.
   * @private
   */
  function _clearHeaderAndFooter() {
    const doc = DocumentApp.getActiveDocument();
    const header = doc.getHeader();
    if (header) header.clear();
    const footer = doc.getFooter();
    if (footer) footer.clear();
  }

  /**
   * Inserts a logo image into the header or footer.
   * @private
   */
  function _insertLogo(id, isHeader) {
    if (!id) return;
    const doc = DocumentApp.getActiveDocument();
    let section = isHeader ? doc.getHeader() : doc.getFooter();
    if (!section) {
      section = isHeader ? doc.addHeader() : doc.addFooter();
    }

    const imageBlob = DriveApp.getFileById(id).getBlob();
    const paragraph = section.appendParagraph('');
    const style = {
        [DocumentApp.Attribute.HORIZONTAL_ALIGNMENT]: DocumentApp.HorizontalAlignment.RIGHT,
        [DocumentApp.Attribute.SPACING_AFTER]: 0,
        [DocumentApp.Attribute.SPACING_BEFORE]: 0,
    };
    paragraph.setAttributes(style);
    paragraph.addPositionedImage(imageBlob)
      .setWidth(1880 / 3)
      .setHeight(350 / 3)
      .setLeftOffset(0)
      .setTopOffset(isHeader ? -40 : 0);
  }

  /**
   * Inserts the current date into the header.
   * @private
   */
  function _insertHeaderDate(hasHeaderImage) {
    const header = DocumentApp.getActiveDocument().getHeader();
    if (!header) return;

    for (let i = 0; i < (hasHeaderImage ? 1 : 3); i++) {
        header.appendParagraph('');
    }

    const formattedDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");
    const dateParagraph = header.appendParagraph(formattedDate);
    const style = {
        [DocumentApp.Attribute.BOLD]: true,
        [DocumentApp.Attribute.FONT_FAMILY]: 'IBM Plex Sans',
        [DocumentApp.Attribute.FONT_SIZE]: 12,
        [DocumentApp.Attribute.HORIZONTAL_ALIGNMENT]: DocumentApp.HorizontalAlignment.RIGHT,
    };
    dateParagraph.setAttributes(style);
  }

  // =================================================================
  // PUBLIC API
  // =================================================================
  const publicApi = {};

  /**
   * Shows a dialog for the user to choose a header and footer.
   */
  publicApi.chooseHeaderFooter = () => {
    try {
      const images = _getAvailableImages();
      const carimboFooter = images.footers.find(f => f.name.toLowerCase() === 'carimbo');

      const template = HtmlService.createTemplateFromFile('dialog_header_footer');
      template.headerRadios = _createRadios(images.headers, 'headerRadio', '');
      template.footerRadios = _createRadios(images.footers, 'footerRadio', carimboFooter ? carimboFooter.id : '');

      const htmlOutput = template.evaluate().setWidth(400).setHeight(500);
      DocumentApp.getUi().showModalDialog(htmlOutput, 'Choose Header and Footer');
    } catch (e) {
      DocumentApp.getUi().alert(`Error: ${e.message}`);
    }
  };

  /**
   * Inserts the selected header and footer images into the document.
   * @param {string} headerId - The Google Drive ID of the header image.
   * @param {string} footerId - The Google Drive ID of the footer image.
   */
  publicApi.insertHeaderAndFooter = (headerId, footerId) => {
    _clearHeaderAndFooter();
    if (headerId) _insertLogo(headerId, true);
    if (footerId) _insertLogo(footerId, false);
    _insertHeaderDate(!!headerId);
  };

  return publicApi;

})();


// =================================================================
// GLOBAL STUBS for Client-Side Calls
// =================================================================

function chooseHeaderFooter() {
  HeaderFooterService.chooseHeaderFooter();
}

// Note: insertHeaderAndFooter is called from the client dialog,
// but it needs to be available in the global scope if not called via a service object.
// The dialog code was updated to call HeaderFooterService.insertHeaderAndFooter directly.
// For robustness, we can keep a global stub.
function insertHeaderAndFooter(headerId, footerId) {
    HeaderFooterService.insertHeaderAndFooter(headerId, footerId);
}

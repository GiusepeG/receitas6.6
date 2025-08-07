const HeaderFooterService = {
  getAvailableImages: function() {
    const FOLDER_ID = PropertiesService.getScriptProperties().getProperty('imageFolderId');
    const folder = DriveApp.getFolderById(FOLDER_ID);
    const files = folder.getFiles();
    const images = { headers: [], footers: [] };

    while (files.hasNext()) {
      const file = files.next();
      const fileName = file.getName();
      const fileId = file.getId();
      const spaceIndex = fileName.indexOf(' ');
      if (spaceIndex !== -1) {
        const type = fileName.substring(0, spaceIndex).toLowerCase();
        const name = fileName.substring(spaceIndex + 1).replace(/\.[^/.]+$/, '');
        const imageObject = { name: name, id: fileId };
        if (type === 'header') {
          images.headers.push(imageObject);
        } else if (type === 'footer') {
          images.footers.push(imageObject);
        }
      }
    }
    return images;
  },

  createRadios: function(items, name, preselectedId) {
    let html = `<div class="form-check">
      <input class="form-check-input" type="radio" name="${name}" id="${name}-none" value="" ${!preselectedId ? 'checked' : ''}>
      <label class="form-check-label" for="${name}-none">Nenhum</label>
    </div>`;
    for (let i = 0; i < items.length; i++) {
      const isChecked = items[i].id === preselectedId ? 'checked' : '';
      html += `<div class="form-check">
        <input class="form-check-input" type="radio" name="${name}" id="${name}-${i}" value="${items[i].id}" ${isChecked}>
        <label class="form-check-label" for="${name}-${i}">${items[i].name}</label>
      </div>`;
    }
    return html;
  },

  chooseHeaderFooter: function() {
    const images = this.getAvailableImages();
    const carimboFooter = images.footers.find(footer => footer.name.toLowerCase() === 'carimbo');
    const preselectedFooterId = carimboFooter ? carimboFooter.id : '';

    const htmlContent = `
      <!DOCTYPE html>
      <html>
      <head>
        <base target="_top">
        <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@4.6.2/dist/css/bootstrap.min.css" integrity="sha384-xOolHFLEh07PJGoPkLv1IbcEPTNtaed2xpHsD9ESMhqIYd0nLMwNLD69Npy4HI+N" crossorigin="anonymous">
      </head>
      <body>
        <div class="container">
          <h5>Escolha o Cabeçalho</h5>
          ${this.createRadios(images.headers, 'headerRadio', '')}
          <h5 class="mt-3">Escolha o Rodapé</h5>
          ${this.createRadios(images.footers, 'footerRadio', preselectedFooterId)}
          <button class="btn btn-primary btn-block" id="applyBtn" onclick="submitSelection();">Aplicar Seleção</button>
          <button class="btn btn-secondary btn-block" onclick="google.script.host.close();">Cancelar</button>
        </div>
        <script>
          function submitSelection() {
            document.getElementById('applyBtn').disabled = true;
            document.getElementById('applyBtn').textContent = 'Aplicando...';
            const selectedHeaderId = document.querySelector('input[name="headerRadio"]:checked').value;
            const selectedFooterId = document.querySelector('input[name="footerRadio"]:checked').value;
            google.script.run
              .withSuccessHandler(() => google.script.host.close())
              .withFailureHandler(error => {
                alert('Falha ao aplicar: ' + error.message);
                google.script.host.close();
              })
              .HeaderFooterService.insertHeaderAndFooter(selectedHeaderId, selectedFooterId);
          }
        </script>
      </body>
      </html>`;

    const htmlOutput = HtmlService.createHtmlOutput(htmlContent)
      .setWidth(400)
      .setHeight(500);
    DocumentApp.getUi().showModalDialog(htmlOutput, 'Escolha o cabeçalho e rodapé');
  },

  insertHeaderAndFooter: function(headerId, footerId) {
    this.clearHeaderAndFooter();
    if (headerId) this.insertHeaderLogo(headerId);
    if (footerId) this.insertFooterLogo(footerId);
    this.insertHeaderDate(!!headerId);
  },

  clearHeaderAndFooter: function() {
    const header = DocumentApp.getActiveDocument().getHeader();
    const footer = DocumentApp.getActiveDocument().getFooter();
    if (header) this.clearSection(header);
    if (footer) this.clearSection(footer);
  },

  clearSection: function(section) {
    const paragraphs = section.getParagraphs();
    for (let i = 0; i < paragraphs.length - 1; i++) {
      section.removeChild(paragraphs[i]);
    }
    this.clearParagraph(section.getParagraphs()[0]);
  },

  clearParagraph: function(paragraph) {
    const images = paragraph.getPositionedImages();
    for (let i = 0; i < images.length; i++) {
      paragraph.removePositionedImage(images[i].getId());
    }
    paragraph.clear();
  },

  insertLogo: function(id, isHeader) {
    if (!id) return;
    let section = isHeader ? DocumentApp.getActiveDocument().getHeader() : DocumentApp.getActiveDocument().getFooter();
    if (!section) {
      section = isHeader ? DocumentApp.getActiveDocument().addHeader() : DocumentApp.getActiveDocument().addFooter();
    }
    const image = DriveApp.getFileById(id).getBlob();
    const style = {
      [DocumentApp.Attribute.HORIZONTAL_ALIGNMENT]: DocumentApp.HorizontalAlignment.RIGHT,
      [DocumentApp.Attribute.FONT_FAMILY]: 'IBM Plex Sans',
      [DocumentApp.Attribute.FONT_SIZE]: 6,
      [DocumentApp.Attribute.BOLD]: true,
      [DocumentApp.Attribute.SPACING_AFTER]: 0,
      [DocumentApp.Attribute.SPACING_BEFORE]: 0,
      [DocumentApp.Attribute.LINE_SPACING]: 0,
      [DocumentApp.Attribute.INDENT_START]: 0,
      [DocumentApp.Attribute.INDENT_FIRST_LINE]: 0,
      [DocumentApp.Attribute.INDENT_END]: 0
    };
    const paragraph = section.appendParagraph('');
    paragraph.setAttributes(style);
    const positionedImage = paragraph.addPositionedImage(image).setWidth(1880/3).setHeight(350/3).setLeftOffset(0);
    if(isHeader) {
      positionedImage.setTopOffset(-40);
    } else {
      positionedImage.setTopOffset(0);
    }
  },

  insertHeaderLogo: function(id_header) {
    this.insertLogo(id_header, true);
  },

  insertFooterLogo: function(id_footer) {
    this.insertLogo(id_footer, false);
  },

  insertHeaderDate: function(hasHeaderImage) {
    let header = DocumentApp.getActiveDocument().getHeader();
    if (!header) {
      header = DocumentApp.getActiveDocument().addHeader();
    }
    let numberOfParagraphs = hasHeaderImage ? 1 : 3;
    for (let i = 0; i < numberOfParagraphs; i++) {
      header.appendParagraph('');
    }
    const formattedDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");
    const dateParagraph = header.appendParagraph(formattedDate);
    const style = {
      [DocumentApp.Attribute.HORIZONTAL_ALIGNMENT]: DocumentApp.HorizontalAlignment.RIGHT,
      [DocumentApp.Attribute.BOLD]: true,
      [DocumentApp.Attribute.FONT_FAMILY]: 'IBM Plex Sans',
      [DocumentApp.Attribute.FONT_SIZE]: 12,
      [DocumentApp.Attribute.SPACING_AFTER]: 0,
      [DocumentApp.Attribute.SPACING_BEFORE]: 0
    };
    dateParagraph.setAttributes(style);
  }
};

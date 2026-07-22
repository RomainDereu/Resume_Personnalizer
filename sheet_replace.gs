/**
 * Paste either:
 *   - the complete Google Docs/Slides URL, or
 *   - the file ID
 */
const SOURCE_FILE_URL_OR_ID =
  'PASTE_SOURCE_FILE_URL_OR_ID_HERE';

/**
 * Replaces the contents of the current Google Doc or Google Slides file
 * with the contents of SOURCE_FILE_URL_OR_ID.
 *
 * The source and target must be the same type:
 *   Google Doc    -> Google Doc
 *   Google Slides -> Google Slides
 */
function replaceCurrentFileFromSource() {
  const sourceId = extractFileId_(SOURCE_FILE_URL_OR_ID);
  const target = getCurrentBoundFile_();

  if (sourceId === target.id) {
    throw new Error('The source file and target file cannot be the same.');
  }

  const sourceMimeType = DriveApp.getFileById(sourceId).getMimeType();

  if (target.type === 'DOCUMENT') {
    if (sourceMimeType !== MimeType.GOOGLE_DOCS) {
      throw new Error(
        'The current file is a Google Doc, so the source must also be a Google Doc.'
      );
    }

    const sourceDocument = DocumentApp.openById(sourceId);
    replaceDocumentContent_(target.file, sourceDocument);
  } else if (target.type === 'PRESENTATION') {
    if (sourceMimeType !== MimeType.GOOGLE_SLIDES) {
      throw new Error(
        'The current file is a Google Slides presentation, so the source must also be a presentation.'
      );
    }

    const sourcePresentation = SlidesApp.openById(sourceId);
    replacePresentationContent_(target.file, sourcePresentation);
  }

  console.log('File content replaced successfully.');
}


/**
 * Determines whether the script is bound to a Google Doc or presentation.
 */
function getCurrentBoundFile_() {
  try {
    const document = DocumentApp.getActiveDocument();

    if (document) {
      return {
        type: 'DOCUMENT',
        id: document.getId(),
        file: document
      };
    }
  } catch (error) {
    // The script is not bound to a Google Doc.
  }

  try {
    const presentation = SlidesApp.getActivePresentation();

    if (presentation) {
      return {
        type: 'PRESENTATION',
        id: presentation.getId(),
        file: presentation
      };
    }
  } catch (error) {
    // The script is not bound to Google Slides.
  }

  throw new Error(
    'This script must be opened from Extensions > Apps Script inside a Google Doc or Google Slides presentation.'
  );
}


/**
 * Replaces the active tab's body in the target document with the body of
 * the first tab in the source document.
 */
function replaceDocumentContent_(targetDocument, sourceDocument) {
  const sourceTabs = sourceDocument.getTabs();

  if (sourceTabs.length === 0) {
    throw new Error('The source document does not contain any document tabs.');
  }

  const sourceBody = sourceTabs[0]
    .asDocumentTab()
    .getBody();

  const targetBody = targetDocument
    .getActiveTab()
    .asDocumentTab()
    .getBody();

  // Erase the existing body.
  targetBody.clear();

  let insertionIndex = 0;

  for (let i = 0; i < sourceBody.getNumChildren(); i++) {
    const sourceElement = sourceBody.getChild(i);
    const elementType = sourceElement.getType();

    switch (elementType) {
      case DocumentApp.ElementType.PARAGRAPH:
        targetBody.insertParagraph(
          insertionIndex++,
          sourceElement.asParagraph().copy()
        );
        break;

      case DocumentApp.ElementType.LIST_ITEM:
        targetBody.insertListItem(
          insertionIndex++,
          sourceElement.asListItem().copy()
        );
        break;

      case DocumentApp.ElementType.TABLE:
        targetBody.insertTable(
          insertionIndex++,
          sourceElement.asTable().copy()
        );
        break;

      case DocumentApp.ElementType.TABLE_OF_CONTENTS:
        /*
         * DocumentApp does not provide an insertion method for a live
         * table of contents. Preserve its visible text instead.
         */
        targetBody.insertParagraph(
          insertionIndex++,
          sourceElement.asTableOfContents().getText()
        );
        break;

      default:
        console.warn(
          `Skipped unsupported document element: ${elementType}`
        );
    }
  }

  /*
   * Body.clear() can leave an empty placeholder paragraph. Remove it when
   * copied content has been inserted.
   */
  if (
    insertionIndex > 0 &&
    targetBody.getNumChildren() > insertionIndex
  ) {
    const leftoverElement = targetBody.getChild(insertionIndex);

    if (
      leftoverElement.getType() === DocumentApp.ElementType.PARAGRAPH &&
      leftoverElement.asParagraph().getText() === ''
    ) {
      targetBody.removeChild(leftoverElement);
    }
  }

  targetDocument.saveAndClose();
}


/**
 * Copies all source slides first, then deletes all original target slides.
 *
 * Copying first prevents the presentation from temporarily having no slides.
 */
function replacePresentationContent_(
  targetPresentation,
  sourcePresentation
) {
  const sourceSlides = sourcePresentation.getSlides();

  if (sourceSlides.length === 0) {
    throw new Error('The source presentation contains no slides.');
  }

  // Keep references to the slides that must be erased.
  const originalTargetSlides = targetPresentation.getSlides();

  // Append copies of all source slides in their original order.
  sourceSlides.forEach(function (sourceSlide) {
    targetPresentation.appendSlide(sourceSlide);
  });

  // Delete the previous target content.
  originalTargetSlides.forEach(function (targetSlide) {
    targetSlide.remove();
  });

  targetPresentation.saveAndClose();
}


/**
 * Accepts either a complete Google file URL or a bare file ID.
 */
function extractFileId_(urlOrId) {
  const value = String(urlOrId || '').trim();

  if (!value || value === 'PASTE_SOURCE_FILE_URL_OR_ID_HERE') {
    throw new Error(
      'Set SOURCE_FILE_URL_OR_ID before running the function.'
    );
  }

  // Standard Google file URLs contain /d/FILE_ID/.
  const urlMatch = value.match(/\/d\/([a-zA-Z0-9_-]+)/);

  if (urlMatch) {
    return urlMatch[1];
  }

  // Accept a bare Drive file ID.
  if (/^[a-zA-Z0-9_-]{20,}$/.test(value)) {
    return value;
  }

  throw new Error('The source file URL or ID is not valid.');
}
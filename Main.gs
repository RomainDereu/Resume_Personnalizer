/**
 * Confirms that the Google account running this script
 * can open the target presentation by ID.
 */
function testTargetPresentationAccess() {
  const presentationUrl =
    'https://docs.google.com/presentation/d/1fZoeg628A8gaKly_Y4Yv0Pt8fjkbaUi8cDCjTgBKJMw/edit';

  const presentationId =
    extractGoogleFileId_(presentationUrl);

  console.log(
    'Extracted presentation ID: ' +
    presentationId
  );

  console.log(
    'Effective Google account: ' +
    Session.getEffectiveUser().getEmail()
  );

  const presentation =
    SlidesApp.openById(presentationId);

  console.log(
    'Opened presentation: ' +
    presentation.getName()
  );

  console.log(
    'Confirmed ID: ' +
    presentation.getId()
  );

  console.log(
    'Presentation URL: ' +
    presentation.getUrl()
  );

  presentation.saveAndClose();
}

const OPENAI_API_URL = 'https://api.openai.com/v1/responses';
const OPENAI_MODEL = 'gpt-5.6-luna';

const WEBHOOK_SECRET_PROPERTY = 'WEBHOOK_SECRET';
const MAX_CHANGES_PER_REQUEST = 50;


/**
 * Adds a menu to the Google Slides presentation.
 *
 * These menu commands are local tests.
 * Direct ChatGPT orders arrive through doPost().
 */
function onOpen() {
  SlidesApp.getUi()
    .createMenu('Job Application Assistant')
    .addItem(
      'Test OpenAI connection',
      'testOpenAIConnection'
    )
    .addSeparator()
    .addItem(
      'Local test: Los Angeles → New York in Slides',
      'testAIReplacementInPresentation'
    )
    .addItem(
      'Local test: Los Angeles → New York in Google Doc',
      'testAIReplacementInDocument'
    )
    .addToUi();
}


/**
 * Tests the OpenAI API connection.
 */
function testOpenAIConnection() {
  const apiKey = PropertiesService
    .getScriptProperties()
    .getProperty('OPENAI_API_KEY');

  if (!apiKey) {
    throw new Error(
      'OPENAI_API_KEY was not found in Apps Script project properties.'
    );
  }

  const payload = {
    model: OPENAI_MODEL,
    input: 'Reply with exactly: Connection successful',
    store: false,
    max_output_tokens: 30
  };

  const response = UrlFetchApp.fetch(OPENAI_API_URL, {
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization: 'Bearer ' + apiKey
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const statusCode = response.getResponseCode();
  const responseBody = response.getContentText();

  console.log('Status code:', statusCode);
  console.log('Response:', responseBody);

  if (statusCode < 200 || statusCode >= 300) {
    throw new Error(
      'OpenAI API error ' + statusCode + ':\n' + responseBody
    );
  }

  const data = JSON.parse(responseBody);
  const outputText = extractOutputText_(data);

  if (!outputText) {
    throw new Error(
      'The request succeeded, but no output text was found.\n' +
      responseBody
    );
  }

  try {
    SlidesApp.getUi().alert(
      'OpenAI result: ' + outputText
    );
  } catch (error) {
    console.log(
      'OpenAI result: ' + outputText
    );
  }

  return outputText;
}

/**
 * Extracts text from an OpenAI Responses API response.
 */
function extractOutputText_(data) {
  if (
    typeof data.output_text === 'string' &&
    data.output_text.trim()
  ) {
    return data.output_text.trim();
  }

  const textParts = [];

  for (const item of data.output || []) {
    for (const content of item.content || []) {
      if (
        content.type === 'output_text' &&
        typeof content.text === 'string'
      ) {
        textParts.push(content.text);
      }
    }
  }

  return textParts.join('\n').trim();
}




/* =====================================================
 * CHATGPT ACTION WEB ENDPOINT
 * =====================================================
 */


/**
 * Simple health-check endpoint.
 *
 * Opening the deployed web-app URL in a browser should
 * return a small JSON response showing that it is ready.
 */
function doGet() {
  return jsonResponse_({
    ok: true,
    service: 'Job Application Assistant',
    supportedFileTypes: [
      'slides',
      'docs'
    ]
  });
}


/**
 * Receives direct editing orders from a Custom GPT Action.
 *
 * Expected request:
 *
 * {
 *   "webhookSecret": "...",
 *   "fileType": "slides",
 *   "fileId": "Google file URL or ID",
 *   "changes": [
 *     {
 *       "before": "Funabashi",
 *       "after": "New York"
 *     }
 *   ]
 * }
 */
function doPost(e) {
  try {
    const route = String(
      e && e.pathInfo
        ? e.pathInfo
        : ''
    )
      .replace(/^\/+|\/+$/g, '')
      .toLowerCase();

    const request =
      parseJsonRequest_(e);

    authenticateActionRequest_(request);

    /*
     * LIVE READ ROUTE
     *
     * URL:
     * /exec/read
     */
    if (route === 'read') {
      const command =
        validateReadRequest_(request);

      let result;

      if (command.fileType === 'slides') {
        result =
          readSlidesContent_(
            command.fileId
          );
      } else {
        result =
          readDocumentContent_(
            command.fileId
          );
      }

      return jsonResponse_({
        ok: true,
        operation: 'read',
        fileType: command.fileType,
        fileId: command.fileId,
        name: result.name,
        content: result.content
      });
    }

    /*
     * EDIT ROUTE
     *
     * URL:
     * /exec
     */
    if (route !== '') {
      throw new Error(
        'Unknown endpoint: /' + route
      );
    }

    const command =
      validateActionRequest_(request);

    let result;

    if (command.fileType === 'slides') {
      result =
        applyActionChangesToSlides_(
          command.fileId,
          command.changes
        );
    } else {
      result =
        applyActionChangesToDocument_(
          command.fileId,
          command.changes
        );
    }

    return jsonResponse_({
      ok: true,
      operation: 'edit',
      fileType: command.fileType,
      fileId: command.fileId,
      totalReplacements:
        result.totalReplacements,
      results:
        result.results
    });
  } catch (error) {
    console.error(
      error && error.stack
        ? error.stack
        : String(error)
    );

    return jsonResponse_({
      ok: false,
      error:
        error && error.message
          ? error.message
          : String(error)
    });
  }
}


/**
 * Reads and parses the incoming JSON request body.
 */
function parseJsonRequest_(e) {
  if (
    !e ||
    !e.postData ||
    !e.postData.contents
  ) {
    throw new Error(
      'The request body was empty.'
    );
  }

  try {
    return JSON.parse(
      e.postData.contents
    );
  } catch (error) {
    throw new Error(
      'The request body was not valid JSON.'
    );
  }
}


/**
 * Confirms that the request contains the correct
 * shared secret.
 */
function authenticateActionRequest_(request) {
  const expectedSecret =
    PropertiesService
      .getScriptProperties()
      .getProperty(
        WEBHOOK_SECRET_PROPERTY
      );

  if (!expectedSecret) {
    throw new Error(
      'WEBHOOK_SECRET is not configured ' +
      'in Script Properties.'
    );
  }

  const receivedSecret = String(
    request.webhookSecret || ''
  );

  if (receivedSecret !== expectedSecret) {
    throw new Error(
      'Unauthorized request.'
    );
  }
}

/**
 * Validates a live file-reading request.
 */
function validateReadRequest_(request) {
  const fileType = String(
    request.fileType || ''
  )
    .trim()
    .toLowerCase();

  if (
    fileType !== 'slides' &&
    fileType !== 'docs'
  ) {
    throw new Error(
      'fileType must be either ' +
      '"slides" or "docs".'
    );
  }

  const fileId =
    extractGoogleFileId_(
      request.fileId
    );

  return {
    fileType: fileType,
    fileId: fileId
  };
}


/**
 * Validates and normalizes the incoming command.
 */
function validateActionRequest_(request) {
  const fileType = String(
    request.fileType || ''
  )
    .trim()
    .toLowerCase();

  if (
    fileType !== 'slides' &&
    fileType !== 'docs'
  ) {
    throw new Error(
      'fileType must be either ' +
      '"slides" or "docs".'
    );
  }

  const fileId = extractGoogleFileId_(
    request.fileId
  );

  if (!Array.isArray(request.changes)) {
    throw new Error(
      'changes must be an array.'
    );
  }

  if (request.changes.length === 0) {
    throw new Error(
      'At least one change is required.'
    );
  }

  if (
    request.changes.length >
    MAX_CHANGES_PER_REQUEST
  ) {
    throw new Error(
      'A maximum of ' +
      MAX_CHANGES_PER_REQUEST +
      ' changes is allowed per request.'
    );
  }

  const normalizedChanges =
    request.changes.map(
      function(change, index) {
        if (
          !change ||
          typeof change.before !== 'string' ||
          typeof change.after !== 'string'
        ) {
          throw new Error(
            'Change ' +
            (index + 1) +
            ' must contain before and after strings.'
          );
        }

        if (change.before.length === 0) {
          throw new Error(
            'Change ' +
            (index + 1) +
            ' has an empty before value.'
          );
        }

        if (change.after.length === 0) {
          throw new Error(
            'Change ' +
            (index + 1) +
            ' has an empty after value.'
          );
        }

        if (
          change.before ===
          change.after
        ) {
          throw new Error(
            'Change ' +
            (index + 1) +
            ' has identical before and after values.'
          );
        }

        return {
          before: change.before,
          after: change.after
        };
      }
    );

  return {
    fileType: fileType,
    fileId: fileId,
    changes: normalizedChanges
  };
}


/**
 * Opens a Google Slides presentation by ID
 * and applies all requested changes.
 */
function applyActionChangesToSlides_(
  fileId,
  changes
) {
  const presentation =
    SlidesApp.openById(fileId);

  const results = [];
  let totalReplacements = 0;

  try {
    changes.forEach(function(change) {
      const replacementCount =
        replaceTextInPresentation_(
          presentation,
          change.before,
          change.after
        );

      totalReplacements +=
        replacementCount;

      results.push({
        before: change.before,
        after: change.after,
        replacements:
          replacementCount
      });
    });
  } finally {
    presentation.saveAndClose();
  }

  return {
    totalReplacements:
      totalReplacements,
    results: results
  };
}


/**
 * Opens a Google Docs document by ID
 * and applies all requested changes.
 */
function applyActionChangesToDocument_(
  fileId,
  changes
) {
  const document =
    DocumentApp.openById(fileId);

  const results = [];
  let totalReplacements = 0;

  try {
    changes.forEach(function(change) {
      const replacementCount =
        replaceTextInDocument_(
          document,
          change.before,
          change.after
        );

      totalReplacements +=
        replacementCount;

      results.push({
        before: change.before,
        after: change.after,
        replacements:
          replacementCount
      });
    });
  } finally {
    document.saveAndClose();
  }

  return {
    totalReplacements:
      totalReplacements,
    results: results
  };
}


/**
 * Returns a JSON response to the caller.
 */
function jsonResponse_(data) {
  return ContentService
    .createTextOutput(
      JSON.stringify(data)
    )
    .setMimeType(
      ContentService.MimeType.JSON
    );
}

/* =====================================================
 * LIVE FILE READING
 * =====================================================
 */

/**
 * Reads the current text from a Google Slides file.
 */
function readSlidesContent_(fileId) {
  const presentation =
    SlidesApp.openById(fileId);

  try {
    const slides =
      presentation
        .getSlides()
        .map(function(slide, slideIndex) {
          const blocks = [];

          slide
            .getPageElements()
            .forEach(function(element) {
              collectSlidesTextBlocks_(
                element,
                blocks
              );
            });

          return {
            slideNumber: slideIndex + 1,
            slideId: slide.getObjectId(),
            textBlocks: blocks
          };
        });

    return {
      name: presentation.getName(),
      content: {
        slides: slides,

        /*
         * Convenient plain-text version for the GPT.
         */
        plainText: slides
          .map(function(slide) {
            const slideText =
              slide.textBlocks
                .map(function(block) {
                  return block.text;
                })
                .join('\n');

            return (
              '--- Slide ' +
              slide.slideNumber +
              ' ---\n' +
              slideText
            );
          })
          .join('\n\n')
      }
    };
  } finally {
    presentation.saveAndClose();
  }
}


/**
 * Recursively extracts text from Slides elements.
 *
 * Supports:
 * - Shapes and text boxes
 * - Tables
 * - Groups
 */
function collectSlidesTextBlocks_(
  element,
  blocks
) {
  const type =
    element.getPageElementType();

  if (
    type ===
    SlidesApp.PageElementType.SHAPE
  ) {
    const text =
      element
        .asShape()
        .getText()
        .asString();

    if (text.trim()) {
      blocks.push({
        elementId:
          element.getObjectId(),
        elementType: 'shape',
        text: text
      });
    }

    return;
  }

  if (
    type ===
    SlidesApp.PageElementType.TABLE
  ) {
    const table =
      element.asTable();

    for (
      let row = 0;
      row < table.getNumRows();
      row++
    ) {
      for (
        let column = 0;
        column < table.getNumColumns();
        column++
      ) {
        const textRange =
          table
            .getCell(row, column)
            .getText();

        if (!textRange) {
          continue;
        }

        const text =
          textRange.asString();

        if (text.trim()) {
          blocks.push({
            elementId:
              element.getObjectId(),
            elementType: 'tableCell',
            row: row,
            column: column,
            text: text
          });
        }
      }
    }

    return;
  }

  if (
    type ===
    SlidesApp.PageElementType.GROUP
  ) {
    element
      .asGroup()
      .getChildren()
      .forEach(function(child) {
        collectSlidesTextBlocks_(
          child,
          blocks
        );
      });
  }
}


/* =====================================================
 * GOOGLE SLIDES
 * =====================================================
 */

/**
 * Searches every slide in the presentation.
 *
 * Supports:
 * - Text boxes
 * - Shapes
 * - Placeholders
 * - Tables
 * - Grouped shapes
 */
function replaceTextInPresentation_(
  presentation,
  before,
  after
) {
  const pattern =
    escapeRegularExpression_(before);

  let count = 0;

  presentation
    .getSlides()
    .forEach(function(slide) {
      slide
        .getPageElements()
        .forEach(function(element) {
          count +=
            replaceTextInSlideElement_(
              element,
              pattern,
              after
            );
        });
    });

  return count;
}


/**
 * Handles an individual Slides page element.
 */
function replaceTextInSlideElement_(
  element,
  pattern,
  after
) {
  const type =
    element.getPageElementType();

  if (
    type ===
    SlidesApp.PageElementType.SHAPE
  ) {
    return replaceTextInSlidesRange_(
      element.asShape().getText(),
      pattern,
      after
    );
  }

  if (
    type ===
    SlidesApp.PageElementType.TABLE
  ) {
    const table =
      element.asTable();

    let count = 0;

    for (
      let row = 0;
      row < table.getNumRows();
      row++
    ) {
      for (
        let column = 0;
        column < table.getNumColumns();
        column++
      ) {
        const textRange =
          table
            .getCell(row, column)
            .getText();

        // Merged non-head cells may return null.
        if (textRange) {
          count += replaceTextInSlidesRange_(
            textRange,
            pattern,
            after
          );
        }
      }
    }

    return count;
  }

  if (
    type ===
    SlidesApp.PageElementType.GROUP
  ) {
    let count = 0;

    element
      .asGroup()
      .getChildren()
      .forEach(function(child) {
        count += replaceTextInSlideElement_(
          child,
          pattern,
          after
        );
      });

    return count;
  }

  return 0;
}


/**
 * Replaces matching text inside one Slides text range.
 *
 * Existing font family, size, bold, italic,
 * underline and other starting text properties remain.
 * Only the replacement font color is changed.
 */
function replaceTextInSlidesRange_(
  textRange,
  pattern,
  after
) {
  const matches =
    textRange.find(pattern);

  /*
   * Process backwards so changing the length of one
   * match does not disturb the positions of later matches.
   */
  for (
    let index = matches.length - 1;
    index >= 0;
    index--
  ) {
    const replacementRange =
      matches[index].setText(after);

    replacementRange
      .getTextStyle()
      .setForegroundColor('#ff0000');
  }

  return matches.length;
}


/* =====================================================
 * GOOGLE DOCS
 * =====================================================
 */

/**
 * Searches every tab, including nested tabs,
 * in a Google Docs document.
 */
function replaceTextInDocument_(
  document,
  before,
  after
) {
  let count = 0;

  document
    .getTabs()
    .forEach(function(tab) {
      count += replaceTextInDocumentTab_(
        tab,
        before,
        after
      );
    });

  return count;
}


/**
 * Searches one document tab and its nested child tabs.
 */
function replaceTextInDocumentTab_(
  tab,
  before,
  after
) {
  let count = 0;

  if (
    tab.getType() ===
    DocumentApp.TabType.DOCUMENT_TAB
  ) {
    const body =
      tab
        .asDocumentTab()
        .getBody();

    count += replaceTextInDocumentBody_(
      body,
      before,
      after
    );
  }

  tab
    .getChildTabs()
    .forEach(function(childTab) {
      count += replaceTextInDocumentTab_(
        childTab,
        before,
        after
      );
    });

  return count;
}


/**
 * Finds every occurrence before modifying the document.
 *
 * Replacements are then performed backwards, preserving
 * the original character attributes and changing only
 * the foreground color to red.
 */
function replaceTextInDocumentBody_(
  body,
  before,
  after
) {
  const pattern =
    escapeRegularExpression_(before);

  const matches = [];

  let found =
    body.findText(pattern);

  while (found) {
    matches.push({
      textElement:
        found.getElement().asText(),

      start:
        found.getStartOffset(),

      end:
        found.getEndOffsetInclusive()
    });

    found =
      body.findText(pattern, found);
  }

  for (
    let index = matches.length - 1;
    index >= 0;
    index--
  ) {
    const match =
      matches[index];

    const originalAttributes =
      match.textElement.getAttributes(
        match.start
      );

    match.textElement.deleteText(
      match.start,
      match.end
    );

    match.textElement.insertText(
      match.start,
      after
    );

    if (after.length > 0) {
      const newEnd =
        match.start +
        after.length -
        1;

      /*
       * Restore font, size, bold, italic, links
       * and other original attributes.
       */
      match.textElement.setAttributes(
        match.start,
        newEnd,
        originalAttributes
      );

      /*
       * Then change only the font color.
       */
      match.textElement.setForegroundColor(
        match.start,
        newEnd,
        '#ff0000'
      );
    }
  }

  return matches.length;
}


/* =====================================================
 * SHARED HELPERS
 * =====================================================
 */

/**
 * Escapes text because Slides find() and Docs findText()
 * interpret the search value as a regular expression.
 */
function escapeRegularExpression_(text) {
  return text.replace(
    /[.*+?^${}()|[\]\\]/g,
    '\\$&'
  );
}


/**
 * Accepts either a complete Google URL or a raw file ID.
 */
function extractGoogleFileId_(urlOrId) {
  const value =
    String(urlOrId || '').trim();

  const match =
    value.match(/[-\w]{25,}/);

  if (!match) {
    throw new Error(
      'A valid Google Slides or Google Docs ' +
      'URL or file ID was not provided.'
    );
  }
  return match[0];
}
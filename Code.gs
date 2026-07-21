const OPENAI_API_URL = 'https://api.openai.com/v1/responses';
const OPENAI_MODEL = 'gpt-5.6-luna';

/**
 * Adds a menu to the spreadsheet.
 */
/**
 * Adds a menu to the Google Slides presentation.
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
      'Replace Los Angeles → New York in presentation',
      'testAIReplacementInPresentation'
    )
    .addItem(
      'Replace Los Angeles → New York in Google Doc',
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

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  if (spreadsheet) {
    spreadsheet.toast(
      outputText,
      'OpenAI connection test',
      10
    );
  } else {
    console.log('OpenAI result: ' + outputText);
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






/**
 * End-to-end test for the current Google Slides presentation.
 */
function testAIReplacementInPresentation() {
  const change = getExampleChangeFromOpenAI_();

  validateExampleChange_(change);

  const presentation =
    SlidesApp.getActivePresentation();

  if (!presentation) {
    throw new Error(
      'No active Google Slides presentation was found. ' +
      'This Apps Script project must be opened through ' +
      'Extensions → Apps Script inside the presentation.'
    );
  }

  const count = replaceTextInPresentation_(
    presentation,
    change.before,
    change.after
  );

  if (count === 0) {
    throw new Error(
      '"' + change.before + '" was not found ' +
      'anywhere in the presentation.'
    );
  }

  SlidesApp.getUi().alert(
    'Replacement complete.\n\n' +
    'Changed "' + change.before + '" to "' +
    change.after + '" in ' + count +
    ' location(s).\n\n' +
    'The replacement text should be red.'
  );
}


/**
 * End-to-end test for a Google Docs document.
 *
 * The script asks for the Google Doc URL or file ID.
 */
function testAIReplacementInDocument() {
  const ui = SlidesApp.getUi();

  const promptResult = ui.prompt(
    'Select Google Doc',
    'Paste the Google Doc URL or file ID:',
    ui.ButtonSet.OK_CANCEL
  );

  if (
    promptResult.getSelectedButton() !==
    ui.Button.OK
  ) {
    return;
  }

  const documentId = extractGoogleFileId_(
    promptResult.getResponseText()
  );

  const change = getExampleChangeFromOpenAI_();

  validateExampleChange_(change);

  const document =
    DocumentApp.openById(documentId);

  const count = replaceTextInDocument_(
    document,
    change.before,
    change.after
  );

  document.saveAndClose();

  if (count === 0) {
    throw new Error(
      '"' + change.before + '" was not found ' +
      'anywhere in the Google Doc.'
    );
  }

  ui.alert(
    'Replacement complete.\n\n' +
    'Changed "' + change.before + '" to "' +
    change.after + '" in ' + count +
    ' location(s).\n\n' +
    'The replacement text should be red.'
  );
}


/**
 * Requests this exact example from OpenAI:
 *
 * {
 *   "before": "Los Angeles",
 *   "after": "New York"
 * }
 */
function getExampleChangeFromOpenAI_() {
  const apiKey = PropertiesService
    .getScriptProperties()
    .getProperty('OPENAI_API_KEY');

  if (!apiKey) {
    throw new Error(
      'OPENAI_API_KEY was not found in Apps Script ' +
      'project properties.'
    );
  }

  const payload = {
    model: OPENAI_MODEL,

    instructions: [
      'Return one before-and-after text replacement.',
      'The before value must be exactly: Los Angeles',
      'The after value must be exactly: New York',
      'Return data matching the supplied JSON schema.'
    ].join('\n'),

    input:
      'Provide the requested text replacement.',

    store: false,
    max_output_tokens: 100,

    text: {
      format: {
        type: 'json_schema',
        name: 'text_replacement',
        strict: true,

        schema: {
          type: 'object',
          additionalProperties: false,

          properties: {
            before: {
              type: 'string'
            },

            after: {
              type: 'string'
            }
          },

          required: [
            'before',
            'after'
          ]
        }
      }
    }
  };

  const response = UrlFetchApp.fetch(
    OPENAI_API_URL,
    {
      method: 'post',
      contentType: 'application/json',

      headers: {
        Authorization: 'Bearer ' + apiKey
      },

      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    }
  );

  const statusCode =
    response.getResponseCode();

  const responseBody =
    response.getContentText();

  console.log(
    'Status code: ' + statusCode
  );

  console.log(
    'Response: ' + responseBody
  );

  if (
    statusCode < 200 ||
    statusCode >= 300
  ) {
    throw new Error(
      'OpenAI API error ' +
      statusCode +
      ':\n' +
      responseBody
    );
  }

  const data =
    JSON.parse(responseBody);

  const outputText =
    extractOutputText_(data);

  if (!outputText) {
    throw new Error(
      'OpenAI returned no output text.\n' +
      responseBody
    );
  }

  try {
    return JSON.parse(outputText);
  } catch (error) {
    throw new Error(
      'OpenAI output was not valid JSON:\n' +
      outputText
    );
  }
}


/**
 * Safety check for this specific test.
 */
function validateExampleChange_(change) {
  if (
    !change ||
    change.before !== 'Los Angeles' ||
    change.after !== 'New York'
  ) {
    throw new Error(
      'OpenAI returned an unexpected change:\n' +
      JSON.stringify(change, null, 2)
    );
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
      'A valid Google Doc URL or file ID ' +
      'was not provided.'
    );
  }

  return match[0];
}
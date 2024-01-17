/**
 * @OnlyCurrentDoc  Limits the script to only accessing the current document.
 */

var DIALOG_TITLE = 'Modify Images';
var SIDEBAR_TITLE = 'ReadEasy';

/**
 * Adds a custom menu with items to show the sidebar
 *
 * @param {Object} e The event parameter for a simple onOpen trigger.
 */
function onOpen(e) {
  DocumentApp.getUi()
      .createAddonMenu()
      .addItem('Open ReadEasy', 'showSidebar')
      .addToUi();
}

/**
 * Runs when the add-on is installed; calls onOpen() to ensure menu creation and
 * any other initializion work is done immediately.
 *
 * @param {Object} e The event parameter for a simple onInstall trigger.
 */
function onInstall(e) {
  onOpen(e);
}

/**
 * Opens a sidebar. The sidebar structure is described in the Sidebar.html
 * project file.
 */
function showSidebar() {
  var ui = HtmlService.createTemplateFromFile('Sidebar')
      .evaluate()
      .setTitle(SIDEBAR_TITLE)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  DocumentApp.getUi().showSidebar(ui); // shows the sidebar
}

/**
 * Opens a dialog. The dialog structure is described in the Dialog.html
 * project file.
 * Passes in an array of current alt text of images
 */
function showDialogWithImages(imageAlt) {
  var html = HtmlService.createTemplateFromFile('Dialog');
  html.imageAlt = imageAlt;
  var htmlOutput = html.evaluate().setWidth(400).setHeight(300);
  DocumentApp.getUi().showModalDialog(htmlOutput, 'Add/Modify Image Alt Texts'); // shows the dialog box
}

/**
 * Helper function to give all the images a corresponding alt text
 */
function setAltText(altTextData) {
  var body = DocumentApp.getActiveDocument().getBody();
  var images = body.getImages();
  let altTexts = altTextData.map(item => item.altText);
  
  for (var i = 0; i < altTexts.length; i++) { // sets the alt data for each image to the corresponding alt text
    images[i].setAltDescription(altTexts[i])
    }
  }

/**
 *  Sets document font family to arial.and the font size to 12
 */
function setDocumentFontToArial() {
  var body = DocumentApp.getActiveDocument().getBody();
  var paragraphs = body.getParagraphs();

  paragraphs.forEach(eachPara => {// sets each paragraph to Arial
    eachPara.setFontFamily("Arial");

    // Check if the paragraph style is "Normal"
    if (eachPara.getHeading() === DocumentApp.ParagraphHeading.NORMAL) {
      eachPara.setFontSize(12);
    }
  });
}
/**
 *  Checks to see if the string is formatted in Title Case
 */
function isTitleCase(str) {
  // Regular expression for strict title case
  const strictTitleCaseRegex = /^[A-Z][a-z]+(?:\s[A-Z][a-z]+)*$/;
  if (strictTitleCaseRegex.test(str)) {
    return true;
  }

  // More flexible title case check
  const words = str.split(" ");
  for (const word of words) {
    if (word.length > 1 && !word.match(/^[A-Z][a-z]+$/)) {
      return false;
    }
  }

  // Handles common exceptions (e.g., articles, conjunctions)
  const exceptions = ["a", "an", "the", "and", "but", "or", "for", "nor", "so", "yet"];
  for (const word of words) {
    if (exceptions.indexOf(word.toLowerCase()) === -1 && !word.match(/^[A-Z][a-z]+$/)) {
      return false;
    }
  }

  // If all checks pass, it's likely title case
  return true;
}

/**
 * Replaces text formatted with italic or underline with bold text in a Google Document.
 */
function replaceWithBold() {
  var body = DocumentApp.getActiveDocument().getBody();
  var paragraphs = body.getParagraphs();

  paragraphs.forEach(function(paragraph) {
    var text = paragraph.editAsText();
    var textLength = text.getText().length;

    for (var i = 0; i < textLength; i++) { // loops through each char in the document
      var isItalic = text.isItalic(i);
      var isUnderline = text.isUnderline(i);
      var linkUrl = text.getLinkUrl(i);
      if(!linkUrl) { // makes sure the string is not part of a link
        if (isItalic || isUnderline) { // replaces the text with bold
          text.setBold(i, i, true);
          text.setItalic(i, i, false);
          text.setUnderline(i, i, false);
        }
      }
    }
  });
}

/**
 * Change font to have no backround with white text
 */
function setDocumentColorsToHighContrast() {
  var body = DocumentApp.getActiveDocument().getBody();
  var paragraphs = body.getParagraphs();

  paragraphs.forEach(function(paragraph) { // loops through each paragraph
    var text = paragraph.editAsText();
    var textLength = text.getText().length;

    for (var i = 0; i < textLength; i++) { // loops through each individual char
      var backgroundColor = text.getBackgroundColor(i);
      var foregroundColor = text.getForegroundColor(i);
      var linkUrl = text.getLinkUrl(i);

      // Check if the current text is not a hyperlink before changing its color
      if (foregroundColor !== "#000000" && !linkUrl) {
        text.setForegroundColor(i, i, '#000000');
      }

      if (backgroundColor !== null && backgroundColor !== "") {
        text.setBackgroundColor(i, i, null); // Set background to transparent
      }
    }
  });
}

/**
 * Automatically formats document headings based on certain criteria
 */
function applyHeadings() {
  var body = DocumentApp.getActiveDocument().getBody();
  var paragraphs = body.getParagraphs();
  var header1FontSize = 0;

  paragraphs.forEach(paragraph => { // Loops through each paragraph
    var text = paragraph.getText();
    var textStyle = paragraph.getAttributes();
    var fontSize = textStyle[DocumentApp.Attribute.FONT_SIZE];
    var points = 0;

    // Check criteria and assign points
    points += text.length < 30 ? 1 : 0;
    points += fontSize > 12 ? 1 : 0;
    points += text.endsWith(":") ? 1 : 0;
    points += isTitleCase(text) ? 1 : 0; 

    if (points >= 2) {
      if (fontSize >= header1FontSize) {
        paragraph.setHeading(DocumentApp.ParagraphHeading.HEADING1);
        header1FontSize = fontSize;
      } else {
        paragraph.setHeading(DocumentApp.ParagraphHeading.HEADING2);
      }
    }
  });
}


/**
 * Locates all images in the document and their descriptions
 */
function findImagesAndPrompt() {
  var body = DocumentApp.getActiveDocument().getBody();
  var allInlineImages = body.getImages(); // gets all the images in the doc
  var imageAlt = [];

  for (var i = 0; i < allInlineImages.length; i++) { // pushes all images to imageAlt
    imageAlt.push(allInlineImages[i].getAltDescription());
  }

  if (imageAlt.length > 0) {
    showDialogWithImages(imageAlt);
  }
}

/**
 *  Extract the URL and Text connected to that URL int an array
 */
function extractUrlTextAndLinkFromParagraph(paragraphText) {
  var urlData = [];
  var textLength = paragraphText.getText().length;

  for (var i = 0; i < textLength; i++) { // loops through each char in the paragraph
    var url = paragraphText.getLinkUrl(i);
    if (url) {
      // Find the end of the linked text
      var start = i;
      while (i < textLength && paragraphText.getLinkUrl(i) === url) {
        i++;
      }
      var end = i;
      var linkedText = paragraphText.getText().substring(start, end);

      urlData.push({ linkedText: linkedText, url: url });
    }
  }

  return urlData;
}

/**
 * Links text in a paragraph with the specified URL
 */
function linkTextToUrl(paragraphText, linkedText, url) {
  var textString = paragraphText.getText();
  var startIndex = textString.indexOf(linkedText);

  if (startIndex !== -1) {
    var endIndex = startIndex + linkedText.length - 1;
    paragraphText.setLinkUrl(startIndex, endIndex, url);
  }
}

function aiSummarizer() {
  var doc = DocumentApp.getActiveDocument();
  var selectedText = doc.getBody().getText();
  var body = doc.getBody();

  var apiKey = "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";

  var model = "gpt-4"; // intializes the model
  var temperature = 0;
  var maxTokens = 1000;

  let paragraphs = body.getParagraphs();
  paragraphs.forEach(para => {
    if((para.getText() !== null && para.getText().trim() !== "" && para.getText().trim() !== "\n") && para.getHeading()==DocumentApp.ParagraphHeading.NORMAL && para.getText().trim().length > 60) { // goes through each paragraphb only if they contain non heading text are longer than 60 characters
      var urlData = extractUrlTextAndLinkFromParagraph(para.editAsText());
      var urls = urlData.map(item => item.linkedText); // grabs all the links in the paragraph
      selectedText = para.getText();
      var urlExist = false;

      if(urls.length > 0) { // if there is a link in the paragraph runs this
        var prompt = "I have section of text that I need to make more accessible for individuals with dyslexia and those using screen readers. I want you to rewrite the text with simplified sentences that are slightly easier for humans and screen readers to understand. Maintain the same formatting and return only a rewritten section of text it must be one paragraph. There are also some peices of text that contains links here are the linked texts in array form please put this text in your response: " + urls.toString(); +  "|| Here is the text keep the formatting the same in your response  MAKE SURE IT IS ALL ONE PARAGRAPH in your response: \n\n " + selectedText;
        urlExist = true;
      }
      else {
        var prompt = "I have section of text that I need to make more accessible for individuals with dyslexia and those using screen readers. I want you to rewrite the text with simplified sentences that are slightly easier for humans and screen readers to understand. Maintain the same formatting and return only a rewritten section of text. Here is the text: \n\n " + selectedText;
      }
      const requestBody = {
        "model": model,
        "messages": [{"role": "system", "content": prompt}],
        "temperature": temperature,
        "max_tokens": maxTokens,
      };

      const requestOptions = {
        "method": "POST",
        "headers": {
          "Content-Type": "application/json",
          "Authorization": "Bearer " + apiKey
        },
        "payload": JSON.stringify(requestBody) // formats the data into a form that GPT will accept
      };

      const response = UrlFetchApp.fetch("https://api.openai.com/v1/chat/completions", requestOptions); // Gets a response for our API

      var responseText = response.getContentText();
      var json = JSON.parse(responseText);
      Logger.log(json['choices'][0]['message']['content']);

      para.setText(json['choices'][0]['message']['content']);
      }
      if(urlExist) { // If there were links in the original paragraph then you should go back to the GPT response and relink the data
        urlData.forEach(urlControl => {
          linkTextToUrl(para.editAsText(), urlControl.linkedText, urlControl.url);
        })
      }
    })
}

function onOpen() {
  // Create a custom menu in the Google Docs UI
  DocumentApp.getUi()
    .createMenu('MLAiffy')
    .addItem('2. Format Document', 'formatDocument')
    .addToUi();
}

function formatDocument() {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();

  // Set document properties
  doc.setName("MLA_Formatted_Document");

  // Set page margin
  body.setMarginTop(60).setMarginBottom(60).setMarginLeft(60).setMarginRight(60);

  // Add template data to the top left corner
  var headerText = "Your Name\nTeacher's Name\nClass Name and Period\n" + getCurrentDate();
  body.insertParagraph(0, headerText).setHeading(DocumentApp.ParagraphHeading.HEADING6);

  // Set font to Times New Roman
  body.setFontFamily("Times New Roman");

  // Add page numbers in the header
  addHeaderWithPageNumbers();

  // Set custom styles under "Custom > MLAiffy"
  var customStyle = {};
  customStyle[DocumentApp.Attribute.FONT_SIZE] = 12;
  customStyle[DocumentApp.Attribute.LINE_SPACING] = 2;

  // Apply custom styles
  body.setAttributes(customStyle);

  Logger.log("MLA Formatting Complete!");
}

function addHeaderWithPageNumbers() {
  var doc = DocumentApp.getActiveDocument();

  // Count the number of section breaks in the document
  var totalSects = countSectionBreaks(doc);

  // Loop through each section and add the header with page number
  for (var i = 1; i <= totalSects; i++) {
    // Create a new header for each section
    var headerSection = doc.addHeader();


    // Add a paragraph to the header for the page number
    var pageNumPar = headerSection.appendParagraph("Last Name " + i);
    pageNumPar.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
    pageNumPar.setFontFamily("Times New Roman");
    pageNumPar.setFontSize(12);

    // Add a section break after each section except the last one
    if (i < totalSects) {
      doc.getBody().appendSectionBreak(DocumentApp.SectionType.NEXT_PAGE);
    }
  }

  Logger.log("Header with Page Numbers added successfully!");
}

// Helper function to count section breaks in the document
function countSectionBreaks(doc) {
  var body = doc.getBody();
  var elements = body.getNumChildren();
  var sectionBreakCount = 1; // Start with one section, assuming no section breaks at the beginning

  for (var i = 0; i < elements; i++) {
    var element = body.getChild(i);
    var elementType = element.getType();

    if (elementType === DocumentApp.ElementType.SECTION_BREAK) {
      sectionBreakCount++;
    }
  }

  return sectionBreakCount;
}

function getCurrentDate() {
  var now = new Date();
  var formattedDate = Utilities.formatDate(now, "GMT", "dd MMMM yyyy");
  return formattedDate;
}

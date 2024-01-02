function onOpen() {
  SlidesApp.getUi()
    .createMenu('Markdown Slides')
    .addItem('Paste Markdown', 'showDialog')
    .addToUi();
}

function showDialog() {
  var html = HtmlService.createHtmlOutputFromFile('Page')
      .setWidth(500)
      .setHeight(300);
  SlidesApp.getUi()
      .showModalDialog(html, 'Paste Markdown File');
}

function processMarkdown(markdownContent) {
  var presentation = SlidesApp.getActivePresentation();
  var slides = presentation.getSlides();
  var lines = markdownContent.split('\n');
  var currentSlide = null;
  var notes = '';
  var title = '';
  var imgUrl = '';

  // Get the height and width of the slide
  var slideWidth = presentation.getPageWidth();
  var slideHeight = presentation.getPageHeight();

  lines.forEach(function(line) {
    if ((line.startsWith('---')) || (line.startsWith('***')) || (line.startsWith('___'))) {
      // Handle previous slide's notes
 
      if (title.length > 0) {
        currentSlide = presentation.appendSlide(SlidesApp.PredefinedLayout.SECTION_HEADER);
        // Insert a text box that spans the entire slide
        //var textBox = currentSlide.insertTextBox(title, 20, 0, slideWidth-40, slideHeight);
        // Set the title text
        var textBox = currentSlide.getPlaceholder(SlidesApp.PlaceholderType.TITLE);
        if (textBox) {
          textBox.asShape().getText().setText(title);
          textBox.setWidth(slideWidth-36);
          textBox.setLeft(18);
          currentBox = textBox;
          textBox.bringToFront();
          centerAlignTextInShape(textBox);
        } else {
          Logger.log("No title placeholder found on this layout.");
        }
      } else {
        currentSlide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
      }
      Logger.log("New Slide ImageURL: " + imageUrl)
      if (imageUrl != '') {
        try {
        var image = currentSlide.insertImage(imageUrl);
        // Resize and position the image
        if (title == "") {
          Logger.log("No title:" + slideWidth + "X" + slideHeight)
          setBackgroundImage(image,slideWidth,slideHeight);
        } else {
        resizeAndCropImage(image,slideWidth,slideHeight,currentBox);
        }
        // Send the image to back
        image.sendToBack(); 
        } catch (e) {
          Logger.log("Couldn't get image...")
        } 

      }

        currentSlide.getNotesPage().getSpeakerNotesShape().getText().setText(notes);
        notes = '';
        title = '';
        imageURL = '';

    }

      if (line.startsWith('#')) {
        title=line.substring(1).trim();
      }
      else if (line.startsWith('![')) {
      // Extract and insert image
      imageUrl = line.substring(line.indexOf('(') + 1, line.lastIndexOf(')'));
      Logger.log("Got ImageURL: " + imageUrl)
    } else {
      notes += line + '\n';
    }
  });

  // Handle notes for the last slide
  if (currentSlide && notes) {
    currentSlide.getNotesPage().getSpeakerNotesShape().getText().setText(notes);
  }
}

function setBackgroundImage(image, slideWidth, slideHeight) {
  // Calculate aspect ratio
  var imageWidth = image.getWidth();
  var imageHeight = image.getHeight();
  var aspectRatio = imageWidth / imageHeight;
  if (aspectRatio > 1) {
    image.setWidth = slideWidth;
    iw=image.getWidth();
    Logger.log("Imagewidth vs SlideWidth: " + iw + " vs " + slideWidth);
    if (iw < slideWidth) {
        image.setLeft((slideWidth-iw)/2);
      } else {
        image.setTop((slideHeight-imageHeight)/2);
      }
    }
    else {
      image.setHeight = slideHeight;
      iw=image.getWidth();
      image.setLeft((slideWidth-iw)/2);
    }

}


function resizeAndCropImage(image, slideWidth, slideHeight,textBox) {
  var imageWidth = image.getWidth();
  var imageHeight = image.getHeight();
  var halfWidth = slideWidth / 2;
  
  // Calculate aspect ratio
  var aspectRatio = imageWidth / imageHeight;

  // Position image on the right half of the slide
  var newImageWidth = image.getWidth();
  var positionX = slideWidth - newImageWidth + (newImageWidth - halfWidth) / 2;
  var positionY = (slideHeight - newImageHeight) / 2;

  if (aspectRatio > 1) { // Image is wider than tall
    // Resize image to the height of the slide
    image.setHeight(slideHeight);
    image.setWidth(aspectRatio * slideHeight);

    image.setLeft(positionX > 0 ? positionX : 0);
    //Logger.log(textBox)
    textBox.setWidth(positionX-30);
  } else { // Image is taller than wide
    // Resize image to half the width of the slide
    image.setWidth(halfWidth);
    image.setHeight(halfWidth / aspectRatio);

    // Position image on the right and vertically center
    var newImageHeight = image.getHeight();

    image.setLeft(slideWidth - halfWidth);
    image.setTop(positionY > 0 ? positionY : 0);
    textBox.setWidth(positionX-54);
  }
}

function centerAlignTextInShape(textShape) {
  var textRange = textShape.asShape().getText(); // Get the text range
  var paragraphStyle = textRange.getParagraphStyle(); // Get the paragraph style
  paragraphStyle.setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER); // Center align the text
}




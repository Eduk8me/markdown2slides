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
  var imageUrl = '';
  var subtitle = '';
  var titleBox = null;
  var subtitleBox = null;

  // Get the height and width of the slide
  var slideWidth = presentation.getPageWidth();
  var slideHeight = presentation.getPageHeight();

  lines.forEach(function(line) {
    if ((line.startsWith('---')) || (line.startsWith('***')) || (line.startsWith('___'))) {
      //Start new slide
      if (title.length > 0) {
        currentSlide = presentation.appendSlide(SlidesApp.PredefinedLayout.TITLE);
        var titleBox = currentSlide.getPlaceholder(SlidesApp.PlaceholderType.CENTERED_TITLE);
        if (titleBox && title.length > 0 ) {
          titleBox.asShape().getText().setText(title);
          titleBox.setWidth(slideWidth-36);
          titleBox.setLeft(18);
          titleBox.bringToFront();
          //centerAlignTextInShape(titleBox);
        } else {
          Logger.log("No title placeholder found on this layout.");
        }
        
        var subtitleBox = currentSlide.getPlaceholder(SlidesApp.PlaceholderType.SUBTITLE);
        if (subtitleBox && subtitle.length > 0 ) {
          subtitleBox.asShape().getText().setText(subtitle);
          subtitleBox.setWidth(slideWidth-36);
          subtitleBox.setLeft(18);
          subtitleBox.bringToFront();
          //centerAlignTextInShape(subtitleBox);
        } else {
          Logger.log("No title placeholder found on this layout.");
        }
      } else {
        currentSlide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
      }
      // Is there an image for the slide?
      Logger.log("New Slide ImageURL: " + imageUrl)
      if (imageUrl != '') {
        try {
        var image = currentSlide.insertImage(imageUrl);
        // Resize and position the image
        if (title == "") {
          Logger.log("No title:" + slideWidth + "X" + slideHeight)
          setBackgroundImage(image,slideWidth,slideHeight);
        } else {
        resizeAndCropImage(image,slideWidth,slideHeight,titleBox,subtitleBox);
        }
        // Send the image to back
        image.sendToBack(); 
        } catch (e) {
          Logger.log("Couldn't get image...")
        } 

      }
      //Add speaker notes
      currentSlide.getNotesPage().getSpeakerNotesShape().getText().setText(notes);

      //Reset everything for next slide
      notes = '';
      title = '';
      subtitle = '';
      imageUrl = '';
      titleBox = null;
      subtitleBox = null;

    }
    else {
        if (line.startsWith('##')) {
          subtitle += line.substring(2).trim() + '\n';
        }
        else if (line.startsWith('#')) {
          title+=line.substring(1).trim() +'\n';
        }
        else if (line.startsWith('![')) {
        // Extract and insert image
        imageUrl = line.substring(line.indexOf('(') + 1, line.lastIndexOf(')'));
        Logger.log("Got ImageURL: " + imageUrl)
        } else {
          notes += line + '\n';
        }
    }
  });

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


function resizeAndCropImage(image, slideWidth, slideHeight,titleBox,subtitleBox) {
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
    if (titleBox) { titleBox.setWidth(positionX-30); }
    if (subtitleBox) { subtitleBox.setWidth(positionX-30); }
  } else { // Image is taller than wide
    // Resize image to half the width of the slide
    image.setWidth(halfWidth);
    image.setHeight(halfWidth / aspectRatio);

    // Position image on the right and vertically center
    var newImageHeight = image.getHeight();

    image.setLeft(slideWidth - halfWidth);
    image.setTop(positionY > 0 ? positionY : 0);
    if (titleBox) { titleBox.setWidth(positionX-54); }
    if (subtitleBox) { subtitleBox.setWidth(positionX-54); }
  }
}

function centerAlignTextInShape(textShape) {
  var textRange = textShape.asShape().getText(); // Get the text range
  var paragraphStyle = textRange.getParagraphStyle(); // Get the paragraph style
  paragraphStyle.setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER); // Center align the text
}




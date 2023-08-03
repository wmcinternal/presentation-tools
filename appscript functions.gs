
function addTextToSheet(slides, columnnum){ 
  let counter = 2;
  for(var j = 0; j<slides.length; j++){
    var pageElements = slides[j].getPageElements(); 
    
    var a;
    for(var i = 0; i<pageElements.length; i++){
      if(pageElements[i].getPageElementType() == 'SHAPE'){

        a = pageElements[i].asShape().getText();
        // Logger.log("this is a, %s", a.asRenderedString()); 
        // Logger.log("test %s", translations.getRange(1,1).getValue());
        var range = translations.getRange(counter, columnnum);
        translations.getRange(counter, columnnum-1).setValue(j+1);
        range.setValue(a.asString());
        counter++;
    }
  }
}
}

function updateTranslationsSheet(){
  var slides = SlidesApp.openById(masterPPT).getSlides();
  addTextToSheet(slides, 2) // Edit column number
}

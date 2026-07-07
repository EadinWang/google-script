function replaceTextColor() {
  // === CONFIGURE THESE ===
  const TARGET_COLOR = '#0177c6';   // color to find (hex)
  const NEW_COLOR    = '#425de0';   // color to replace it with (hex)
  // 

  const presentation = SlidesApp.getActivePresentation();
  const slides = presentation.getSlides();

  let changedCount = 0;

  function processTextRange(textRange) {
    if (!textRange || textRange.asString().trim() === '') return;
    const runs = textRange.getRuns();
    runs.forEach(run => {
      const style = run.getTextStyle();
      const colorObj = style.getForegroundColor();
      if (colorObj && colorObj.getColorType() === SlidesApp.ColorType.RGB) {
        const hex = rgbToHex(colorObj.asRgbColor());
        if (hex.toUpperCase() === TARGET_COLOR.toUpperCase()) {
          style.setForegroundColor(NEW_COLOR);
          changedCount++;
        }
      }
    });
  }

  slides.forEach(slide => {
    slide.getShapes().forEach(shape => {
      if (shape.getText) processTextRange(shape.getText());
    });

    slide.getTables().forEach(table => {
      const numRows = table.getNumRows();
      const numCols = table.getNumColumns();
      for (let r = 0; r < numRows; r++) {
        for (let c = 0; c < numCols; c++) {
          processTextRange(table.getCell(r, c).getText());
        }
      }
    });
  });

  Logger.log(`Done. Changed ${changedCount} text run(s) from ${TARGET_COLOR} to ${NEW_COLOR}.`);
  SlidesApp.getUi().alert(`Replaced color on ${changedCount} text run(s).`);
}

function rgbToHex(rgbColor) {
  const r = rgbColor.getRed();
  const g = rgbColor.getGreen();
  const b = rgbColor.getBlue();
  return '#' + [r, g, b].map(v => v.toString(16).padStart(2, '0')).join('');
}

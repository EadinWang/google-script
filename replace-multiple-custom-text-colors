function replaceMultipleTextColors() {
  // === CONFIGURE YOUR COLOR MAP HERE ===
  const COLOR_MAP = {
    '0087e4': '3D69F0', // 'original color': 'color to replace with'
    '0044d4': '525fc7',
    'a0a1a7': '6E7687',
    'a5a5a5': '6E7687' //add more pairs if needed
  };
  // =======================================

  const normalizedMap = {};
  Object.keys(COLOR_MAP).forEach(key => {
    const cleanKey = key.replace('#', '').toLowerCase();
    const cleanVal = COLOR_MAP[key].replace('#', '').toLowerCase();
    normalizedMap[cleanKey] = cleanVal;
  });

  const presentation = SlidesApp.getActivePresentation();
  const slides = presentation.getSlides();

  let changedCount = 0;
  const changeLog = {};

  function processTextRange(textRange) {
    if (!textRange || textRange.asString().trim() === '') return;
    const runs = textRange.getRuns();
    runs.forEach(run => {
      const style = run.getTextStyle();
      const colorObj = style.getForegroundColor();
      if (colorObj && colorObj.getColorType() === SlidesApp.ColorType.RGB) {
        const hex = rgbToHex(colorObj.asRgbColor()).replace('#', '').toLowerCase(); // <-- FIX: strip # here too

        if (normalizedMap.hasOwnProperty(hex)) {
          const newHex = '#' + normalizedMap[hex];
          style.setForegroundColor(newHex);
          changedCount++;
          changeLog[hex] = (changeLog[hex] || 0) + 1;
        }
      }
    });
  }

  function scanShape(shape) {
    if (shape.getText) processTextRange(shape.getText());
  }

  slides.forEach(slide => {
    slide.getShapes().forEach(scanShape);

    slide.getGroups().forEach(group => {
      group.getChildren().forEach(child => {
        if (child.asShape) {
          try { scanShape(child.asShape()); } catch (e) {}
        }
      });
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

  let summary = `Total changed: ${changedCount}\n`;
  Object.keys(changeLog).forEach(oldHex => {
    summary += `#${oldHex} → #${normalizedMap[oldHex]}: ${changeLog[oldHex]}\n`;
  });

  Logger.log(summary);
  SlidesApp.getUi().alert(summary);
}

function rgbToHex(rgbColor) {
  const r = rgbColor.getRed();
  const g = rgbColor.getGreen();
  const b = rgbColor.getBlue();
  return '#' + [r, g, b].map(v => v.toString(16).padStart(2, '0')).join('');
}

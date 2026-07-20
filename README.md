# Slides Color Replacer

Two Google Apps Script that replaces custom text colors across a Google Slides presentation (single or in bulk). Scans all slides including shapes, grouped shapes, and table cells — and swaps matching RGB text colors for new ones.
 
Two versions are included:
 
- **Single color** — replace one specific color with another.
- **Multi-color** — replace several colors at once using a configurable map, ideal for remapping a whole palette (e.g. syntax-highlighting themes) in one pass.

## How to use
 
1. Open your Google Slides presentation.
2. Go to **Extensions → Apps Script**.
3. Delete any boilerplate code and paste in one of the scripts below.

**Configure:** replace the color codes based on your needs. 

   ```javascript
  // replace multiple colors
   const COLOR_MAP = {
     '0044d4': '525fc7', // '[old color code]': '[new color code]'
     'a0a1a7': '6E7687',
     'a5a5a5': '6E7687', //add more pairs if needed
   };
   ```
  ```javascript
  // replace single color
  function replaceTextColor() {
    // === CONFIGURE THESE ===
    const TARGET_COLOR = '#0087e4';   // color to find (hex)
    const NEW_COLOR    = '#486EE0';   // color to replace it with (hex)
  ```

4. Click **Save** (Ctrl/Cmd + S).
5. Select the function to run from the dropdown at the top, then click **Run**.
6. Authorize the script the first time you run it.
7. Check the alert popup or **Executions → Logs** for a summary. Output looks like:
```
RGB colors found in this presentation:
#0087e4 — 2 run(s)
#0044d4 — 2 run(s)
#a0a1a7 — 2 run(s)
#a5a5a5 — 2 run(s)
```

> ⚠️ **Tip:** If you paste code from a source that uses "smart quotes" (`‘ ’` / `“ ”`), Apps Script will throw a syntax error. Make sure all quotes are straight (`'` / `"`) before running.

## Finding your current colors

Not sure what hex codes are actually used in your deck? Run the included `debugAllColors()` function first. It scans the whole presentation and logs every custom RGB color it finds, along with how many runs use each one. Use this to confirm the exact hex values before adding them to `COLOR_MAP`.

## Finding your target colors
 
Not sure what hex codes are currently used in your deck? Run this diagnostic script first — it scans the presentation and logs every RGB color found along with how many text runs use it:
 
```javascript
function debugAllColors() {
  const slides = SlidesApp.getActivePresentation().getSlides();
  const foundColors = {};
 
  function scan(textRange) {
    if (!textRange || textRange.asString().trim() === '') return;
    textRange.getRuns().forEach(run => {
      const style = run.getTextStyle();
      const colorObj = style.getForegroundColor();
      if (colorObj && colorObj.getColorType() === SlidesApp.ColorType.RGB) {
        const hex = rgbToHex(colorObj.asRgbColor()).toLowerCase();
        foundColors[hex] = (foundColors[hex] || 0) + 1;
      }
    });
  }
 
  function scanShape(shape) {
    if (shape.getText) scan(shape.getText());
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
          scan(table.getCell(r, c).getText());
        }
      }
    });
  });
 
  let output = 'RGB colors found in this presentation:\n';
  Object.keys(foundColors).sort().forEach(hex => {
    output += `#${hex} — ${foundColors[hex]} run(s)\n`;
  });
 
  Logger.log(output);
  SlidesApp.getUi().alert(output);
}
 
function rgbToHex(rgbColor) {
  const r = rgbColor.getRed();
  const g = rgbColor.getGreen();
  const b = rgbColor.getBlue();
  return '#' + [r, g, b].map(v => v.toString(16).padStart(2, '0')).join('');
}
```


## Notes
 
- Colors must match exactly (case-insensitive hex comparison). Use `debugAllColors` above to confirm exact hex values before configuring your target colors.
- Theme colors (e.g. colors inherited from the slide theme, shown as `THEME` type rather than `RGB`) are not matched by these scripts.
- Changes apply directly to the active presentation — consider duplicating your file first if you want a backup before running.
## Credits
 
Written by Eadin Wang with Claude (Anthropic).

## License

MIT — free to use, modify, and share.

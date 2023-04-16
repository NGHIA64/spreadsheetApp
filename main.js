function formatSpreadsheet() {
    var darkColor = ["#60a3bc", "#dcbe99", "#91c4c4", "#f1a3a3", "#bdb7d6", "#beada7", "#a5d6a7", "#f0d7c2", "#a2d7dd", "#d6d7a7", "#d7a7c3", "#c7b7b7", "#7cb8e2", "#d1c3b1", "#c8d1b1", "#f0b8b8", "#b1afd1", "#b1a9a9", "#a2d6d7", "#b5d1b5", "#d6a7a2", "#a9b1d1", "#c7c7b7", "#7cb8d2", "#b1c3d1", "#c8b1d1", "#e2b87c", "#d1b1c3", "#c3d1b1", "#b8d27c"]
    var colorSpentForTab
    var rowTable = 1
    Object.values(controlSheet).forEach((category, index) => {
      colorSpentForTab = []
      Object.values(category).forEach((sheet, index) => {
        var randomColorForTab = darkColor[Math.floor(Math.random() * darkColor.length)];
        if (colorSpentForTab.includes(randomColorForTab)) {
          randomColorForTab = darkColor.filter(item => !colorSpentForTab.includes(item))[Math.floor(Math.random() * darkColor.length)]
        }
        try {
          sheet.getRange(1,1,sheet.getMaxRows(),sheet.getMaxColumns()).setVerticalAlignment("middle").setHorizontalAlignment("center").setBorder(false, false, false, false, false, false)
          sheet.setTabColor(randomColorForTab);
          sheet.getRange(1, 1, 1, sheet.getLastColumn()).setBackground(randomColorForTab).setVerticalAlignment("middle").setHorizontalAlignment("center").setFontWeight("bold").setBorder(true, true, true, true, true, true)
          sheet.setColumnWidths(1, sheet.getLastColumn(), 150).setRowHeights(1, 1, 35)
          colorSpentForTab.push(randomColorForTab)
        } catch (err) {
  
        }
        rowTable+=1
      });
    });
  }
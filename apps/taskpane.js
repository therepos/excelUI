Office.onReady(() => {
  // Office is ready
});

function setFont(fontName) {
  Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    sheets.load("items");
    await context.sync();

    sheets.items.forEach(sheet => {
      const range = sheet.getUsedRange();
      range.format.font.name = fontName;
      range.format.font.size = 8;
    });

    await context.sync();
  });
}

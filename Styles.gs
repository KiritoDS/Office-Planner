class Styles {

  static apply() {

    const ss = SpreadsheetApp.getActiveSpreadsheet();

    MONTHS.forEach(month => {

      const sheet = ss.getSheetByName(month);

      if (!sheet) return;

      Styles.styleMonth(sheet);

    });

    SpreadsheetApp.getUi().alert("Оформлення застосовано.");

  }

  static styleMonth(sheet) {

    const lastRow = sheet.getLastRow();

    // Ширина колонок
    sheet.setColumnWidths(1, 7, 200);
    sheet.setRowHeights(4, sheet.getMaxRows() - 3, 28);

    // Назва місяця
    sheet.getRange("A1:G1")
      .setBackground("#1F4E78")
      .setFontColor("white")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setFontWeight("bold");

    // Дні тижня
    sheet.getRange("A3:G3")
      .setBackground("#D9EAD3")
      .setHorizontalAlignment("center")
      .setFontWeight("bold");

    // Проходимо по всіх днях
    for (let row = 4; row <= lastRow; row += 10) {

      for (let col = 1; col <= 7; col++) {

        if (sheet.getRange(row, col).getValue() == "")
          continue;

        // День
        sheet.getRange(row, col)
          .setBackground("#B6D7A8")
          .setFontWeight("bold")
          .setHorizontalAlignment("center");

        // Кімната 1
        sheet.getRange(row + 1, col)
          .setBackground("#D0E0F5")
          .setFontWeight("bold");

        // Кімната 2
        sheet.getRange(row + 4, col)
          .setBackground("#D9EAD3")
          .setFontWeight("bold");

        // Кімната 3
        sheet.getRange(row + 7, col)
          .setBackground("#FFF2CC")
          .setFontWeight("bold");

        // Рамка навколо всього дня
        sheet.getRange(row, col, 10, 1)
          .setBorder(
            true,
            true,
            true,
            true,
            true,
            true,
            "black",
            SpreadsheetApp.BorderStyle.SOLID_MEDIUM
          );

      }

    }

  }

}

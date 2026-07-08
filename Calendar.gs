class CalendarBuilder {

  static buildAllMonths() {

    const ss = SpreadsheetApp.getActiveSpreadsheet();

    MONTHS.forEach((monthName, index) => {

      const sheet = ss.getSheetByName(monthName);

      if (sheet) {
        CalendarBuilder.buildMonth(sheet, SETTINGS.YEAR, index + 1);
      }

    });

  }

  static buildMonth(sheet, year, month) {

    sheet.clear();

    sheet.setColumnWidths(1, 7, 170);

    sheet.getRange("A1:G1")
      .merge()
      .setValue(`${MONTHS[month - 1]} ${year}`)
      .setFontSize(18)
      .setFontWeight("bold")
      .setHorizontalAlignment("center");

    // Дні тижня
    sheet.getRange(3, 1, 1, 7).setValues([DAYS]);

    const firstDay = new Date(year, month - 1, 1);

    let startColumn = firstDay.getDay();
    startColumn = startColumn === 0 ? 7 : startColumn;

    const daysInMonth = new Date(year, month, 0).getDate();

    let row = 4;
    let col = startColumn;

    for (let day = 1; day <= daysInMonth; day++) {

      sheet.getRange(row, col)
        .setValue(day)
        .setFontWeight("bold")
        .setBackground("#D9EAD3");

      col++;

      if (col > 7) {
        col = 1;
        row += 15; // залишаємо місце під кімнати
      }

    }

  }

}

class Rooms {

  static buildRooms() {

    const ss = SpreadsheetApp.getActiveSpreadsheet();

    MONTHS.forEach(month => {

      const sheet = ss.getSheetByName(month);

      if (!sheet) return;

      Rooms.buildMonthRooms(sheet);

    });

    SpreadsheetApp.getUi().alert("Кімнати створені.");
  }

  static buildMonthRooms(sheet) {

    const lastRow = sheet.getLastRow();

    for (let row = 4; row <= lastRow; row += 11) {

      for (let col = 1; col <= 7; col++) {

        const day = sheet.getRange(row, col).getValue();

        if (day == "") continue;

        sheet.getRange(row + 1, col).setValue("🏢 Кімната 1");
        sheet.getRange(row + 2, col).setValue("A");
        sheet.getRange(row + 3, col).setValue("B");

        sheet.getRange(row + 4, col).setValue("🏢 Кімната 2");
        sheet.getRange(row + 5, col).setValue("A");
        sheet.getRange(row + 6, col).setValue("B");

        sheet.getRange(row + 7, col).setValue("🏢 Кімната 3");
        sheet.getRange(row + 8, col).setValue("");
        sheet.getRange(row + 9, col).setValue("");
        
        sheet.getRange(row + 10, col)
        .setValue("🟢 Вільно: 18")
        .setFontWeight("bold")
       .setHorizontalAlignment("center")
        .setBackground("#E2F0D9");

        

      }

    }

  }

}

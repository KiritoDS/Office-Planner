class Validation {

  static buildDropdowns() {

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const groupSheet = ss.getSheetByName("Групи");

    const groupRange = groupSheet.getRange(
      2,
      1,
      groupSheet.getLastRow() - 1,
      1
    );

    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInRange(groupRange, true)
      .setAllowInvalid(false)
      .build();

    MONTHS.forEach(month => {

      const sheet = ss.getSheetByName(month);

      if (!sheet) return;

      const lastRow = sheet.getLastRow();

      for (let row = 4; row <= lastRow; row += 11) {

        for (let col = 1; col <= 7; col++) {

          if (sheet.getRange(row, col).getValue() == "") continue;

          const dropdownRows = [
            row + 2,
            row + 3,
            row + 5,
            row + 6,
            row + 8,
            row + 9
          ];

          dropdownRows.forEach(r => {

            const cell = sheet.getRange(r, col);

            cell.clearContent();
            cell.setDataValidation(rule);
            cell.setValue("—");

          });

        }

      }

    });

    SpreadsheetApp.getUi().alert("Dropdown створені!");

  }

}

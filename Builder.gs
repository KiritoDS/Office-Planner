class Builder {

  static build() {

    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Видаляємо всі аркуші, крім Dashboard
    while (ss.getSheets().length > 1) {
      ss.deleteSheet(ss.getSheets()[1]);
    }

    const dashboard = ss.getSheets()[0];
    dashboard.clear();
    dashboard.setName("Dashboard");

    Builder.buildDashboard_(dashboard);
    Builder.buildSettings_(ss);
    Builder.buildGroups_(ss);

    MONTHS.forEach(month => {
      Builder.createMonthSheet_(ss, month);
    });
    CalendarBuilder.buildAllMonths();

    SpreadsheetApp.getUi().alert("Структуру Planner створено.");
  }

  static buildDashboard_(sheet) {

    sheet.getRange("A1").setValue("🏢 OFFICE PLANNER");
    sheet.getRange("A2").setValue("Рік");
    sheet.getRange("B2").setValue(SETTINGS.YEAR);

    sheet.getRange("A4").setValue("Статистика");
    sheet.getRange("A5").setValue("Всього місць");
    sheet.getRange("A6").setValue("Зайнято");
    sheet.getRange("A7").setValue("Вільно");

    sheet.getRange("A1:B7").setFontWeight("bold");

  }

  static buildSettings_(ss) {

    const sh = ss.insertSheet("Налаштування");

    sh.getRange("A1:B5").setValues([
      ["Параметр","Значення"],
      ["Рік",SETTINGS.YEAR],
      ["Кімнат",SETTINGS.ROOMS],
      ["Блоків",SETTINGS.BLOCKS_PER_ROOM],
      ["Місць",SETTINGS.SEATS_PER_BLOCK]
    ]);

  }

  static buildGroups_(ss){

  const sh = ss.insertSheet("Групи");

  sh.getRange("A1").setValue("Група");

  sh.getRange("A2:A7").setValues([
    ["Team Alpha"],
    ["Team Bravo"],
    ["Team Charlie"],
    ["Team Delta"],
    ["Team Echo"],
    ["—"]
  ]);

  sh.setColumnWidth(1,220);

}

  static createMonthSheet_(ss,name){

    const sh = ss.insertSheet(name);

    sh.getRange("A1")
      .setValue(name + " " + SETTINGS.YEAR)
      .setFontSize(18)
      .setFontWeight("bold");

  }

}

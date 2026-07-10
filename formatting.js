
function customizeSheet(sheet, numberOfMembers, validRehearsals) {

  const lastColumn = validRehearsals + 3;
  const lastRow = numberOfMembers + 1;

  //-----------------------------
  // Fonte
  //-----------------------------

  sheet.getRange(1, 1, lastRow, lastColumn)
       .setFontFamily("Roboto")
       .setFontSize(10);

  //-----------------------------
  // Cabeçalho
  //-----------------------------

  sheet.getRange(1, 1, 1, lastColumn)
      .setBackground("#1F4E78")
      .setFontColor("white")
      .setFontWeight("bold")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");

  //-----------------------------
  // Coluna Nome
  //-----------------------------

  sheet.getRange(2, 1, numberOfMembers, 1)
      .setFontWeight("bold")
      .setHorizontalAlignment("left");

  //-----------------------------
  // Pontos + Total + %
  //-----------------------------

  sheet.getRange(2, 2, numberOfMembers, validRehearsals + 2)
      .setHorizontalAlignment("center");

  //-----------------------------
  // Fundo Total e %
  //-----------------------------

  sheet.getRange(1, validRehearsals + 2, lastRow, 2)
      .setBackground("#F2F2F2");

  sheet.getRange(1, validRehearsals + 2, 1, 2)
      .setBackground("#1F4E78")
      .setFontColor("white");

  //-----------------------------
  // Borda vertical antes do Total
  //-----------------------------

  sheet.getRange(1, validRehearsals + 2, lastRow, 1)
      .setBorder(
          null,
          true,
          null,
          true,
          null,
          null,
          "black",
          SpreadsheetApp.BorderStyle.SOLID_MEDIUM
      );

  //-----------------------------
  // Bordas gerais
  //-----------------------------

  sheet.getRange(1, 1, lastRow, lastColumn)
      .setBorder(
          true,
          true,
          true,
          true,
          true,
          true,
          "#D9D9D9",
          SpreadsheetApp.BorderStyle.SOLID
      );

  //-----------------------------
  // Zebra stripes
  //-----------------------------

  for (let row = 2; row <= lastRow; row++) {

    if (row % 2 === 0) {

      sheet.getRange(row, 1, 1, lastColumn)
          .setBackground("#F7F7F7");

    }

  }

  //-----------------------------
  // Congelar
  //-----------------------------

  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(1);

  //-----------------------------
  // Largura das colunas
  //-----------------------------

  sheet.setColumnWidth(1, 220);

  for (let col = 2; col <= validRehearsals + 1; col++) {
    sheet.setColumnWidth(col, 65);
  }

  sheet.setColumnWidth(validRehearsals + 2, 70);
  sheet.setColumnWidth(validRehearsals + 3, 90);

  //-----------------------------
  // Formato das datas
  //-----------------------------

  sheet.getRange(
      1,
      2,
      1,
      validRehearsals
  ).setNumberFormat("dd/MM");

  //-----------------------------
  // Formato da percentagem
  //-----------------------------

  sheet.getRange(
      2,
      validRehearsals + 3,
      numberOfMembers,
      1
  ).setNumberFormat("0.0%");

  //-----------------------------
  // Filtro
  //-----------------------------

  const range = sheet.getRange(1, 1, lastRow, lastColumn);

  if (sheet.getFilter()) {
    sheet.getFilter().remove();
  }

  range.createFilter();

  //-----------------------------
  // Ajustar altura das linhas
  //-----------------------------

  sheet.autoResizeRows(1, lastRow);

}
function applyConditionalFormatting(sheet, numberOfMembers, validRehearsals){

  // Remover regras antigas
  sheet.clearConditionalFormatRules();

  const rules = [];

  //-----------------------------
  // Pontos (5 = Verde)
  //-----------------------------

  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberEqualTo(5)
      .setBackground("#C6EFCE")
      .setFontColor("#006100")
      .setRanges([
        sheet.getRange(
          2,
          2,
          numberOfMembers,
          validRehearsals
        )
      ])
      .build()
  );

  //-----------------------------
  // Pontos (4 = Amarelo)
  //-----------------------------

  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberEqualTo(4)
      .setBackground("#FFEB9C")
      .setFontColor("#9C6500")
      .setRanges([
        sheet.getRange(
          2,
          2,
          numberOfMembers,
          validRehearsals
        )
      ])
      .build()
  );

  //-----------------------------
  // Pontos (0 = Vermelho)
  //-----------------------------

  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberEqualTo(0)
      .setBackground("#FFC7CE")
      .setFontColor("#9C0006")
      .setRanges([
        sheet.getRange(
          2,
          2,
          numberOfMembers,
          validRehearsals
        )
      ])
      .build()
  );

  //-----------------------------
  // Percentagem >= 80%
  //-----------------------------

  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThanOrEqualTo(0.8)
      .setBackground("#C6EFCE")
      .setFontColor("#006100")
      .setRanges([
        sheet.getRange(
          2,
          validRehearsals + 3,
          numberOfMembers,
          1
        )
      ])
      .build()
  );

  //-----------------------------
  // Percentagem entre 50% e 80%
  //-----------------------------

  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberBetween(0.5, 0.799999)
      .setBackground("#FFEB9C")
      .setFontColor("#9C6500")
      .setRanges([
        sheet.getRange(
          2,
          validRehearsals + 3,
          numberOfMembers,
          1
        )
      ])
      .build()
  );

  //-----------------------------
  // Percentagem < 50%
  //-----------------------------

  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(0.5)
      .setBackground("#FFC7CE")
      .setFontColor("#9C0006")
      .setRanges([
        sheet.getRange(
          2,
          validRehearsals + 3,
          numberOfMembers,
          1
        )
      ])
      .build()
  );

  sheet.setConditionalFormatRules(rules);

}

function gerarResumo() {

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const responseSheet = spreadsheet.getSheetByName(SHEETS.RESPONSES);

  if (!responseSheet) {
    throw new Error(`Não existe a folha "${SHEETS.RESPONSES}".`);
  }


  const lastColumn = responseSheet.getLastColumn();
  const data = responseSheet.getDataRange().getValues();

  if (data.length < 2) {
    throw new Error("Não existem respostas.");
  }

  //-----------------------------
  // Criar ou limpar folha Resumo
  //-----------------------------

  let summarySheet = spreadsheet.getSheetByName(SHEETS.SUMMARY);

  if (!summarySheet) {
    summarySheet = spreadsheet.insertSheet(SHEETS.SUMMARY);
  } else {
    summarySheet.clear();
  }

  const headers = data[0];

  const firstMemberColumn = 2;
  const numberOfMembers = headers.length - firstMemberColumn;
  const numberOfRehearsals = data.length - 1;

  //-----------------------------
  // Cabeçalhos
  //-----------------------------

  summarySheet.getRange(1,1).setValue("Nome");

  let summaryColumn = 2;
  let validRehearsals = 0;

  for (let rehearsal = 0; rehearsal < numberOfRehearsals; rehearsal++) {

    const rehearsalDate = new Date(data[rehearsal + 1][1]);

    if (rehearsalDate < START_DATE || rehearsalDate > END_DATE) {
      continue;
    }

    summarySheet
      .getRange(1, summaryColumn)
      .setValue(rehearsalDate);

    summaryColumn++;
    validRehearsals++;

  }

  summarySheet.getRange(1, summaryColumn).setValue("Total");
  summarySheet.getRange(1, summaryColumn + 1).setValue("%");

  //-----------------------------
  // Membros
  //-----------------------------

  for (let member = 0; member < numberOfMembers; member++) {

    const column = member + firstMemberColumn;

    const fullHeader = headers[column];

    const name = fullHeader
      .replace("Registo de Presenças [", "")
      .replace("]", "");

    summarySheet.getRange(member + 2, 1).setValue(name);

    let totalPoints = 0;
    summaryColumn = 2;

    //-------------------------
    // Ensaios
    //-------------------------

    for (let rehearsal = 0; rehearsal < numberOfRehearsals; rehearsal++) {

      const rehearsalDate = new Date(data[rehearsal + 1][1]);

      if (rehearsalDate < START_DATE || rehearsalDate > END_DATE) {
        continue;
      }

      const state = data[rehearsal + 1][column];


      const point = POINTS[state] || 0;

      totalPoints += points;

      summarySheet
        .getRange(member + 2, summaryColumn)
        .setValue(points);

      summaryColumn++;

    }

    //-------------------------
    // Total
    //-------------------------

    summarySheet
      .getRange(member + 2, summaryColumn)
      .setValue(totalPoints);

    //-------------------------
    // Percentagem
    //-------------------------

    const percentage = totalPoints / (5 * validRehearsals);

    summarySheet
      .getRange(member + 2, summaryColumn + 1)
      .setValue(percentage)
      .setNumberFormat("0.0%");

  }

  //-----------------------------
  // Formatação
  //-----------------------------

  customizeSheet(
    summarySheet,
    numberOfMembers,
    validRehearsals
  );

  applyConditionalFormatting(
    summarySheet,
    numberOfMembers,
    validRehearsals
  );

}
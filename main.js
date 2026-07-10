function gerarResumo() {

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const responseSheet = spreadsheet.getSheetByName("Form_Responses");

  if (!responseSheet) {
    throw new Error("Folha 'Form_Responses' não encontrada.");
  }

  // Intervalo de datas a considerar
  const startDate = new Date(2026, 4, 1); // 1 Maio
  const endDate   = new Date(2026, 5, 30); // 30 Junho

  const data = responseSheet.getDataRange().getValues();

  if (data.length < 2) {
    throw new Error("Não existem respostas.");
  }

  //-----------------------------
  // Criar ou limpar folha Resumo
  //-----------------------------

  let summarySheet = spreadsheet.getSheetByName("Resumo");

  if (!summarySheet) {
    summarySheet = spreadsheet.insertSheet("Resumo");
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

    if (rehearsalDate < startDate || rehearsalDate > endDate) {
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

      if (rehearsalDate < startDate || rehearsalDate > endDate) {
        continue;
      }

      const state = data[rehearsal + 1][column];

      let points = 0;

      if (state == "Presente") {
        points = 5;
      }
      else if (state == "Atraso") {
        points = 4;
      }

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
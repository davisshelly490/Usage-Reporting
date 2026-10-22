const { FileBlob, SpreadsheetFile } = require("@oai/artifact-tool");

(async () => {
  const input = await FileBlob.load(process.argv[2]);
  const workbook = await SpreadsheetFile.importXlsx(input);
  workbook.recalculate();
  const summary = await workbook.inspect({
    kind: "table",
    range: "Summary!A1:F18",
    include: "values,formulas",
    tableMaxRows: 18,
    tableMaxCols: 6,
  });
  console.log(summary.ndjson);
  const errors = await workbook.inspect({
    kind: "match",
    searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!",
    options: { useRegex: true, maxResults: 50 },
  });
  console.log(errors.ndjson);
  const summarySheet = workbook.worksheets.getItem("Summary");
  const dataSheet = workbook.worksheets.getItem("Employee Level Data");
  const records = dataSheet.getUsedRange().values;
  const headers = records[0];
  const index = name => headers.indexOf(name);
  const rows = records.slice(1).filter(row => row.some(value => value !== null));
  const fields = ["Job", "BU", "Org 1 Report Category", "Location", "Alternate Title"];
  const cases = [Array(5).fill("All"), Array(5).fill("(Blank)")];
  for (let i = 0; i < fields.length; i++) {
    // Use the actual dropdown source value, not the employee cell's type.
    const dropdownValue = summarySheet.getCell(2, 26 + i).values[0][0];
    const selection = Array(5).fill("All");
    selection[i] = dropdownValue;
    if (!rows.some(row => row[index(fields[i])] === dropdownValue)) {
      throw new Error(`Dropdown/data type mismatch for ${fields[i]}: ${dropdownValue}`);
    }
    cases.push(selection);
  }
  const firstRoster = rows.find(row => row[index("In Roster")] === true);
  cases.push(fields.map(field => firstRoster[index(field)] || "(Blank)"));
  if (rows.some(row => row[index("BU")] === "1054")) {
    cases.push(["All", "1054", "All", "All", "All"]);
  }
  for (const selection of cases) {
    summarySheet.getRange("I2:I6").values = selection.map(value => [value]);
    workbook.recalculate();
    const selected = rows.filter(row => fields.every((field, i) => selection[i] === "All" ||
      (selection[i] === "(Blank)" ? row[index(field)] == null || row[index(field)] === "" : row[index(field)] === selection[i])));
    const expected = [
      selected.filter(row => row[index("In Roster")] === true).length,
      selected.filter(row => row[index("ChatGPT")] === "Enterprise License").length,
      selected.filter(row => row[index("Copilot")] === "Premium License").length,
      selected.filter(row => row[index("AI License")] === "AI License").length,
    ];
    const actual = ["B1", "B5", "D5", "F5"].map(cell => summarySheet.getRange(cell).values[0][0]);
    if (JSON.stringify(actual) !== JSON.stringify(expected)) throw new Error(`Filter mismatch ${JSON.stringify({selection, expected, actual})}`);
  }
  summarySheet.getRange("I2:I6").values = fields.map(() => ["All"]);
  workbook.recalculate();
  console.log(`Passed ${cases.length} filter scenarios against independent row counts.`);
  if (headers.includes("new_chatgpt_user")) {
    const inactiveSheet = workbook.worksheets.getItem("ChatGPT Inactive 4 Weeks");
    const expectedEmails = rows.filter(row => row[index("ChatGPT")] === "Enterprise License" &&
      row[index("new_chatgpt_user")] === 0 && row[index("No Activity in 4 Weeks (ChatGPT)")] === true)
      .map(row => row[index("Employee Email")]).sort();
    const actualEmails = inactiveSheet.getRange(`B2:B${expectedEmails.length + 1}`).values.flat().sort();
    if (JSON.stringify(expectedEmails) !== JSON.stringify(actualEmails)) throw new Error("Inactive list mismatch");
    console.log(`Verified ${actualEmails.length} established, inactive ChatGPT users in the new worksheet.`);
    const listPreview = await workbook.render({ sheetName: "ChatGPT Inactive 4 Weeks", range: "I1:M10", scale: 1.5, format: "png" });
    require("fs").writeFileSync("output/inactive-chatgpt-preview.png", Buffer.from(await listPreview.arrayBuffer()));
  }
  const preview = await workbook.render({ sheetName: "Summary", range: "A1:I18", scale: 1.5, format: "png" });
  require("fs").writeFileSync("output/summary-preview.png", Buffer.from(await preview.arrayBuffer()));
})();

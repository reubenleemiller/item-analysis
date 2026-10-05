import ExcelJS from "exceljs";
import path from "node:path";

export type Part = "A" | "B";
export type ResponseType = "MC" | "NR" | "WR";

export type Question = {
  part: Part;
  responseType: ResponseType;
  maxPoints: number;
  label?: string;
};

export type Student = { number: string; name: string };

export type ExportPayload = {
  examTitle: string;
  questions: Question[];
  students: Student[];
};

const START_QUESTION_COLUMN = 3; // C
const TEMPLATE_QUESTION_COUNT = 36;
const TEMPLATE_HELPER_COLUMN = START_QUESTION_COLUMN + TEMPLATE_QUESTION_COUNT; // AM
const FIRST_STUDENT_ROW = 30;
const TEMPLATE_STUDENT_CAPACITY = 50;
const STATIC_COLUMN_WIDTHS = [60, 16, 18, 17, 19, 15, 18];
const SECTION_SUMMARY_HEADERS = [
  "Current section results",
  "Maximum",
  "Scored students",
  "Mean points",
  "Mean / maximum",
  "Item maxima",
  "Maximum check",
];
const STATIC_SHEET_TITLES: ReadonlyArray<readonly [string, string]> = [
  ["A1", "Item analysis with partial credit"],
  ["A3", "Enter each section maximum in column B, then each item maximum in row 13. Enter earned points in the score-entry rows. Decimals are allowed. Blank = ungraded; 0 = no credit."],
  ["A8", "Section totals update from the question maxima below."],
  ["A9", "Configured items"],
  ["C9", "Total maximum"],
  ["E9", "Complete students"],
  ["A10", "Question setup"],
  ["A12", "Question analysis"],
  ["A13", "Maximum points (input)"],
  ["A14", "Response type"],
  ["A15", "Scored students"],
  ["A16", "Mean points"],
  ["A17", "Facility (mean / max) - Percent Correct"],
  ["A18", "Full-credit count"],
  ["A19", "Partial-credit count"],
  ["A20", "Zero-credit count"],
  ["A21", "Ungraded students"],
  ["A22", "Score SD (sample)"],
  ["A23", "Item–rest correlation - Percent Upper Group Correct"],
  ["A24", "Invalid entries"],
  ["A25", "Status"],
  ["A26", "Conclusion"],
  ["A28", "Score entry (inputs)"],
  ["A29", "Student number"],
  ["B29", "Student name"],
];

function columnLetter(column: number) {
  let result = "";
  let value = column;
  while (value > 0) {
    const remainder = (value - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    value = Math.floor((value - 1) / 26);
  }
  return result;
}

function clone<T>(value: T): T {
  return value ? structuredClone(value) : value;
}

function labelQuestions(questions: Question[]) {
  const counters: Record<string, number> = {};
  return questions.map((question) => {
    const key = `${question.part}-${question.responseType}`;
    counters[key] = (counters[key] ?? 0) + 1;
    return question.label?.trim() || `${question.part}-${question.responseType}${counters[key]}`;
  });
}

function assignFormula(cell: ExcelJS.Cell, formula: string) {
  cell.value = { formula };
}

function clearCell(cell: ExcelJS.Cell) {
  cell.value = null;
}

function addSectionBorder(sheet: ExcelJS.Worksheet, startRow: number, endRow: number, startColumn: number, endColumn: number) {
  const border = { style: "medium" as const, color: { argb: "FF000000" } };
  const applyEdge = (row: number, column: number, edge: "top" | "bottom" | "left" | "right") => {
    const cell = sheet.getCell(row, column);
    cell.border = { ...cell.border, [edge]: border };
  };
  for (let column = startColumn; column <= endColumn; column += 1) {
    applyEdge(startRow, column, "top");
    applyEdge(endRow, column, "bottom");
  }
  for (let row = startRow; row <= endRow; row += 1) {
    applyEdge(row, startColumn, "left");
    applyEdge(row, endColumn, "right");
  }
}

export async function createItemAnalysisWorkbook(payload: ExportPayload) {
  const questions = payload.questions
    .filter((question) => Number.isFinite(question.maxPoints) && question.maxPoints > 0)
    .sort((a, b) => a.part.localeCompare(b.part));

  if (!questions.length) throw new Error("Add at least one question with a positive maximum score.");
  if (questions.length > 150) throw new Error("This template supports up to 150 questions per export.");

  const students = payload.students
    .filter((student) => student.number.trim() || student.name.trim())
    .map((student) => ({ number: String(student.number).trim(), name: String(student.name).trim() }));

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(path.join(process.cwd(), "public/templates/item-analysis-template.xlsx"));
  const sheet = workbook.getWorksheet("Item Analysis");
  if (!sheet) throw new Error("The packaged Item Analysis template is unavailable.");

  const summaryHeaderStyles = SECTION_SUMMARY_HEADERS.map((_, index) => clone(sheet.getCell(5, index + 1).style));
  // These summary/setup rows are static. Preserve their complete reference styling,
  // including the intentionally placed outer edges, after the dynamic item columns
  // replace the template's fixed question area.
  const staticSummaryStyles = Array.from({ length: 6 }, (_, rowOffset) =>
    Array.from({ length: 7 }, (_, columnOffset) => clone(sheet.getCell(rowOffset + 5, columnOffset + 1).style)),
  );
  const sectionMaximumInputStyles = [6, 7].map((row) => clone(sheet.getCell(row, 2).style));
  const itemColumnStyles = Array.from({ length: 100 }, (_, index) => clone(sheet.getCell(index + 1, START_QUESTION_COLUMN).style));
  const helperColumnStyles = [0, 1, 2].map((offset) =>
    Array.from({ length: 100 }, (_, index) => clone(sheet.getCell(index + 1, TEMPLATE_HELPER_COLUMN + offset).style)),
  );
  const itemWidth = sheet.getColumn(START_QUESTION_COLUMN).width;
  const helperWidths = [0, 1, 2].map((offset) => sheet.getColumn(TEMPLATE_HELPER_COLUMN + offset).width);

  // Remove the fixed 36-question area and its three trailing total columns. The replacement is sized to the test.
  sheet.spliceColumns(START_QUESTION_COLUMN, TEMPLATE_QUESTION_COUNT + 3);
  const questionCount = questions.length;
  const firstQuestion = columnLetter(START_QUESTION_COLUMN);
  const lastQuestionColumn = START_QUESTION_COLUMN + questionCount - 1;
  const lastQuestion = columnLetter(lastQuestionColumn);
  const totalColumn = columnLetter(lastQuestionColumn + 1);
  const percentColumn = columnLetter(lastQuestionColumn + 2);
  const invalidColumn = columnLetter(lastQuestionColumn + 3);

  for (let column = START_QUESTION_COLUMN; column <= lastQuestionColumn; column += 1) {
    sheet.getColumn(column).width = itemWidth;
    for (let row = 1; row <= itemColumnStyles.length; row += 1) {
      sheet.getCell(row, column).style = clone(itemColumnStyles[row - 1]);
    }
  }
  for (let offset = 0; offset < 3; offset += 1) {
    const column = lastQuestionColumn + 1 + offset;
    sheet.getColumn(column).width = helperWidths[offset];
    for (let row = 1; row <= helperColumnStyles[offset].length; row += 1) {
      sheet.getCell(row, column).style = clone(helperColumnStyles[offset][row - 1]);
    }
  }
  for (let rowOffset = 0; rowOffset < staticSummaryStyles.length; rowOffset += 1) {
    for (let columnOffset = 0; columnOffset < staticSummaryStyles[rowOffset].length; columnOffset += 1) {
      sheet.getCell(rowOffset + 5, columnOffset + 1).style = clone(staticSummaryStyles[rowOffset][columnOffset]);
    }
  }
  // The summary and setup panels end at column G. The question/total columns
  // below them must not inherit any of the template's stale panel borders.
  for (let row = 5; row <= 10; row += 1) {
    for (let column = 8; column <= lastQuestionColumn + 3; column += 1) {
      sheet.getCell(row, column).border = {};
    }
  }
  for (let index = 0; index < SECTION_SUMMARY_HEADERS.length; index += 1) {
    const cell = sheet.getCell(5, index + 1);
    cell.value = SECTION_SUMMARY_HEADERS[index];
    cell.style = clone(summaryHeaderStyles[index]);
    cell.alignment = { ...cell.alignment, vertical: "middle", wrapText: false };
  }
  STATIC_COLUMN_WIDTHS.forEach((width, index) => { sheet.getColumn(index + 1).width = width; });
  sheet.getRow(5).height = 18;
  for (const address of ["A3", "A17", "A23"]) {
    const cell = sheet.getCell(address);
    cell.alignment = { ...cell.alignment, vertical: "top", wrapText: false };
  }
  for (const [address, title] of STATIC_SHEET_TITLES) {
    sheet.getCell(address).value = title;
  }

  if (!students.length) throw new Error("Import at least one student before exporting.");
  const scoreRowStyles = Array.from({ length: lastQuestionColumn + 3 }, (_, index) =>
    clone(sheet.getCell(FIRST_STUDENT_ROW + TEMPLATE_STUDENT_CAPACITY - 1, index + 1).style),
  );
  const scoreRowHeight = sheet.getRow(FIRST_STUDENT_ROW + TEMPLATE_STUDENT_CAPACITY - 1).height;
  const requiredRows = students.length;
  const rowDifference = requiredRows - TEMPLATE_STUDENT_CAPACITY;
  if (rowDifference > 0) {
    sheet.spliceRows(FIRST_STUDENT_ROW + TEMPLATE_STUDENT_CAPACITY, 0, ...Array.from({ length: rowDifference }, () => []));
  } else if (rowDifference < 0) {
    sheet.spliceRows(FIRST_STUDENT_ROW + requiredRows, -rowDifference);
  }
  const lastStudentRow = FIRST_STUDENT_ROW + requiredRows - 1;

  // Restore score-entry row styles after resizing the roster area of the fixed template.
  for (let row = FIRST_STUDENT_ROW; row <= lastStudentRow; row += 1) {
    sheet.getRow(row).height = scoreRowHeight;
    for (let column = 1; column <= lastQuestionColumn + 3; column += 1) {
      sheet.getCell(row, column).style = clone(scoreRowStyles[column - 1]);
    }
  }

  const labels = labelQuestions(questions);
  const partA = questions.filter((question) => question.part === "A");
  const partB = questions.filter((question) => question.part === "B");
  const questionColumns = questions.map((_, index) => columnLetter(START_QUESTION_COLUMN + index));
  const partMaximumFormula = (part: Part) => questionColumns
    .map((column) => `IF(LEFT(${column}$12,2)="${part}-",${column}$13,0)`)
    .join("+");
  const partMeanFormula = (part: Part) => questionColumns
    .map((column) => `IF(LEFT(${column}$12,2)="${part}-",SUM(${column}${FIRST_STUDENT_ROW}:${column}${lastStudentRow}),0)`)
    .join("+");

  sheet.getCell("A2").value = payload.examTitle.trim() || "Item Analysis";
  sheet.getCell("A6").value = partA.length ? "Part A" : "";
  sheet.getCell("A7").value = partB.length ? "Part B" : "";
  for (const row of [6, 7]) {
    const part = row === 6 ? "A" : "B";
    if (!(part === "A" ? partA : partB).length) {
      for (let column = 2; column <= 7; column += 1) clearCell(sheet.getCell(row, column));
      continue;
    }
    const maximumCell = sheet.getCell(row, 2);
    clearCell(maximumCell);
    maximumCell.style = clone(sectionMaximumInputStyles[row - 6]);
    assignFormula(maximumCell, partMaximumFormula(part));
    assignFormula(sheet.getCell(row, 3), `COUNT(${totalColumn}${FIRST_STUDENT_ROW}:${totalColumn}${lastStudentRow})`);
    assignFormula(sheet.getCell(row, 4), `IF(C${row}=0,"",(${partMeanFormula(part)})/C${row})`);
    assignFormula(sheet.getCell(row, 5), `IF(C${row}=0,"",IF(B${row}>0,D${row}/B${row},AVERAGE(${percentColumn}${FIRST_STUDENT_ROW}:${percentColumn}${lastStudentRow})))`);
    assignFormula(sheet.getCell(row, 6), partMaximumFormula(part));
    assignFormula(sheet.getCell(row, 7), `IF(F${row}=B${row},"Matches","Review maxima")`);
  }
  assignFormula(sheet.getCell("B9"), `COUNTIF(${firstQuestion}13:${lastQuestion}13,">0")`);
  assignFormula(sheet.getCell("D9"), `SUM(${firstQuestion}13:${lastQuestion}13)`);
  assignFormula(sheet.getCell("F9"), `COUNT(${totalColumn}${FIRST_STUDENT_ROW}:${totalColumn}${lastStudentRow})`);
  assignFormula(sheet.getCell("B10"), `IF(B9=0,"Add question maxima","Ready")`);

  for (let offset = 0; offset < questionCount; offset += 1) {
    const column = START_QUESTION_COLUMN + offset;
    const letter = columnLetter(column);
    const question = questions[offset];
    sheet.getCell(12, column).value = labels[offset];
    sheet.getCell(13, column).value = question.maxPoints;
    sheet.getCell(14, column).value = question.responseType;
    assignFormula(sheet.getCell(15, column), `IF(OR(NOT(ISNUMBER(${letter}$13)),${letter}$13<=0,${letter}$24>0),"",COUNTIFS(${letter}$${FIRST_STUDENT_ROW}:${letter}$${lastStudentRow},">=0",${letter}$${FIRST_STUDENT_ROW}:${letter}$${lastStudentRow},"<="&${letter}$13,${letter}$${FIRST_STUDENT_ROW}:${letter}$${lastStudentRow},"<>",$A$${FIRST_STUDENT_ROW}:$A$${lastStudentRow},"<>"))`);
    assignFormula(sheet.getCell(16, column), `IF(OR(OR(NOT(ISNUMBER(${letter}$13)),${letter}$13<=0,${letter}$24>0),${letter}$15=0),"",AVERAGEIFS(${letter}$${FIRST_STUDENT_ROW}:${letter}$${lastStudentRow},$A$${FIRST_STUDENT_ROW}:$A$${lastStudentRow},"<>",${letter}$${FIRST_STUDENT_ROW}:${letter}$${lastStudentRow},">=0",${letter}$${FIRST_STUDENT_ROW}:${letter}$${lastStudentRow},"<="&${letter}$13))`);
    assignFormula(sheet.getCell(17, column), `IF(OR(OR(NOT(ISNUMBER(${letter}$13)),${letter}$13<=0,${letter}$24>0),${letter}$15=0),"",${letter}16/${letter}$13)`);
    assignFormula(sheet.getCell(18, column), `IF(OR(NOT(ISNUMBER(${letter}$13)),${letter}$13<=0,${letter}$24>0),"",COUNTIFS(${letter}$${FIRST_STUDENT_ROW}:${letter}$${lastStudentRow},${letter}$13,$A$${FIRST_STUDENT_ROW}:$A$${lastStudentRow},"<>"))`);
    assignFormula(sheet.getCell(19, column), `IF(OR(NOT(ISNUMBER(${letter}$13)),${letter}$13<=0,${letter}$24>0),"",COUNTIFS(${letter}$${FIRST_STUDENT_ROW}:${letter}$${lastStudentRow},">0",${letter}$${FIRST_STUDENT_ROW}:${letter}$${lastStudentRow},"<"&${letter}$13,$A$${FIRST_STUDENT_ROW}:$A$${lastStudentRow},"<>"))`);
    assignFormula(sheet.getCell(20, column), `IF(OR(NOT(ISNUMBER(${letter}$13)),${letter}$13<=0,${letter}$24>0),"",COUNTIFS(${letter}$${FIRST_STUDENT_ROW}:${letter}$${lastStudentRow},0,${letter}$${FIRST_STUDENT_ROW}:${letter}$${lastStudentRow},"<>",$A$${FIRST_STUDENT_ROW}:$A$${lastStudentRow},"<>"))`);
    assignFormula(sheet.getCell(21, column), `IF(OR(NOT(ISNUMBER(${letter}$13)),${letter}$13<=0,${letter}$24>0),"",COUNTIF($A$${FIRST_STUDENT_ROW}:$A$${lastStudentRow},"<>")-${letter}15)`);
    assignFormula(sheet.getCell(22, column), `IF(OR(OR(NOT(ISNUMBER(${letter}$13)),${letter}$13<=0,${letter}$24>0),${letter}$15<2),"",STDEV(${letter}$${FIRST_STUDENT_ROW}:${letter}$${lastStudentRow}))`);
    assignFormula(sheet.getCell(23, column), `IF(OR(OR(NOT(ISNUMBER(${letter}$13)),${letter}$13<=0,${letter}$24>0),COUNT(${totalColumn}$${FIRST_STUDENT_ROW}:${totalColumn}$${lastStudentRow})<3),"",IFERROR(CORREL(FILTER(${letter}$${FIRST_STUDENT_ROW}:${letter}$${lastStudentRow},ISNUMBER(${totalColumn}$${FIRST_STUDENT_ROW}:${totalColumn}$${lastStudentRow})),FILTER(${totalColumn}$${FIRST_STUDENT_ROW}:${totalColumn}$${lastStudentRow}-${letter}$${FIRST_STUDENT_ROW}:${letter}$${lastStudentRow},ISNUMBER(${totalColumn}$${FIRST_STUDENT_ROW}:${totalColumn}$${lastStudentRow}))),"n.a."))`);
    assignFormula(sheet.getCell(24, column), `IF(OR(NOT(ISNUMBER(${letter}$13)),${letter}$13<=0),COUNTA(${letter}$${FIRST_STUDENT_ROW}:${letter}$${lastStudentRow}),COUNTA(${letter}$${FIRST_STUDENT_ROW}:${letter}$${lastStudentRow})-COUNTIFS(${letter}$${FIRST_STUDENT_ROW}:${letter}$${lastStudentRow},">=0",${letter}$${FIRST_STUDENT_ROW}:${letter}$${lastStudentRow},"<="&${letter}$13,${letter}$${FIRST_STUDENT_ROW}:${letter}$${lastStudentRow},"<>",$A$${FIRST_STUDENT_ROW}:$A$${lastStudentRow},"<>"))`);
    assignFormula(sheet.getCell(25, column), `IF(${letter}24>0,"Check scores",IF(OR(NOT(ISNUMBER(${letter}$13)),${letter}$13<=0),"Set maximum",IF(${letter}$15=0,"Enter scores","Ready")))`);
    sheet.getCell(29, column).value = labels[offset];
  }

  sheet.getCell(29, lastQuestionColumn + 1).value = "Complete total";
  sheet.getCell(29, lastQuestionColumn + 2).value = "Percent";
  sheet.getCell(29, lastQuestionColumn + 3).value = "Invalid scores";
  [16, 12, 16].forEach((width, offset) => {
    const column = lastQuestionColumn + 1 + offset;
    const staticWidth = STATIC_COLUMN_WIDTHS[column - 1] ?? 0;
    sheet.getColumn(column).width = Math.max(width, staticWidth);
  });

  for (let row = FIRST_STUDENT_ROW; row <= lastStudentRow; row += 1) {
    const student = students[row - FIRST_STUDENT_ROW];
    sheet.getCell(row, 1).value = student?.number ?? "";
    sheet.getCell(row, 2).value = student?.name ?? "";
    for (let column = START_QUESTION_COLUMN; column <= lastQuestionColumn; column += 1) {
      const letter = columnLetter(column);
      sheet.getCell(row, column).value = null;
      sheet.getCell(row, column).numFmt = "0.##";
      sheet.getCell(row, column).dataValidation = {
        type: "custom",
        allowBlank: true,
        formulae: [`AND(OR(${letter}${row}="",ISNUMBER(${letter}${row})),${letter}${row}>=0,${letter}${row}<=${letter}$13)`],
        showErrorMessage: true,
        errorTitle: "Invalid score",
        error: "Enter a number from 0 to the question maximum, or leave the cell blank.",
      };
    }
    const scoreRange = `${firstQuestion}${row}:${lastQuestion}${row}`;
    const maximaRange = `$${firstQuestion}$13:$${lastQuestion}$13`;
    assignFormula(sheet.getCell(row, lastQuestionColumn + 1), `IF(OR($A${row}="",COUNT(${scoreRange})<>${questionCount},${invalidColumn}${row}>0),"",SUM(${scoreRange}))`);
    assignFormula(sheet.getCell(row, lastQuestionColumn + 2), `IF(ISNUMBER(${totalColumn}${row}),${totalColumn}${row}/$D$9,"")`);
    assignFormula(sheet.getCell(row, lastQuestionColumn + 3), `IF(AND($A${row}="",COUNTA(${scoreRange})=0),"",IF($A${row}="",COUNTA(${scoreRange}),COUNTA(${scoreRange})-SUMPRODUCT(--ISNUMBER(${scoreRange}),--ISNUMBER(${maximaRange}),--(${maximaRange}>0),--(${scoreRange}>=0),--(${scoreRange}<=${maximaRange}))))`);
  }

  addSectionBorder(sheet, 12, 26, 1, lastQuestionColumn);
  addSectionBorder(sheet, 29, lastStudentRow, 1, lastQuestionColumn + 3);

  // The original workbook contains a source/class-list sheet. This export keeps
  // only the self-contained Item Analysis worksheet used by the app.
  for (const worksheet of workbook.worksheets.filter((worksheet) => worksheet.id !== sheet.id)) {
    workbook.removeWorksheet(worksheet.id);
  }

  // Start new Google Sheets at the summary instead of at the score-entry grid.
  // Avoid frozen panes so the import does not create a divider that obscures it.
  sheet.views = [{ state: "normal", activeCell: "A1" }];
  return workbook.xlsx.writeBuffer();
}

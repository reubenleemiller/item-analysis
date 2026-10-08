import ExcelJS from "exceljs";
import path from "node:path";

export type ResponseType = "MC" | "NR" | "WR";

export type Question = {
  part: string;
  responseType: ResponseType;
  maxPoints: number;
  label?: string;
};

export type Student = { number: string; name: string };

export type ExportPayload = {
  examTitle: string;
  questions: Question[];
  students: Student[];
  options?: {
    includePartTotals?: boolean;
    calculateCorrelationByPart?: boolean;
    calculateStatisticsByPart?: boolean;
  };
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
const STATIC_SHEET_TITLES: ReadonlyArray<readonly [number, number, string]> = [
  [1, 1, "Item analysis with partial credit"],
  [3, 1, "Enter each section maximum in column B, then each item maximum in row 13. Enter earned points in the score-entry rows. Decimals are allowed. Blank = ungraded; 0 = no credit."],
  [8, 1, "Section totals update from the question maxima below."],
  [9, 1, "Configured items"],
  [9, 3, "Total maximum"],
  [9, 5, "Complete students"],
  [10, 1, "Question setup"],
  [12, 1, "Question analysis"],
  [13, 1, "Maximum points (input)"],
  [14, 1, "Response type"],
  [15, 1, "Scored students"],
  [16, 1, "Mean points"],
  [17, 1, "Facility (mean / max) - Percent Correct"],
  [18, 1, "Full-credit count"],
  [19, 1, "Partial-credit count"],
  [20, 1, "Zero-credit count"],
  [21, 1, "Ungraded students"],
  [22, 1, "Score SD (sample)"],
  [23, 1, "Item–rest correlation - Percent Upper Group Correct"],
  [24, 1, "Invalid entries"],
  [25, 1, "Status"],
  [26, 1, "Conclusion"],
  [28, 1, "Score entry (inputs)"],
  [29, 1, "Student number"],
  [29, 2, "Student name"],
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
    .map((question) => ({ ...question, part: String(question.part || "").trim() || "Part 1" }));

  if (!questions.length) throw new Error("Add at least one question with a positive maximum score.");
  if (questions.length > 150) throw new Error("This template supports up to 150 questions per export.");
  const parts = [...new Set(questions.map((question) => question.part))];
  const includePartTotals = payload.options?.includePartTotals ?? true;
  const calculateCorrelationByPart = payload.options?.calculateCorrelationByPart ?? false;
  const calculateStatisticsByPart = payload.options?.calculateStatisticsByPart ?? false;

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

  // The reference workbook reserves two part-summary rows. Add rows only when
  // needed, so every distinct part selected in the builder has its own summary.
  const extraPartRows = Math.max(0, parts.length - 2);
  if (extraPartRows) sheet.spliceRows(8, 0, ...Array.from({ length: extraPartRows }, () => []));
  const rowOf = (baseRow: number) => baseRow >= 8 ? baseRow + extraPartRows : baseRow;
  const firstStudentRow = rowOf(FIRST_STUDENT_ROW);
  const row12 = rowOf(12);
  const row13 = rowOf(13);
  const row14 = rowOf(14);
  const row15 = rowOf(15);
  const row16 = rowOf(16);
  const row17 = rowOf(17);
  const row18 = rowOf(18);
  const row19 = rowOf(19);
  const row20 = rowOf(20);
  const row21 = rowOf(21);
  const row22 = rowOf(22);
  const row23 = rowOf(23);
  const row24 = rowOf(24);
  const row25 = rowOf(25);
  const row26 = rowOf(26);
  const row29 = rowOf(29);

  // Remove the fixed 36-question area and its three trailing total columns. The replacement is sized to the test.
  sheet.spliceColumns(START_QUESTION_COLUMN, TEMPLATE_QUESTION_COUNT + 3);
  const questionCount = questions.length;
  const firstQuestion = columnLetter(START_QUESTION_COLUMN);
  const lastQuestionColumn = START_QUESTION_COLUMN + questionCount - 1;
  const lastQuestion = columnLetter(lastQuestionColumn);
  const totalColumn = columnLetter(lastQuestionColumn + 1);
  const partTotalColumns = includePartTotals
    ? parts.map((_, index) => columnLetter(lastQuestionColumn + 2 + index))
    : [];
  const percentColumn = columnLetter(lastQuestionColumn + 2 + partTotalColumns.length);
  const invalidColumn = columnLetter(lastQuestionColumn + 3 + partTotalColumns.length);
  const finalColumn = lastQuestionColumn + 3 + partTotalColumns.length;

  for (let column = START_QUESTION_COLUMN; column <= lastQuestionColumn; column += 1) {
    sheet.getColumn(column).width = itemWidth;
    for (let baseRow = 1; baseRow <= itemColumnStyles.length; baseRow += 1) {
      sheet.getCell(rowOf(baseRow), column).style = clone(itemColumnStyles[baseRow - 1]);
    }
  }
  const helperStyleIndexes = [0, ...partTotalColumns.map(() => 0), 1, 2];
  for (let offset = 0; offset < helperStyleIndexes.length; offset += 1) {
    const column = lastQuestionColumn + 1 + offset;
    const styleIndex = helperStyleIndexes[offset];
    sheet.getColumn(column).width = helperWidths[styleIndex];
    for (let baseRow = 1; baseRow <= helperColumnStyles[styleIndex].length; baseRow += 1) {
      sheet.getCell(rowOf(baseRow), column).style = clone(helperColumnStyles[styleIndex][baseRow - 1]);
    }
  }
  for (let rowOffset = 0; rowOffset < staticSummaryStyles.length; rowOffset += 1) {
    for (let columnOffset = 0; columnOffset < staticSummaryStyles[rowOffset].length; columnOffset += 1) {
      const baseRow = rowOffset + 5;
      sheet.getCell(rowOf(baseRow), columnOffset + 1).style = clone(staticSummaryStyles[rowOffset][columnOffset]);
    }
  }
  // The summary and setup panels end at column G. The question/total columns
  // below them must not inherit any of the template's stale panel borders.
  for (let row = 5; row <= 10; row += 1) {
    for (let column = 8; column <= finalColumn; column += 1) {
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
  for (const baseRow of [3, 17, 23]) {
    const cell = sheet.getCell(rowOf(baseRow), 1);
    cell.alignment = { ...cell.alignment, vertical: "top", wrapText: false };
  }
  for (const [baseRow, column, title] of STATIC_SHEET_TITLES) {
    sheet.getCell(rowOf(baseRow), column).value = title;
  }

  if (!students.length) throw new Error("Import at least one student before exporting.");
  const scoreRowStyles = Array.from({ length: finalColumn }, (_, index) =>
    clone(sheet.getCell(firstStudentRow + TEMPLATE_STUDENT_CAPACITY - 1, index + 1).style),
  );
  const scoreRowHeight = sheet.getRow(firstStudentRow + TEMPLATE_STUDENT_CAPACITY - 1).height;
  const requiredRows = students.length;
  const rowDifference = requiredRows - TEMPLATE_STUDENT_CAPACITY;
  if (rowDifference > 0) {
    sheet.spliceRows(firstStudentRow + TEMPLATE_STUDENT_CAPACITY, 0, ...Array.from({ length: rowDifference }, () => []));
  } else if (rowDifference < 0) {
    sheet.spliceRows(firstStudentRow + requiredRows, -rowDifference);
  }
  const lastStudentRow = firstStudentRow + requiredRows - 1;

  // Restore score-entry row styles after resizing the roster area of the fixed template.
  for (let row = firstStudentRow; row <= lastStudentRow; row += 1) {
    sheet.getRow(row).height = scoreRowHeight;
    for (let column = 1; column <= finalColumn; column += 1) {
      sheet.getCell(row, column).style = clone(scoreRowStyles[column - 1]);
    }
  }

  const labels = labelQuestions(questions);
  const questionColumns = questions.map((_, index) => columnLetter(START_QUESTION_COLUMN + index));
  const partQuestionColumns = (part: string) => questions
    .map((question, index) => question.part === part ? questionColumns[index] : "")
    .filter(Boolean);
  const partMaximumFormula = (part: string) => partQuestionColumns(part)
    .map((column) => `${column}$${row13}`)
    .join("+") || "0";
  const scoreRange = (column: string) => `${column}${firstStudentRow}:${column}${lastStudentRow}`;
  const scoreSumFormula = (columns: string[]) => columns.map(scoreRange).join("+") || "0";
  const completeCohortCondition = (columns: string[]) => [
    `($A$${firstStudentRow}:$A$${lastStudentRow}<>"")`,
    ...columns.flatMap((column) => [`ISNUMBER(${scoreRange(column)})`, `(${scoreRange(column)}>=0)`, `(${scoreRange(column)}<=${column}$${row13})`]),
  ].join("*");
  const cohortCountFormula = (columns: string[]) => `SUMPRODUCT(--(${completeCohortCondition(columns)}))`;
  const cohortPointsFormula = (scoreColumns: string[], cohortColumns: string[]) =>
    `SUMPRODUCT((${scoreSumFormula(scoreColumns)})*--(${completeCohortCondition(cohortColumns)}))`;

  sheet.getCell("A2").value = payload.examTitle.trim() || "Item Analysis";
  for (let index = 0; index < parts.length; index += 1) {
    const row = 6 + index;
    const part = parts[index];
    for (let column = 1; column <= 7; column += 1) {
      sheet.getCell(row, column).style = clone(staticSummaryStyles[Math.min(index + 1, 2)][column - 1]);
    }
    sheet.getCell(row, 1).value = part;
    const maximumCell = sheet.getCell(row, 2);
    const summaryCohortColumns = calculateStatisticsByPart ? partQuestionColumns(part) : questionColumns;
    clearCell(maximumCell);
    maximumCell.style = clone(sectionMaximumInputStyles[Math.min(index, 1)]);
    assignFormula(maximumCell, partMaximumFormula(part));
    assignFormula(sheet.getCell(row, 3), cohortCountFormula(summaryCohortColumns));
    assignFormula(sheet.getCell(row, 4), `IF(C${row}=0,"",${cohortPointsFormula(partQuestionColumns(part), summaryCohortColumns)}/C${row})`);
    assignFormula(sheet.getCell(row, 5), `IF(C${row}=0,"",IF(B${row}>0,D${row}/B${row},AVERAGE(${percentColumn}${firstStudentRow}:${percentColumn}${lastStudentRow})))`);
    assignFormula(sheet.getCell(row, 6), partMaximumFormula(part));
    assignFormula(sheet.getCell(row, 7), `IF(F${row}=B${row},"Matches","Review maxima")`);
  }
  assignFormula(sheet.getCell(rowOf(9), 2), `COUNTIF(${firstQuestion}${row13}:${lastQuestion}${row13},">0")`);
  assignFormula(sheet.getCell(rowOf(9), 4), `SUM(${firstQuestion}${row13}:${lastQuestion}${row13})`);
  assignFormula(sheet.getCell(rowOf(9), 6), `COUNT(${totalColumn}${firstStudentRow}:${totalColumn}${lastStudentRow})`);
  assignFormula(sheet.getCell(rowOf(10), 2), `IF(B${rowOf(9)}=0,"Add question maxima","Ready")`);
  sheet.getCell(row15, 1).value = calculateStatisticsByPart ? "Scored students (complete part)" : "Scored students (complete test)";
  sheet.getCell(row21, 1).value = calculateStatisticsByPart ? "Incomplete part students" : "Incomplete test students";
  sheet.getCell(row23, 1).value = calculateCorrelationByPart ? "Item–rest correlation - within part" : "Item–rest correlation - whole test";

  for (let offset = 0; offset < questionCount; offset += 1) {
    const column = START_QUESTION_COLUMN + offset;
    const letter = columnLetter(column);
    const question = questions[offset];
    sheet.getCell(row12, column).value = labels[offset];
    sheet.getCell(row13, column).value = question.maxPoints;
    sheet.getCell(row14, column).value = question.responseType;
    const statisticsColumns = calculateStatisticsByPart ? partQuestionColumns(question.part) : questionColumns;
    const statisticsCondition = completeCohortCondition(statisticsColumns);
    const statisticsCount = cohortCountFormula(statisticsColumns);
    const correlationColumns = calculateCorrelationByPart ? partQuestionColumns(question.part) : questionColumns;
    const correlationCondition = completeCohortCondition(correlationColumns);
    const correlationCount = cohortCountFormula(correlationColumns);
    const correlationTotal = scoreSumFormula(correlationColumns);
    assignFormula(sheet.getCell(row15, column), `IF(OR(NOT(ISNUMBER(${letter}$${row13})),${letter}$${row13}<=0),"",${statisticsCount})`);
    assignFormula(sheet.getCell(row16, column), `IF(OR(NOT(ISNUMBER(${letter}$${row13})),${letter}$${row13}<=0,${letter}$${row15}=0),"",SUMPRODUCT(${scoreRange(letter)}*--(${statisticsCondition}))/${letter}$${row15})`);
    assignFormula(sheet.getCell(row17, column), `IF(OR(OR(NOT(ISNUMBER(${letter}$${row13})),${letter}$${row13}<=0),${letter}$${row15}=0),"",${letter}${row16}/${letter}$${row13})`);
    assignFormula(sheet.getCell(row18, column), `IF(OR(NOT(ISNUMBER(${letter}$${row13})),${letter}$${row13}<=0),"",SUMPRODUCT(--(${statisticsCondition}),--(${scoreRange(letter)}=${letter}$${row13})))`);
    assignFormula(sheet.getCell(row19, column), `IF(OR(NOT(ISNUMBER(${letter}$${row13})),${letter}$${row13}<=0),"",SUMPRODUCT(--(${statisticsCondition}),--(${scoreRange(letter)}>0),--(${scoreRange(letter)}<${letter}$${row13})))`);
    assignFormula(sheet.getCell(row20, column), `IF(OR(NOT(ISNUMBER(${letter}$${row13})),${letter}$${row13}<=0),"",SUMPRODUCT(--(${statisticsCondition}),--(${scoreRange(letter)}=0)))`);
    assignFormula(sheet.getCell(row21, column), `IF(OR(NOT(ISNUMBER(${letter}$${row13})),${letter}$${row13}<=0),"",COUNTIF($A$${firstStudentRow}:$A$${lastStudentRow},"<>")-${letter}${row15})`);
    assignFormula(sheet.getCell(row22, column), `IF(OR(OR(NOT(ISNUMBER(${letter}$${row13})),${letter}$${row13}<=0),${letter}$${row15}<2),"",IFERROR(STDEV(FILTER(${scoreRange(letter)},${statisticsCondition})),""))`);
    assignFormula(sheet.getCell(row23, column), `IF(OR(NOT(ISNUMBER(${letter}$${row13})),${letter}$${row13}<=0,${correlationCount}<3),"",IFERROR(CORREL(FILTER(${scoreRange(letter)},${correlationCondition}),FILTER(${correlationTotal}-${scoreRange(letter)},${correlationCondition})),"n.a."))`);
    assignFormula(sheet.getCell(row24, column), `IF(OR(NOT(ISNUMBER(${letter}$${row13})),${letter}$${row13}<=0),COUNTA(${letter}$${firstStudentRow}:${letter}$${lastStudentRow}),COUNTA(${letter}$${firstStudentRow}:${letter}$${lastStudentRow})-COUNTIFS(${letter}$${firstStudentRow}:${letter}$${lastStudentRow},">=0",${letter}$${firstStudentRow}:${letter}$${lastStudentRow},"<="&${letter}$${row13},${letter}$${firstStudentRow}:${letter}$${lastStudentRow},"<>",$A$${firstStudentRow}:$A$${lastStudentRow},"<>"))`);
    assignFormula(sheet.getCell(row25, column), `IF(${letter}${row24}>0,"Check scores",IF(OR(NOT(ISNUMBER(${letter}$${row13})),${letter}$${row13}<=0),"Set maximum",IF(${letter}${row15}=0,"Enter scores","Ready")))`);
    sheet.getCell(row29, column).value = labels[offset];
  }

  sheet.getCell(row29, lastQuestionColumn + 1).value = "Complete total";
  partTotalColumns.forEach((_, index) => { sheet.getCell(row29, lastQuestionColumn + 2 + index).value = `${parts[index]} total`; });
  sheet.getCell(row29, lastQuestionColumn + 2 + partTotalColumns.length).value = "Percent";
  sheet.getCell(row29, lastQuestionColumn + 3 + partTotalColumns.length).value = "Invalid scores";
  [16, ...partTotalColumns.map(() => 16), 12, 16].forEach((width, offset) => {
    const column = lastQuestionColumn + 1 + offset;
    const staticWidth = STATIC_COLUMN_WIDTHS[column - 1] ?? 0;
    sheet.getColumn(column).width = Math.max(width, staticWidth);
  });

  for (let row = firstStudentRow; row <= lastStudentRow; row += 1) {
    const student = students[row - firstStudentRow];
    sheet.getCell(row, 1).value = student?.number ?? "";
    sheet.getCell(row, 2).value = student?.name ?? "";
    for (let column = START_QUESTION_COLUMN; column <= lastQuestionColumn; column += 1) {
      const letter = columnLetter(column);
      sheet.getCell(row, column).value = null;
      sheet.getCell(row, column).numFmt = "0.##";
      sheet.getCell(row, column).dataValidation = {
        type: "custom",
        allowBlank: true,
        formulae: [`AND(OR(${letter}${row}="",ISNUMBER(${letter}${row})),${letter}${row}>=0,${letter}${row}<=${letter}$${row13})`],
        showErrorMessage: true,
        errorTitle: "Invalid score",
        error: "Enter a number from 0 to the question maximum, or leave the cell blank.",
      };
    }
    const scoreRange = `${firstQuestion}${row}:${lastQuestion}${row}`;
    const maximaRange = `$${firstQuestion}$${row13}:$${lastQuestion}$${row13}`;
    assignFormula(sheet.getCell(row, lastQuestionColumn + 1), `IF(OR($A${row}="",COUNT(${scoreRange})<>${questionCount},${invalidColumn}${row}>0),"",SUM(${scoreRange}))`);
    partTotalColumns.forEach((partColumn, index) => {
      const partScores = partQuestionColumns(parts[index]).map((column) => `${column}${row}`).join(",");
      assignFormula(sheet.getCell(row, lastQuestionColumn + 2 + index), `IF($A${row}="","",SUM(${partScores}))`);
    });
    assignFormula(sheet.getCell(row, lastQuestionColumn + 2 + partTotalColumns.length), `IF(ISNUMBER(${totalColumn}${row}),${totalColumn}${row}/$D$${rowOf(9)},"")`);
    assignFormula(sheet.getCell(row, lastQuestionColumn + 3 + partTotalColumns.length), `IF(AND($A${row}="",COUNTA(${scoreRange})=0),"",IF($A${row}="",COUNTA(${scoreRange}),COUNTA(${scoreRange})-SUMPRODUCT(--ISNUMBER(${scoreRange}),--ISNUMBER(${maximaRange}),--(${maximaRange}>0),--(${scoreRange}>=0),--(${scoreRange}<=${maximaRange}))))`);
  }

  addSectionBorder(sheet, row12, row26, 1, lastQuestionColumn);
  addSectionBorder(sheet, row29, lastStudentRow, 1, finalColumn);

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

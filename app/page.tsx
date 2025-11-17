"use client";

import { ChangeEvent, useCallback, useMemo, useState } from "react";
import * as XLSX from "xlsx";
import { SheetCell, SheetGrid } from "./components/SheetGrid";

interface StatusState {
  type: "success" | "error" | "info";
  message: string;
}

const SAMPLE_DATA: SheetCell[][] = [
  ["Employee", "Department", "Hours", "Rate", "Total"],
  ["Anita", "Finance", 32, 45, null],
  ["Rohan", "Marketing", 40, 38, null],
  ["Zoya", "Finance", 28, 52, null],
  ["Karan", "Sales", 45, 41, null],
];

const EXAMPLES = [
  "set cell e2 to c2 * d2",
  "increment column c by 5",
  "rename column e to payroll",
  "fill empty cells in column e with c * d",
  "add column f as sum of c and e",
  "clear column d",
];

export default function Page() {
  const [sheetData, setSheetData] = useState<SheetCell[][] | null>(null);
  const [sheetName, setSheetName] = useState("Sheet1");
  const [fileName, setFileName] = useState<string | null>(null);
  const [instruction, setInstruction] = useState("");
  const [status, setStatus] = useState<StatusState | null>(null);
  const [history, setHistory] = useState<string[]>([]);

  const handleFileUpload = useCallback(async (event: ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) {
      return;
    }

    try {
      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: "array" });
      const firstSheet = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheet];
      const rows = XLSX.utils.sheet_to_json<SheetCell[]>(worksheet, { header: 1, raw: false, defval: "" });
      const normalized = ensureUniformColumns(rows as SheetCell[][]);
      setSheetData(normalized);
      setSheetName(firstSheet);
      setFileName(file.name.replace(/\.xlsx?$/i, ""));
      setStatus({ type: "success", message: `Loaded ${file.name}` });
      setHistory([]);
    } catch (error) {
      console.error(error);
      setStatus({ type: "error", message: "Failed to read Excel file." });
    }
  }, []);

  const handleSampleLoad = useCallback(() => {
    setSheetData(ensureUniformColumns(SAMPLE_DATA));
    setSheetName("Workforce");
    setFileName("sample-workforce");
    setStatus({ type: "info", message: "Sample sheet loaded." });
    setHistory([]);
  }, []);

  const handleDownload = useCallback(() => {
    if (!sheetData) {
      setStatus({ type: "error", message: "Nothing to export yet." });
      return;
    }

    const workbook = XLSX.utils.book_new();
    const worksheet = XLSX.utils.aoa_to_sheet(sheetData);
    XLSX.utils.book_append_sheet(workbook, worksheet, sheetName || "Sheet1");
    const downloadName = `${fileName ?? "excel-agent"}-updated.xlsx`;
    XLSX.writeFile(workbook, downloadName);
    setStatus({ type: "success", message: `Exported ${downloadName}` });
  }, [fileName, sheetData, sheetName]);

  const handleInstructionRun = useCallback(() => {
    if (!sheetData) {
      setStatus({ type: "error", message: "Upload a sheet before running instructions." });
      return;
    }
    const text = instruction.trim();
    if (!text) {
      setStatus({ type: "error", message: "Enter an instruction." });
      return;
    }

    try {
      const result = applyInstruction(sheetData, text);
      setSheetData(result.data);
      setStatus({ type: "success", message: result.message });
      setHistory((prev) => [result.message, ...prev]);
      setInstruction("");
    } catch (error) {
      const message = error instanceof Error ? error.message : "Could not understand instruction.";
      setStatus({ type: "error", message });
    }
  }, [instruction, sheetData]);

  const instructionPlaceholder = useMemo(() => {
    if (!sheetData) {
      return "Try: load sample data and ask me to add totals";
    }
    return "Example: set cell e2 to c2 * d2";
  }, [sheetData]);

  return (
    <main className="mx-auto flex min-h-screen max-w-6xl flex-col gap-10 px-6 py-10">
      <header className="flex flex-col gap-4 rounded-3xl bg-gradient-to-r from-primary-500 to-primary-700 px-8 py-10 text-white shadow-2xl">
        <div className="flex flex-col items-start justify-between gap-4 sm:flex-row">
          <div>
            <p className="text-sm uppercase tracking-widest text-primary-100">Excel AI Agent</p>
            <h1 className="text-3xl font-semibold leading-tight sm:text-4xl">
              Update spreadsheets in Hinglish, powered by on-device rules.
            </h1>
          </div>
          <div className="flex items-center gap-3">
            <button
              type="button"
              onClick={handleSampleLoad}
              className="rounded-full bg-white px-4 py-2 text-sm font-semibold text-primary-700 shadow-lg shadow-primary-900/20 transition hover:bg-primary-50"
            >
              Load Sample Sheet
            </button>
            <label className="cursor-pointer rounded-full border border-white/40 bg-white/10 px-4 py-2 text-sm font-semibold backdrop-blur transition hover:bg-white/20">
              Upload Excel
              <input
                type="file"
                accept=".xlsx,.xls"
                onChange={handleFileUpload}
                className="hidden"
              />
            </label>
          </div>
        </div>
        <p className="max-w-2xl text-sm text-primary-100">
          Describe the updates you need using natural language — like &ldquo;column C ko 5 se increase karo&rdquo; — and the agent rewrites your sheet instantly.
        </p>
        {status && (
          <div
            className={`rounded-2xl px-4 py-3 text-sm font-medium shadow-lg backdrop-blur transition ${
              status.type === "success"
                ? "bg-white/20 text-white"
                : status.type === "error"
                ? "bg-red-500/70 text-white"
                : "bg-white/30 text-white"
            }`}
          >
            {status.message}
          </div>
        )}
      </header>

      <section className="grid gap-8 lg:grid-cols-[2fr,1fr]">
        <div className="flex flex-col gap-6">
          {sheetData ? (
            <SheetGrid data={sheetData} sheetName={sheetName} />
          ) : (
            <div className="rounded-3xl border border-dashed border-slate-300 bg-white/70 p-12 text-center text-slate-500 shadow-inner">
              Upload an Excel workbook or load the sample sheet to get started.
            </div>
          )}
          <div className="flex flex-col gap-3 rounded-3xl bg-white p-6 shadow-xl shadow-slate-900/5">
            <label htmlFor="instruction" className="text-sm font-semibold uppercase tracking-widest text-slate-500">
              Natural Language Instruction
            </label>
            <textarea
              id="instruction"
              value={instruction}
              onChange={(event) => setInstruction(event.target.value)}
              placeholder={instructionPlaceholder}
              rows={3}
              className="w-full resize-none rounded-2xl border border-slate-200 bg-slate-50/60 px-4 py-3 text-sm text-slate-700 outline-none transition focus:border-primary-400 focus:bg-white focus:ring-2 focus:ring-primary-200"
            />
            <div className="flex flex-wrap items-center justify-between gap-3">
              <div className="flex flex-wrap gap-2">
                {EXAMPLES.map((example) => (
                  <button
                    key={example}
                    type="button"
                    onClick={() => setInstruction(example)}
                    className="rounded-full border border-slate-200 bg-slate-100 px-3 py-1 text-xs font-medium text-slate-600 transition hover:border-primary-200 hover:bg-primary-50 hover:text-primary-700"
                  >
                    {example}
                  </button>
                ))}
              </div>
              <div className="flex items-center gap-2">
                <button
                  type="button"
                  onClick={handleDownload}
                  className="rounded-full border border-slate-200 bg-white px-4 py-2 text-sm font-semibold text-slate-600 shadow-sm transition hover:border-primary-300 hover:text-primary-600"
                >
                  Export Excel
                </button>
                <button
                  type="button"
                  onClick={handleInstructionRun}
                  className="rounded-full bg-primary-600 px-6 py-2 text-sm font-semibold text-white shadow-lg shadow-primary-900/30 transition hover:bg-primary-500"
                >
                  Run Instruction
                </button>
              </div>
            </div>
          </div>
        </div>
        <aside className="flex h-full flex-col gap-6">
          <div className="rounded-3xl bg-white p-6 shadow-xl shadow-slate-900/5">
            <h2 className="text-base font-semibold text-slate-800">Instruction History</h2>
            <ol className="mt-4 space-y-3 text-sm text-slate-600">
              {history.length === 0 && <li className="text-slate-400">No instructions run yet.</li>}
              {history.map((entry, idx) => (
                <li key={`${entry}-${idx}`} className="rounded-2xl border border-slate-100 bg-slate-50/60 px-3 py-2">
                  {entry}
                </li>
              ))}
            </ol>
          </div>
          <div className="rounded-3xl border border-dashed border-primary-200 bg-primary-50/70 p-6 text-sm text-primary-900 shadow-inner">
            <h3 className="mb-2 text-base font-semibold">Supported actions</h3>
            <ul className="space-y-2">
              <li>Set or update any cell. Example: &ldquo;set cell b4 to 120&rdquo;</li>
              <li>Increase, decrease, multiply, or divide a number column.</li>
              <li>Add calculated columns using addition, subtraction, or difference.</li>
              <li>Fill empty cells with a value or another column expression.</li>
              <li>Rename headers to keep things organized.</li>
            </ul>
          </div>
        </aside>
      </section>
    </main>
  );
}

function applyInstruction(data: SheetCell[][], rawInstruction: string) {
  const instruction = rawInstruction.trim();
  if (!instruction) {
    throw new Error("Instruction cannot be empty.");
  }

  const working = data.map((row) => [...row]);
  const normalized = instruction.toLowerCase();

  const setCellMatch = instruction.match(/(?:set|update)\s+cell\s+([a-z]+)(\d+)\s+(?:to|as)\s+(.+)/i);
  if (setCellMatch) {
    const [, columnLetters, rowNumberRaw, valueExpression] = setCellMatch;
    const columnIndex = columnLettersToIndex(columnLetters);
    const rowIndex = Number(rowNumberRaw) - 1;
    if (Number.isNaN(rowIndex) || rowIndex < 0) {
      throw new Error("Row reference invalid.");
    }

    ensureRowCapacity(working, rowIndex + 1, working[0]?.length ?? columnIndex + 1);

    const evaluatedValue = evaluateExpression(valueExpression, working, rowIndex, columnIndex);
    working[rowIndex][columnIndex] = evaluatedValue;

    return {
      data: ensureUniformColumns(working),
      message: `Set ${columnLetters.toUpperCase()}${rowIndex + 1} to ${String(evaluatedValue)}`,
    };
  }

  const adjustColumnMatch = normalized.match(/(increase|increment|decrease|reduce|multiply|divide)\s+column\s+([a-z]+)\s+(?:by|with)\s+([\w ./*+-]+)/i);
  if (adjustColumnMatch) {
    const [, action, columnLetters, amountRaw] = adjustColumnMatch;
    const columnIndex = columnLettersToIndex(columnLetters);
    ensureColumnCapacity(working, columnIndex + 1);

    const factor = evaluateNumericExpression(amountRaw, working, 1, columnIndex);
    for (let rowIndex = 1; rowIndex < working.length; rowIndex += 1) {
      const current = toNumber(working[rowIndex][columnIndex]);
      if (current === null) {
        continue;
      }
      switch (action) {
        case "increase":
        case "increment":
          working[rowIndex][columnIndex] = roundNumber(current + factor);
          break;
        case "decrease":
        case "reduce":
          working[rowIndex][columnIndex] = roundNumber(current - factor);
          break;
        case "multiply":
          working[rowIndex][columnIndex] = roundNumber(current * factor);
          break;
        case "divide":
          if (factor === 0) {
            throw new Error("Cannot divide by zero.");
          }
          working[rowIndex][columnIndex] = roundNumber(current / factor);
          break;
        default:
          break;
      }
    }

    return {
      data: ensureUniformColumns(working),
      message: `${capitalizeLabel(action)} column ${columnLetters.toUpperCase()} by ${factor}`,
    };
  }

  const addColumnMatch = instruction.match(/add\s+column\s+([a-z]+)\s+(?:as|with)\s+(?:sum|total|difference|value)\s+of\s+([a-z]+)\s+(?:and|minus)\s+([a-z]+)/i);
  if (addColumnMatch) {
    const [, targetLetters, leftLetters, rightLetters] = addColumnMatch;
    const targetIndex = columnLettersToIndex(targetLetters);
    const leftIndex = columnLettersToIndex(leftLetters);
    const rightIndex = columnLettersToIndex(rightLetters);

    ensureColumnCapacity(working, Math.max(targetIndex, leftIndex, rightIndex) + 1);

    const isDifference = /minus/i.test(addColumnMatch[0]);
    for (let rowIndex = 0; rowIndex < working.length; rowIndex += 1) {
      const left = toNumber(working[rowIndex][leftIndex]);
      const right = toNumber(working[rowIndex][rightIndex]);
      if (rowIndex === 0) {
        working[rowIndex][targetIndex] = isDifference ? `${leftLetters.toUpperCase()}-${rightLetters.toUpperCase()}` : `${leftLetters.toUpperCase()}+${rightLetters.toUpperCase()}`;
        continue;
      }
      if (left === null || right === null) {
        working[rowIndex][targetIndex] = "";
        continue;
      }
      working[rowIndex][targetIndex] = roundNumber(isDifference ? left - right : left + right);
    }

    return {
      data: ensureUniformColumns(working),
      message: `Computed column ${targetLetters.toUpperCase()} from ${leftLetters.toUpperCase()} ${isDifference ? "-" : "+"} ${rightLetters.toUpperCase()}`,
    };
  }

  const fillColumnMatch = instruction.match(/fill\s+(?:empty\s+)?cells?\s+(?:in\s+)?column\s+([a-z]+)\s+(?:with|using)\s+(.+)/i);
  if (fillColumnMatch) {
    const [, columnLetters, valueRaw] = fillColumnMatch;
    const columnIndex = columnLettersToIndex(columnLetters);
    ensureColumnCapacity(working, columnIndex + 1);
    for (let rowIndex = 1; rowIndex < working.length; rowIndex += 1) {
      if (working[rowIndex][columnIndex] === "" || working[rowIndex][columnIndex] === null || working[rowIndex][columnIndex] === undefined) {
        working[rowIndex][columnIndex] = evaluateExpression(valueRaw, working, rowIndex, columnIndex);
      }
    }
    return {
      data: ensureUniformColumns(working),
      message: `Filled empty cells in column ${columnLetters.toUpperCase()}`,
    };
  }

  const renameColumnMatch = instruction.match(/rename\s+column\s+([a-z]+)\s+(?:to|as)\s+(.+)/i);
  if (renameColumnMatch) {
    const [, columnLetters, titleRaw] = renameColumnMatch;
    const columnIndex = columnLettersToIndex(columnLetters);
    ensureColumnCapacity(working, columnIndex + 1);
    const newTitle = titleRaw.trim();
    working[0][columnIndex] = capitalizeWords(newTitle);
    return {
      data: ensureUniformColumns(working),
      message: `Renamed column ${columnLetters.toUpperCase()} to ${capitalizeWords(newTitle)}`,
    };
  }

  const clearColumnMatch = instruction.match(/clear\s+column\s+([a-z]+)/i);
  if (clearColumnMatch) {
    const [, columnLetters] = clearColumnMatch;
    const columnIndex = columnLettersToIndex(columnLetters);
    ensureColumnCapacity(working, columnIndex + 1);
    for (let rowIndex = 1; rowIndex < working.length; rowIndex += 1) {
      working[rowIndex][columnIndex] = "";
    }
    return {
      data: ensureUniformColumns(working),
      message: `Cleared column ${columnLetters.toUpperCase()}`,
    };
  }

  throw new Error("Instruction not recognized. Try a simpler sentence.");
}

function ensureUniformColumns(data: SheetCell[][]): SheetCell[][] {
  if (data.length === 0) {
    return [["Column A"]];
  }
  const widest = Math.max(...data.map((row) => row.length));
  return data.map((row) => {
    const next = [...row];
    while (next.length < widest) {
      next.push("");
    }
    return next;
  });
}

function ensureRowCapacity(data: SheetCell[][], requiredRows: number, targetColumns: number) {
  while (data.length < requiredRows) {
    const newRow: SheetCell[] = Array.from({ length: targetColumns }, () => "");
    data.push(newRow);
  }
  ensureColumnCapacity(data, targetColumns);
}

function ensureColumnCapacity(data: SheetCell[][], requiredColumns: number) {
  for (let rowIndex = 0; rowIndex < data.length; rowIndex += 1) {
    const row = data[rowIndex];
    while (row.length < requiredColumns) {
      row.push("");
    }
  }
}

function columnLettersToIndex(letters: string) {
  const clean = letters.trim().toUpperCase();
  let index = 0;
  for (let i = 0; i < clean.length; i += 1) {
    const charCode = clean.charCodeAt(i);
    if (charCode < 65 || charCode > 90) {
      throw new Error(`Invalid column reference: ${letters}`);
    }
    index = index * 26 + (charCode - 64);
  }
  return index - 1;
}

function evaluateExpression(expression: string, data: SheetCell[][], rowIndex: number, columnIndex: number): SheetCell {
  const trimmed = expression.trim();
  if (/^".*"$/.test(trimmed)) {
    return trimmed.slice(1, -1);
  }
  if (/^'.*'$/.test(trimmed)) {
    return trimmed.slice(1, -1);
  }
  if (/^(blank|empty|null|none)$/i.test(trimmed)) {
    return "";
  }

  if (/^[a-z]+\d+$/i.test(trimmed)) {
    const match = trimmed.match(/([a-z]+)(\d+)/i);
    if (match) {
      const targetColumn = columnLettersToIndex(match[1]);
      const targetRow = Number(match[2]) - 1;
      ensureRowCapacity(data, targetRow + 1, Math.max(data[0]?.length ?? 0, targetColumn + 1));
      return data[targetRow][targetColumn] ?? "";
    }
  }

  const numeric = parseFloat(trimmed.replace(/,/g, "."));
  if (!Number.isNaN(numeric)) {
    return numeric;
  }

  const columnFormulaMatch = trimmed.match(/([a-z]+)\s*([*+\/-])\s*([a-z]+)/i);
  if (columnFormulaMatch) {
    const [, leftLetters, operator, rightLetters] = columnFormulaMatch;
    const leftIndex = columnLettersToIndex(leftLetters);
    const rightIndex = columnLettersToIndex(rightLetters);
    const left = toNumber(data[rowIndex]?.[leftIndex]);
    const right = toNumber(data[rowIndex]?.[rightIndex]);
    if (left === null || right === null) {
      return "";
    }
    switch (operator) {
      case "+":
        return roundNumber(left + right);
      case "-":
        return roundNumber(left - right);
      case "*":
        return roundNumber(left * right);
      case "/":
        return roundNumber(right === 0 ? 0 : left / right);
      default:
        return "";
    }
  }

  const cellFormulaMatch = trimmed.match(/([a-z]+)(\d+)\s*([*+\-\/])\s*([a-z]+)(\d+)/i);
  if (cellFormulaMatch) {
    const [, leftLetters, leftRowRaw, operator, rightLetters, rightRowRaw] = cellFormulaMatch;
    const leftIndex = columnLettersToIndex(leftLetters);
    const rightIndex = columnLettersToIndex(rightLetters);
    const leftRow = Number(leftRowRaw) - 1;
    const rightRow = Number(rightRowRaw) - 1;
    ensureRowCapacity(data, Math.max(leftRow, rightRow) + 1, Math.max(leftIndex, rightIndex) + 1);
    const left = toNumber(data[leftRow]?.[leftIndex]);
    const right = toNumber(data[rightRow]?.[rightIndex]);
    if (left === null || right === null) {
      return "";
    }
    switch (operator) {
      case "+":
        return roundNumber(left + right);
      case "-":
        return roundNumber(left - right);
      case "*":
        return roundNumber(left * right);
      case "/":
        return roundNumber(right === 0 ? 0 : left / right);
      default:
        return "";
    }
  }

  const cellNumberMatch = trimmed.match(/([a-z]+)(\d+)\s*([*+\-\/])\s*([\d.]+)/i);
  if (cellNumberMatch) {
    const [, letters, rowRaw, operator, numberRaw] = cellNumberMatch;
    const column = columnLettersToIndex(letters);
    const row = Number(rowRaw) - 1;
    ensureRowCapacity(data, row + 1, column + 1);
    const base = toNumber(data[row]?.[column]);
    const operand = Number(numberRaw);
    if (base === null || Number.isNaN(operand)) {
      return "";
    }
    switch (operator) {
      case "+":
        return roundNumber(base + operand);
      case "-":
        return roundNumber(base - operand);
      case "*":
        return roundNumber(base * operand);
      case "/":
        return roundNumber(operand === 0 ? 0 : base / operand);
      default:
        return "";
    }
  }

  return capitalizeWords(trimmed);
}

function evaluateNumericExpression(expression: string, data: SheetCell[][], rowIndex: number, columnIndex: number): number {
  const result = evaluateExpression(expression, data, rowIndex, columnIndex);
  const value = toNumber(result);
  if (value === null) {
    throw new Error(`Expected a numeric value from "${expression}".`);
  }
  return value;
}

function toNumber(value: SheetCell): number | null {
  if (value === null || value === undefined || value === "") {
    return null;
  }
  if (typeof value === "number") {
    return value;
  }
  const parsed = Number(String(value).replace(/,/g, "."));
  if (Number.isNaN(parsed)) {
    return null;
  }
  return parsed;
}

function roundNumber(value: number) {
  return Math.round((value + Number.EPSILON) * 100) / 100;
}

function capitalizeWord(value: string) {
  return value.charAt(0).toUpperCase() + value.slice(1).toLowerCase();
}

function capitalizeWords(value: string) {
  return value
    .split(/\s+/)
    .filter(Boolean)
    .map((part) => capitalizeWord(part))
    .join(" ");
}

function capitalizeLabel(action: string) {
  return action.charAt(0).toUpperCase() + action.slice(1);
}

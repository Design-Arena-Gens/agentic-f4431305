"use client";

import clsx from "clsx";
import { memo } from "react";

export type SheetCell = string | number | boolean | null;

interface SheetGridProps {
  data: SheetCell[][];
  sheetName: string;
}

const SheetGridComponent = ({ data, sheetName }: SheetGridProps) => {
  if (!data.length) {
    return (
      <div className="rounded-xl border border-dashed border-slate-300 bg-white/60 p-10 text-center text-slate-500 shadow-sm">
        No data available in {sheetName}.
      </div>
    );
  }

  const columnCount = data[0].length;
  const columnHeaders = Array.from({ length: columnCount }, (_, idx) => indexToColumnLetters(idx));

  return (
    <div className="overflow-hidden rounded-2xl border border-slate-200 shadow-xl shadow-slate-900/5">
      <div className="max-h-[480px] overflow-auto">
        <table className="min-w-full border-collapse bg-white text-sm">
          <thead className="sticky top-0 z-10 bg-slate-900 text-white">
            <tr>
              <th className="sticky left-0 z-20 bg-slate-800 px-3 py-2 text-left font-semibold" aria-label="Row" />
              {columnHeaders.map((col) => (
                <th key={col} className="px-3 py-2 text-left font-semibold tracking-wide">
                  {col}
                </th>
              ))}
            </tr>
            <tr className="bg-slate-800/90 text-slate-200">
              <th className="sticky left-0 z-10 bg-slate-800/90 px-3 py-2 text-left font-medium">
                #
              </th>
              {data[0].map((value, idx) => (
                <th key={`header-${idx}`} className="px-3 py-2 text-left font-medium uppercase tracking-wider">
                  {renderValue(value)}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {data.slice(1).map((row, rowIndex) => (
              <tr
                key={`row-${rowIndex}`}
                className={clsx("border-t border-slate-100", rowIndex % 2 === 0 ? "bg-white" : "bg-slate-50/60")}
              >
                <td className="sticky left-0 bg-white px-3 py-2 font-medium text-slate-400">
                  {rowIndex + 1}
                </td>
                {row.map((cell, colIndex) => (
                  <td key={`cell-${rowIndex}-${colIndex}`} className="px-3 py-2 text-slate-700">
                    {renderValue(cell)}
                  </td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
};

export const SheetGrid = memo(SheetGridComponent);
SheetGrid.displayName = "SheetGrid";

function renderValue(value: SheetCell) {
  if (value === null || value === undefined || value === "") {
    return <span className="text-slate-300">—</span>;
  }
  if (typeof value === "boolean") {
    return value ? "TRUE" : "FALSE";
  }
  return `${value}`;
}

function indexToColumnLetters(index: number) {
  let result = "";
  let current = index;
  while (current >= 0) {
    result = String.fromCharCode((current % 26) + 65) + result;
    current = Math.floor(current / 26) - 1;
  }
  return result;
}

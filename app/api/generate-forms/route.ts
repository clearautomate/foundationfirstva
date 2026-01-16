// app/api/generate-forms/route.ts
import JSZip from "jszip";
import ExcelJS from "exceljs";
import { NextResponse } from "next/server";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

type ReviewType = "mid" | "end";

// ExcelJS typing helpers
type Side = Partial<ExcelJS.Border>;
type Borders = Partial<ExcelJS.Borders>;

function buildPrefix(reviewType: ReviewType, fiscalYear?: string | null) {
    const periodCode = reviewType === "mid" ? "MOY" : "EOY";
    const fy = (fiscalYear ?? "").trim();
    return fy ? `${fy}_${periodCode} ` : `_${periodCode} `;
}

function fitSheetTitle(title: string) {
    const t = (title ?? "").trim();
    return t ? t.slice(0, 31) : "Sheet";
}

function safeFilename(name: string) {
    const cleaned = (name ?? "").replace(/[<>:"/\\|?*]/g, "_").trim();
    return cleaned || "Sheet";
}

function periodLabel(reviewType: ReviewType) {
    return reviewType === "mid" ? "Middle of Year" : "End of Year";
}

function getText(v: unknown) {
    if (v == null) return "";
    if (typeof v === "string") return v;
    if (typeof v === "number") return String(v);
    if (typeof v === "boolean") return v ? "TRUE" : "FALSE";
    try {
        if (typeof v === "object" && v && "text" in (v as any)) {
            return String((v as any).text ?? "");
        }
    } catch { }
    return String(v);
}

function isRowEmpty(ws: ExcelJS.Worksheet, rowNum: number, maxCols: number) {
    for (let c = 1; c <= maxCols; c++) {
        const v = ws.getCell(rowNum, c).value;
        if (v !== null && v !== undefined && getText(v).trim() !== "") return false;
    }
    return true;
}

function setBorder(cell: ExcelJS.Cell, border: Borders) {
    cell.border = border;
}
function setFill(cell: ExcelJS.Cell, fill: ExcelJS.Fill) {
    cell.fill = fill;
}
function setAlignment(cell: ExcelJS.Cell, alignment: Partial<ExcelJS.Alignment>) {
    cell.alignment = alignment;
}
function setFont(cell: ExcelJS.Cell, font: Partial<ExcelJS.Font>) {
    cell.font = font;
}

function appendKpiBlocks(args: {
    srcWs: ExcelJS.Worksheet;
    outWs: ExcelJS.Worksheet;
    startRow: number;

    boldFont: Partial<ExcelJS.Font>;
    scoreFont: Partial<ExcelJS.Font>;

    fillYellow: ExcelJS.Fill;
    fillLightGray: ExcelJS.Fill;
    fillDarkGray: ExcelJS.Fill;

    centerAlign: Partial<ExcelJS.Alignment>;
    centerTopAlign: Partial<ExcelJS.Alignment>;

    blackMedBorderStyle: ExcelJS.BorderStyle;
}) {
    const {
        srcWs,
        outWs,
        boldFont,
        scoreFont,
        fillYellow,
        fillLightGray,
        fillDarkGray,
        centerAlign,
        centerTopAlign,
        blackMedBorderStyle,
    } = args;

    let startRow = args.startRow;

    const labels = [
        "1 – Unsatisfactory",
        "2 – Needs Improvement",
        "3 – Proficient",
        "4 – Strong",
        "5 – Exemplary",
    ];

    const scoreValidation: ExcelJS.DataValidation = {
        type: "whole",
        operator: "between",
        allowBlank: true,
        formulae: [1, 5],
        showInputMessage: true,
        promptTitle: "Score Required",
        prompt: "Enter a whole number from 1 to 5.",
        showErrorMessage: true,
        errorTitle: "Invalid Entry",
        error: "Only whole numbers from 1 to 5 are allowed.",
    };

    const maxRow = srcWs.rowCount;

    for (let r = 2; r <= maxRow; r++) {
        // new Column A added, so consider A..H
        if (isRowEmpty(srcWs, r, 8)) continue;

        // shifted layout:
        // A = KPI Title
        // B = Competency
        // C = Description of Work
        // D..H = rating scale text
        const kpiTitle = getText(srcWs.getCell(r, 1).value);
        const competency = getText(srcWs.getCell(r, 2).value);
        const description = getText(srcWs.getCell(r, 3).value);
        const ratings = [
            getText(srcWs.getCell(r, 4).value),
            getText(srcWs.getCell(r, 5).value),
            getText(srcWs.getCell(r, 6).value),
            getText(srcWs.getCell(r, 7).value),
            getText(srcWs.getCell(r, 8).value),
        ];

        // ---- KPI block header row ----
        outWs.getCell(startRow, 2).value = `KPI - ${kpiTitle}`.trim();
        setFont(outWs.getCell(startRow, 2), boldFont);

        outWs.getCell(startRow, 3).value = "Description of Work";
        setFont(outWs.getCell(startRow, 3), boldFont);

        outWs.getCell(startRow, 4).value = "Score (1–5)";
        setFont(outWs.getCell(startRow, 4), boldFont);
        setAlignment(outWs.getCell(startRow, 4), centerTopAlign);

        // Notes / Goals headers
        outWs.getCell(startRow, 6).value = "Notes";
        setFont(outWs.getCell(startRow, 6), boldFont);
        setAlignment(outWs.getCell(startRow, 6), centerTopAlign);

        outWs.getCell(startRow, 7).value = "Goals";
        setFont(outWs.getCell(startRow, 7), boldFont);
        setAlignment(outWs.getCell(startRow, 7), centerTopAlign);

        // ---- KPI values row ----
        outWs.getCell(startRow + 1, 2).value = competency;
        outWs.getCell(startRow + 1, 3).value = description;

        // Score cell
        const scoreCell = outWs.getCell(startRow + 1, 4);
        scoreCell.value = "";
        setFill(scoreCell, fillYellow);
        setFont(scoreCell, scoreFont);
        setAlignment(scoreCell, centerAlign);
        scoreCell.dataValidation = scoreValidation;

        // Notes / Goals input cells
        const notesCell = outWs.getCell(startRow + 1, 6);
        notesCell.value = "";
        setFill(notesCell, fillYellow);
        setAlignment(notesCell, { wrapText: true, horizontal: "left", vertical: "top" });

        const goalsCell = outWs.getCell(startRow + 1, 7);
        goalsCell.value = "";
        setFill(goalsCell, fillYellow);
        setAlignment(goalsCell, { wrapText: true, horizontal: "left", vertical: "top" });

        // ---- Rating scale title row (above guidelines) ----
        outWs.mergeCells(startRow + 2, 2, startRow + 2, 4);
        const ratingTitleCell = outWs.getCell(startRow + 2, 2);
        ratingTitleCell.value = "Rating Scale (1–5)";
        setFont(ratingTitleCell, boldFont);
        setAlignment(ratingTitleCell, centerAlign);

        // ---- 5 rating rows ----
        for (let i = 0; i < 5; i++) {
            const rr = startRow + 3 + i;
            const fill = i % 2 === 0 ? fillLightGray : fillDarkGray;

            const b = outWs.getCell(rr, 2);
            b.value = labels[i];
            setFill(b, fill);
            setAlignment(b, { wrapText: true, horizontal: "left", vertical: "middle" });

            const c = outWs.getCell(rr, 3);
            c.value = ratings[i];
            setFill(c, fill);
            setAlignment(c, { wrapText: true, horizontal: "left", vertical: "middle" });

            const d = outWs.getCell(rr, 4);
            d.value = "";
            setFill(d, fill);
            setAlignment(d, { wrapText: true, horizontal: "left", vertical: "middle" });
        }

        // ---- Borders (clean + fixes)
        // Goals:
        // 1) Fix missing left border on merged "Rating Scale (1-5)" row
        // 2) Add a separator line between Question (rows startRow..startRow+1) and Rating Scale (startRow+2..)
        // 3) Make Notes/Goals wrap only around the Notes/Goals area (rows startRow..startRow+1), not a tall box
        const blackMed: Side = { style: blackMedBorderStyle, color: { argb: "FF000000" } };
        const whiteThin: Side = { style: "thin", color: { argb: "FFFFFFFF" } };

        const topRow = startRow;
        const midRow = startRow + 1; // last "question" row
        const sepRow = startRow + 2; // first "rating scale" row
        const bottomRow = startRow + 7;

        const mainLeftCol = 2; // B
        const mainRightCol = 4; // D (main block only)
        const notesLeftCol = 6; // F
        const notesRightCol = 7; // G

        // A) Main block outline only (B..D, rows startRow..startRow+7).
        for (let rr = topRow; rr <= bottomRow; rr++) {
            for (let cc = mainLeftCol; cc <= mainRightCol; cc++) {
                const isTop = rr === topRow;
                const isBottom = rr === bottomRow;
                const isLeft = cc === mainLeftCol;
                const isRight = cc === mainRightCol;

                setBorder(outWs.getCell(rr, cc), {
                    left: isLeft ? blackMed : whiteThin,
                    right: isRight ? blackMed : whiteThin,
                    top: isTop ? blackMed : whiteThin,
                    bottom: isBottom ? blackMed : whiteThin,
                });
            }
        }

        // B) Notes/Goals outline only for the first 2 rows (F..G, rows startRow..startRow+1)
        for (let rr = topRow; rr <= midRow; rr++) {
            for (let cc = notesLeftCol; cc <= notesRightCol; cc++) {
                const isTop = rr === topRow;
                const isBottom = rr === midRow;
                const isLeft = cc === notesLeftCol;
                const isRight = cc === notesRightCol;

                setBorder(outWs.getCell(rr, cc), {
                    left: isLeft ? blackMed : whiteThin,
                    right: isRight ? blackMed : whiteThin,
                    top: isTop ? blackMed : whiteThin,
                    bottom: isBottom ? blackMed : whiteThin,
                });
            }
        }

        // C) Make the cells below Notes/Goals (rows startRow+2..startRow+7) stay "blank/white bordered"
        for (let rr = sepRow; rr <= bottomRow; rr++) {
            for (let cc = notesLeftCol; cc <= notesRightCol; cc++) {
                setBorder(outWs.getCell(rr, cc), {
                    left: whiteThin,
                    right: whiteThin,
                    top: whiteThin,
                    bottom: whiteThin,
                });
            }
        }

        // D) Separator line between Question and Rating Scale across the main block (B..D)
        // Put a medium TOP border on row startRow+2 across B..D.
        for (let cc = mainLeftCol; cc <= mainRightCol; cc++) {
            const cell = outWs.getCell(sepRow, cc);
            const existing = cell.border ?? {};
            setBorder(cell, {
                ...existing,
                top: blackMed,
            });
        }

        // E) Fix merged "Rating Scale (1–5)" row left edge by forcing left border on ALL cells in the merged range B..D
        // (Excel sometimes only displays borders reliably when each underlying cell has the edge set.)
        for (let cc = 2; cc <= 4; cc++) {
            const cell = outWs.getCell(sepRow, cc);
            const existing = cell.border ?? {};
            setBorder(cell, {
                ...existing,
                left: blackMed, // ensures the merged row shows a left border
                // keep right edge correct on D as well
                right: cc === 4 ? blackMed : existing.right,
            });
        }

        // move to next KPI block (8 rows tall + 1 spacer row)
        startRow += 9;
    }

    return startRow;
}

function buildFormFromSheet(args: {
    srcWs: ExcelJS.Worksheet;
    outWs: ExcelJS.Worksheet;
    softWs?: ExcelJS.Worksheet | null;
    reviewType: ReviewType;
    fiscalYear?: string | null;
}) {
    const { srcWs, outWs, softWs, reviewType, fiscalYear } = args;

    // ---------- Styles ----------
    const whiteSide: Side = { style: "thin", color: { argb: "FFFFFFFF" } };
    const whiteBorder: Borders = {
        left: whiteSide,
        right: whiteSide,
        top: whiteSide,
        bottom: whiteSide,
    };

    const blackMedStyle: ExcelJS.BorderStyle = "medium";
    const blackMed: Side = { style: blackMedStyle, color: { argb: "FF000000" } };

    const boldFont: Partial<ExcelJS.Font> = { bold: true };
    const titleFont: Partial<ExcelJS.Font> = { bold: true, size: 20 };
    const redFont: Partial<ExcelJS.Font> = { color: { argb: "FFFF0000" } };
    const sectionFont: Partial<ExcelJS.Font> = { bold: true, size: 14 };

    const defaultAlign: Partial<ExcelJS.Alignment> = { wrapText: true, horizontal: "left", vertical: "top" };
    const centerAlign: Partial<ExcelJS.Alignment> = { wrapText: true, horizontal: "center", vertical: "middle" };
    const centerBottomAlign: Partial<ExcelJS.Alignment> = { wrapText: true, horizontal: "center", vertical: "bottom" };
    const centerTopAlign: Partial<ExcelJS.Alignment> = { wrapText: true, horizontal: "center", vertical: "top" };
    const leftCenterAlign: Partial<ExcelJS.Alignment> = { wrapText: true, horizontal: "left", vertical: "middle" };
    const leftVCenterAlign: Partial<ExcelJS.Alignment> = { wrapText: true, horizontal: "left", vertical: "middle" };
    const rightCenterAlign: Partial<ExcelJS.Alignment> = { horizontal: "right", vertical: "middle" };

    const fillLightGray: ExcelJS.Fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF5F5F5" } };
    const fillDarkGray: ExcelJS.Fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFEAEAEA" } };
    const fillYellow: ExcelJS.Fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFF00" } };

    const scoreFont: Partial<ExcelJS.Font> = { size: 20, color: { argb: "FFFF0000" } };

    // ---------- Column widths ----------
    outWs.getColumn(1).width = 3; // A
    outWs.getColumn(2).width = 50; // B
    outWs.getColumn(3).width = 70; // C
    outWs.getColumn(4).width = 16; // D
    outWs.getColumn(5).width = 3; // E spacer
    outWs.getColumn(6).width = 40; // F Notes
    outWs.getColumn(7).width = 40; // G Goals

    // ---------- Pre-draw white grid (blank/white borders everywhere by default) ----------
    for (let r = 1; r <= 600; r++) {
        for (let c = 1; c <= 26; c++) {
            const cell = outWs.getCell(r, c);
            setBorder(cell, whiteBorder);
            setAlignment(cell, defaultAlign);
        }
    }

    // ---------- Title ----------
    outWs.mergeCells("B2:D2");
    outWs.getCell("B2").value = `Performance Review Form for ${srcWs.name}`;
    setFont(outWs.getCell("B2"), titleFont);
    setAlignment(outWs.getCell("B2"), centerBottomAlign);
    outWs.getRow(2).height = 50; // bigger row 2

    outWs.mergeCells("B3:D3");
    outWs.getCell("B3").value = "Please rate each question on a scale of 1–5 using the yellow cell.";
    setAlignment(outWs.getCell("B3"), centerAlign);

    // ---------- Header block ----------
    const startRow = 5;
    const categoryStart = startRow + 1;

    const jobTitle = srcWs.name.endsWith(" KPI") ? srcWs.name.slice(0, -4).trim() : srcWs.name;

    // NEW: "Completed By" row underneath Period (dropdown: Employee/Coordinator)
    const fields: Array<[string, string, boolean, "none" | "role"]> = [
        ["Job Title:", jobTitle, false, "none"],
        ["Fiscal Year:", fiscalYear ?? "", false, "none"],
        ["Period:", periodLabel(reviewType), false, "none"],
        ["Completed By:", "", true, "role"], // NEW row
        ["First Name:", "", true, "none"],
        ["Last Name:", "", true, "none"],
    ];

    const categoryEnd = categoryStart + (fields.length - 1);

    const completedByValidation: ExcelJS.DataValidation = {
        type: "list",
        allowBlank: false,
        formulae: ['"Employee,Coordinator"'],
        showInputMessage: true,
        promptTitle: "Select Role",
        prompt: "Choose Employee or Coordinator.",
        showErrorMessage: true,
        errorTitle: "Invalid Selection",
        error: "Please select either Employee or Coordinator.",
    };

    for (let i = 0; i < fields.length; i++) {
        const [label, value, isYellow, kind] = fields[i];
        const r = categoryStart + i;

        const labelCell = outWs.getCell(r, 2);
        labelCell.value = label;
        setFont(labelCell, boldFont);
        setAlignment(labelCell, rightCenterAlign);

        // merged C:D for header block only
        outWs.mergeCells(r, 3, r, 4);
        const valCell = outWs.getCell(r, 3);
        valCell.value = value;
        setAlignment(valCell, leftVCenterAlign);

        // Yellow input background
        if (isYellow) {
            setFill(valCell, fillYellow);
            setFont(valCell, redFont);
        }

        // Role dropdown
        if (kind === "role") {
            valCell.value = ""; // keep blank until selected
            valCell.dataValidation = completedByValidation;
        }

        // keep borders blank/white inside header block (not "none")
        setBorder(outWs.getCell(r, 2), whiteBorder);
        setBorder(outWs.getCell(r, 3), whiteBorder);
        setBorder(outWs.getCell(r, 4), whiteBorder);
    }

    // outline header block with medium border
    for (let r = categoryStart; r <= categoryEnd; r++) {
        for (let c = 2; c <= 4; c++) {
            setBorder(outWs.getCell(r, c), {
                left: c === 2 ? blackMed : undefined,
                right: c === 4 ? blackMed : undefined,
                top: r === categoryStart ? blackMed : undefined,
                bottom: r === categoryEnd ? blackMed : undefined,
            });
        }
    }

    // ---------- Section A ----------
    let row = categoryEnd + 2;
    outWs.mergeCells(row, 2, row, 4);
    const secA = outWs.getCell(row, 2);
    secA.value = "Section A: Key Performance Indicators (KPIs)";
    setFont(secA, sectionFont);
    setAlignment(secA, leftCenterAlign);
    row += 2;

    row = appendKpiBlocks({
        srcWs,
        outWs,
        startRow: row,
        boldFont,
        scoreFont,
        fillYellow,
        fillLightGray,
        fillDarkGray,
        centerAlign,
        centerTopAlign,
        blackMedBorderStyle: blackMedStyle,
    });

    // ---------- Section B ----------
    row += 1;
    outWs.mergeCells(row, 2, row, 4);
    const secB = outWs.getCell(row, 2);
    secB.value = "Section B: Soft Skills";
    setFont(secB, sectionFont);
    setAlignment(secB, leftCenterAlign);
    row += 2;

    if (softWs) {
        appendKpiBlocks({
            srcWs: softWs,
            outWs,
            startRow: row,
            boldFont,
            scoreFont,
            fillYellow,
            fillLightGray,
            fillDarkGray,
            centerAlign,
            centerTopAlign,
            blackMedBorderStyle: blackMedStyle,
        });
    }
}

export async function POST(req: Request) {
    try {
        const form = await req.formData();

        const file = form.get("file");
        const reviewTypeRaw = getText(form.get("reviewType") ?? "mid").toLowerCase();
        const fiscalYear = getText(form.get("fiscalYear") ?? "").trim() || null;
        const reviewType: ReviewType = reviewTypeRaw === "end" ? "end" : "mid";

        if (!(file instanceof File)) {
            return NextResponse.json({ error: "Missing file" }, { status: 400 });
        }

        if (!file.name.toLowerCase().endsWith(".xlsx")) {
            return NextResponse.json({ error: "File must be .xlsx" }, { status: 400 });
        }

        // Avoid Node22 Buffer<T> mismatch by using Uint8Array
        const ab = await file.arrayBuffer();
        const inputBytes: Uint8Array = new Uint8Array(ab);

        const wb = new ExcelJS.Workbook();
        // @ts-expect-error ExcelJS type defs are too narrow; Uint8Array works at runtime
        await wb.xlsx.load(inputBytes);

        const softWs = wb.getWorksheet("Soft Skills KPI") ?? null;

        const zip = new JSZip();
        let created = 0;

        const prefix = buildPrefix(reviewType, fiscalYear);

        for (const srcWs of wb.worksheets) {
            if (srcWs.state && srcWs.state !== "visible") continue;
            if (srcWs.name === "Dashboard" || srcWs.name === "Soft Skills KPI") continue;

            const outWb = new ExcelJS.Workbook();
            const outWs = outWb.addWorksheet(fitSheetTitle(`${prefix}${srcWs.name}`));

            buildFormFromSheet({
                srcWs,
                outWs,
                softWs,
                reviewType,
                fiscalYear,
            });

            const outFilename = `${safeFilename(prefix + srcWs.name)}.xlsx`;

            const outArrayBuffer = (await outWb.xlsx.writeBuffer()) as ArrayBuffer;
            zip.file(outFilename, outArrayBuffer);

            created++;
        }

        if (created === 0) {
            return NextResponse.json(
                { error: "No visible KPI sheets found to generate forms from." },
                { status: 400 }
            );
        }

        const zipBytes = await zip.generateAsync({ type: "uint8array" });

        // @ts-expect-error ExcelJS type defs are too narrow; Uint8Array works at runtime
        return new NextResponse(zipBytes, {
            status: 200,
            headers: {
                "Content-Type": "application/zip",
                "Content-Disposition": `attachment; filename="generated_forms.zip"`,
            },
        });
    } catch (err: any) {
        return NextResponse.json({ error: err?.message ?? "Server error" }, { status: 500 });
    }
}

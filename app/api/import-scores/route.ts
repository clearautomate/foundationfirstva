import ExcelJS from "exceljs";
import { NextResponse } from "next/server";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

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

function getCellText(ws: ExcelJS.Worksheet, address: string) {
    return getText(ws.getCell(address).value).trim();
}

function normalizeCompletedBy(raw: string): "Employee" | "Coordinator" | "Unknown" {
    const t = (raw ?? "").trim().toLowerCase();
    if (t === "employee") return "Employee";
    if (t === "coordinator") return "Coordinator";
    return "Unknown";
}

function isYellowFill(cell: ExcelJS.Cell) {
    const fill = cell.fill as ExcelJS.Fill | undefined;
    if (!fill || (fill as any).type !== "pattern") return false;

    const fg = (fill as any).fgColor?.argb as string | undefined;
    if (!fg) return false;

    const hex = String(fg).toUpperCase();
    return hex.endsWith("FFFF00");
}

function readNumericScore(cell: ExcelJS.Cell) {
    const v = cell.value;
    if (typeof v === "number") return v;
    if (typeof v === "string") {
        const n = Number(v.trim());
        return Number.isFinite(n) ? n : null;
    }
    return null;
}

function clamp(n: number, min: number, max: number) {
    return Math.max(min, Math.min(max, n));
}

function normName(s: string) {
    return (s ?? "").trim().toLowerCase().replace(/\s+/g, " ");
}

type ImportedRow = {
    fileName: string;

    jobTitle: string;
    fiscalYear: string;
    period: string;
    firstName: string;
    lastName: string;

    completedBy: "Employee" | "Coordinator" | "Unknown";

    avgScore: number | null;
    percent: number | null;

    employeeScorePercent: number | null;
    coordinatorScorePercent: number | null;
};

type CombinedRow = {
    // identity
    jobTitle: string;
    fiscalYear: string;
    period: string;
    firstName: string;
    lastName: string;

    // sources
    employeeFileName: string | null;
    coordinatorFileName: string | null;

    // scores
    employeeScorePercent: number | null;
    coordinatorScorePercent: number | null;
};

function mergeRowsByName(rows: ImportedRow[]): CombinedRow[] {
    // Keyed ONLY by first+last (as you requested).
    // If you want safer grouping, include jobTitle/fiscalYear/period in the key.
    const map = new Map<string, CombinedRow>();

    for (const r of rows) {
        const key = `${normName(r.firstName)}|${normName(r.lastName)}`;

        const existing =
            map.get(key) ??
            ({
                jobTitle: r.jobTitle,
                fiscalYear: r.fiscalYear,
                period: r.period,
                firstName: r.firstName,
                lastName: r.lastName,

                employeeFileName: null,
                coordinatorFileName: null,

                employeeScorePercent: null,
                coordinatorScorePercent: null,
            } satisfies CombinedRow);

        // Keep the first non-empty meta values (optional)
        if (!existing.jobTitle && r.jobTitle) existing.jobTitle = r.jobTitle;
        if (!existing.fiscalYear && r.fiscalYear) existing.fiscalYear = r.fiscalYear;
        if (!existing.period && r.period) existing.period = r.period;
        if (!existing.firstName && r.firstName) existing.firstName = r.firstName;
        if (!existing.lastName && r.lastName) existing.lastName = r.lastName;

        // Merge score into the right slot
        if (r.completedBy === "Employee") {
            // If duplicate employee file exists, last one wins (change if you want)
            existing.employeeScorePercent = r.percent;
            existing.employeeFileName = r.fileName;
        } else if (r.completedBy === "Coordinator") {
            existing.coordinatorScorePercent = r.percent;
            existing.coordinatorFileName = r.fileName;
        } else {
            // Unknown: do nothing, or pick a side if you want
        }

        map.set(key, existing);
    }

    return Array.from(map.values());
}

export async function POST(req: Request) {
    try {
        const form = await req.formData();
        const files = form.getAll("files");
        const excelFiles = files.filter((f): f is File => f instanceof File);

        if (!excelFiles.length) {
            return NextResponse.json({ error: "Missing files" }, { status: 400 });
        }

        const bad = excelFiles.find((f) => !f.name.toLowerCase().endsWith(".xlsx"));
        if (bad) {
            return NextResponse.json({ error: "All files must be .xlsx" }, { status: 400 });
        }

        const rows: ImportedRow[] = [];

        for (const file of excelFiles) {
            const ab = await file.arrayBuffer();
            const inputBytes = new Uint8Array(ab);

            const wb = new ExcelJS.Workbook();
            // @ts-expect-error ExcelJS type defs are too narrow; Uint8Array works at runtime
            await wb.xlsx.load(inputBytes);

            const ws = wb.worksheets?.[0];
            if (!ws) continue;

            const jobTitle = getCellText(ws, "C6");
            const fiscalYear = getCellText(ws, "C7");
            const period = getCellText(ws, "C8");
            const completedBy = normalizeCompletedBy(getCellText(ws, "C9"));
            const firstName = getCellText(ws, "C10");
            const lastName = getCellText(ws, "C11");

            const scores: number[] = [];
            ws.eachRow((row) => {
                row.eachCell((cell) => {
                    if (!isYellowFill(cell)) return;
                    const n = readNumericScore(cell);
                    if (n == null) return;
                    if (n >= 1 && n <= 5) scores.push(n);
                });
            });

            let avgScore: number | null = null;
            let percent: number | null = null;

            if (scores.length) {
                const sum = scores.reduce((s, n) => s + n, 0);
                avgScore = sum / scores.length;
                percent = clamp((avgScore / 5) * 100, 0, 100);
            }

            const employeeScorePercent = completedBy === "Employee" ? percent : null;
            const coordinatorScorePercent = completedBy === "Coordinator" ? percent : null;

            rows.push({
                fileName: file.name,
                jobTitle,
                fiscalYear,
                period,
                firstName,
                lastName,
                completedBy,
                avgScore,
                percent,
                employeeScorePercent,
                coordinatorScorePercent,
            });
        }

        // ✅ COMBINE INTO ONE ROW PER NAME
        const combinedRows = mergeRowsByName(rows);

        return NextResponse.json({ rows: combinedRows }, { status: 200 });
    } catch (err: any) {
        return NextResponse.json({ error: err?.message ?? "Server error" }, { status: 500 });
    }
}

"use client";

import { useMemo, useRef, useState } from "react";
import styles from "./importScores.module.css";

type ImportedRow = {
    fileName: string;

    jobTitle: string;
    fiscalYear: string;
    period: string;
    firstName: string;
    lastName: string;

    completedBy: "Employee" | "Coordinator" | "Unknown";

    avgScore: number | null;     // 1-5
    percent: number | null;      // 0-100

    employeeScorePercent: number | null;
    coordinatorScorePercent: number | null;
};

function formatPercent(v: number | null) {
    if (v == null || Number.isNaN(v)) return "";
    return `${v.toFixed(1)}%`;
}

export default function ImportScores() {
    const inputRef = useRef<HTMLInputElement | null>(null);

    const [files, setFiles] = useState<File[]>([]);
    const [rows, setRows] = useState<ImportedRow[]>([]);

    const [busy, setBusy] = useState(false);
    const [error, setError] = useState<string>("");
    const [copied, setCopied] = useState(false);

    const fileLabel = useMemo(() => {
        if (!files.length) return "No files selected";
        const totalKb = Math.round(files.reduce((s, f) => s + f.size, 0) / 1024);
        return `${files.length} file(s) selected (${totalKb} KB total)`;
    }, [files]);

    function clearFilesOnly() {
        setFiles([]);
        if (inputRef.current) inputRef.current.value = "";
    }

    function clearAll() {
        clearFilesOnly();
        setRows([]);
        setError("");
        setCopied(false);
    }

    function onPickFiles(e: React.ChangeEvent<HTMLInputElement>) {
        setError("");
        setCopied(false);

        const picked = Array.from(e.target.files ?? []);
        if (!picked.length) return;

        // allow add-more behavior; validate only .xlsx
        const bad = picked.find((f) => !f.name.toLowerCase().endsWith(".xlsx"));
        if (bad) {
            setError("Please upload only .xlsx files (Excel workbooks).");
            // don't add anything if invalid
            if (inputRef.current) inputRef.current.value = "";
            return;
        }

        setFiles((prev) => [...prev, ...picked]);
    }

    async function onImport() {
        setError("");
        setCopied(false);

        if (!files.length) {
            setError("Please select one or more .xlsx files first.");
            return;
        }

        setBusy(true);
        try {
            const form = new FormData();
            for (const f of files) form.append("files", f);

            const res = await fetch("/api/import-scores", {
                method: "POST",
                body: form,
            });

            if (!res.ok) {
                const data = await res.json().catch(() => ({}));
                throw new Error(data?.error || "Failed to import scores");
            }

            const data = (await res.json()) as { rows: ImportedRow[] };

            // add rows (do not replace)
            setRows((prev) => [...prev, ...(data.rows ?? [])]);

            // optional: clear file picker after import
            clearFilesOnly();
        } catch (err: any) {
            setError(err?.message || "Something went wrong.");
        } finally {
            setBusy(false);
        }
    }

    async function onCopy() {
        setError("");
        setCopied(false);

        if (!rows.length) {
            setError("Nothing to copy yet. Import some files first.");
            return;
        }

        // Copy ONLY data rows (no header) as TSV for clean paste into Excel.
        const lines = rows.map((r) => {
            const cols = [
                r.jobTitle ?? "",
                r.fiscalYear ?? "",
                r.period ?? "",
                r.firstName ?? "",
                r.lastName ?? "",
                formatPercent(r.employeeScorePercent),
                formatPercent(r.coordinatorScorePercent),
            ];
            return cols.map((c) => String(c).replace(/\r?\n/g, " ")).join("\t");
        });

        const tsv = lines.join("\n");
        await navigator.clipboard.writeText(tsv);

        setCopied(true);
        setTimeout(() => setCopied(false), 1200);
    }

    return (
        <div className={styles.container}>
            <div className={styles.card}>
                <label className={styles.fileGroup}>
                    <span className={styles.label}>Excel files (.xlsx only)</span>

                    <input
                        ref={inputRef}
                        type="file"
                        multiple
                        accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        onChange={onPickFiles}
                        className={styles.fileInput}
                        disabled={busy}
                    />

                    <small className={styles.fileName}>{fileLabel}</small>
                </label>

                {error ? <div className={styles.error}>{error}</div> : null}

                <div className={styles.actions}>
                    <button
                        onClick={clearAll}
                        disabled={busy && !files.length && !rows.length}
                        className={styles.secondaryButton}
                    >
                        Clear
                    </button>

                    <button
                        onClick={onImport}
                        disabled={!files.length || busy}
                        className={styles.button}
                    >
                        {busy ? "Importing..." : "Import Scores"}
                    </button>

                    <button
                        onClick={onCopy}
                        disabled={!rows.length || busy}
                        className={styles.button}
                        title="Copies rows only (no header) for easy paste into Excel"
                    >
                        {copied ? "Copied!" : "Copy"}
                    </button>
                </div>
            </div>

            {rows.length ? (
                <div className={styles.tableCard}>
                    <div className={styles.tableTopRow}>
                        <div className={styles.tableMeta}>
                            <span className={styles.metaPill}>{rows.length} row(s)</span>
                            <span className={styles.metaHint}>
                                Copy button excludes the header.
                            </span>
                        </div>
                    </div>

                    <div className={styles.tableWrap}>
                        <table className={styles.table}>
                            <thead>
                                <tr>
                                    <th>Job Title</th>
                                    <th>Fiscal Year</th>
                                    <th>Period</th>
                                    <th>First Name</th>
                                    <th>Last Name</th>
                                    <th>Employee Score</th>
                                    <th>Coordinator Score</th>
                                </tr>
                            </thead>
                            <tbody>
                                {rows.map((r, idx) => (
                                    <tr key={`${r.fileName}-${idx}`}>
                                        <td className={styles.monoHint} title={r.fileName}>
                                            {r.jobTitle}
                                        </td>
                                        <td>{r.fiscalYear}</td>
                                        <td>{r.period}</td>
                                        <td>{r.firstName}</td>
                                        <td>{r.lastName}</td>
                                        <td>{formatPercent(r.employeeScorePercent)}</td>
                                        <td>{formatPercent(r.coordinatorScorePercent)}</td>
                                    </tr>
                                ))}
                            </tbody>
                        </table>
                    </div>
                </div>
            ) : null}
        </div>
    );
}

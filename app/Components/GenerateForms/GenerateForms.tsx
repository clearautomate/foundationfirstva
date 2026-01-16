"use client";

import { useMemo, useRef, useState } from "react";
import styles from "./generateForms.module.css";

type ReviewType = "mid" | "end";

export default function GenerateForms() {
    const inputRef = useRef<HTMLInputElement | null>(null);

    const [file, setFile] = useState<File | null>(null);
    const [reviewType, setReviewType] = useState<ReviewType>("mid");
    const [fiscalYear, setFiscalYear] = useState<string>("FY25");

    const [busy, setBusy] = useState(false);
    const [error, setError] = useState<string>("");

    const fileLabel = useMemo(() => {
        if (!file) return "No file selected";
        return `${file.name} (${Math.round(file.size / 1024)} KB)`;
    }, [file]);

    function clearFile() {
        setFile(null);
        if (inputRef.current) inputRef.current.value = "";
    }

    function onPickFile(e: React.ChangeEvent<HTMLInputElement>) {
        setError("");
        const picked = e.target.files?.[0] ?? null;

        if (!picked) {
            setFile(null);
            return;
        }

        const lower = picked.name.toLowerCase();

        // Enforce .xlsx ONLY
        if (!lower.endsWith(".xlsx")) {
            setFile(null);
            if (inputRef.current) inputRef.current.value = "";
            setError("Please upload a .xlsx file (Excel workbook).");
            return;
        }

        setFile(picked);
    }

    async function onGenerate() {
        setError("");
        if (!file) {
            setError("Please select a .xlsx file first.");
            return;
        }

        const lower = file.name.toLowerCase();
        if (!lower.endsWith(".xlsx")) {
            setError("Please upload a .xlsx file (Excel workbook).");
            return;
        }

        setBusy(true);
        try {
            const form = new FormData();
            form.append("file", file);
            form.append("reviewType", reviewType);
            form.append("fiscalYear", fiscalYear.trim());

            const res = await fetch("/api/generate-forms", {
                method: "POST",
                body: form,
            });

            if (!res.ok) {
                // API sends JSON on errors
                const data = await res.json().catch(() => ({}));
                throw new Error(data?.error || "Failed to generate forms");
            }

            // ZIP download
            const blob = await res.blob();
            const url = URL.createObjectURL(blob);

            const a = document.createElement("a");
            a.href = url;
            a.download = "generated_forms.zip";
            document.body.appendChild(a);
            a.click();
            a.remove();

            URL.revokeObjectURL(url);
        } catch (err: any) {
            setError(err?.message || "Something went wrong.");
        } finally {
            setBusy(false);
        }
    }

    return (
        <div className={styles.container}>
            <div className={styles.card}>
                <label className={styles.fileGroup}>
                    <span className={styles.label}>Excel file (.xlsx only)</span>

                    <input
                        ref={inputRef}
                        type="file"
                        accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        onChange={onPickFile}
                        className={styles.fileInput}
                    />

                    <small className={styles.fileName}>{fileLabel}</small>
                </label>

                <div className={styles.row}>
                    <label className={styles.field}>
                        <span className={styles.label}>Review Type</span>
                        <select
                            className={styles.select}
                            value={reviewType}
                            onChange={(e) => setReviewType(e.target.value as ReviewType)}
                            disabled={busy}
                        >
                            <option value="mid">Mid (MOY)</option>
                            <option value="end">End (EOY)</option>
                        </select>
                    </label>

                    <label className={styles.field}>
                        <span className={styles.label}>Fiscal Year</span>
                        <input
                            className={styles.input}
                            value={fiscalYear}
                            onChange={(e) => setFiscalYear(e.target.value)}
                            placeholder="FY25"
                            disabled={busy}
                        />
                    </label>
                </div>

                {error ? <div className={styles.error}>{error}</div> : null}

                <div className={styles.actions}>
                    <button
                        onClick={clearFile}
                        disabled={!file || busy}
                        className={styles.secondaryButton}
                    >
                        Clear
                    </button>

                    <button
                        onClick={onGenerate}
                        disabled={!file || busy}
                        className={styles.button}
                    >
                        {busy ? "Generating..." : "Generate Forms"}
                    </button>
                </div>
            </div>
        </div>
    );
}

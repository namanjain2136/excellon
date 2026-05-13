import { createFileRoute } from "@tanstack/react-router";
import { useCallback, useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";
import FileSaver from "file-saver";
const { saveAs } = FileSaver;
import { toast } from "sonner";
import {
  FileSpreadsheet,
  Upload,
  Download,
  CheckCircle2,
  Loader2,
  Sparkles,
  Settings2,
  ListChecks,
  X,
} from "lucide-react";

import { Button } from "@/components/ui/button";
import { Card } from "@/components/ui/card";
import { Checkbox } from "@/components/ui/checkbox";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { RadioGroup, RadioGroupItem } from "@/components/ui/radio-group";
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "@/components/ui/select";
import { Progress } from "@/components/ui/progress";
import { Badge } from "@/components/ui/badge";

import { PrivacyBadge } from "@/components/excellon/PrivacyBadge";
import { detectNumericColumns, processSheet, type ColumnInfo } from "@/lib/workbook";
import type { ConvertOptions, NumberSystem, CaseStyle } from "@/lib/numberToWords";

export const Route = createFileRoute("/")({
  component: ExcellonPage,
});

type Status =
  | { kind: "idle" }
  | { kind: "uploading" }
  | { kind: "reading" }
  | { kind: "detecting" }
  | { kind: "ready" }
  | { kind: "processing"; progress: number }
  | { kind: "done" };

function ExcellonPage() {
  const [file, setFile] = useState<File | null>(null);
  const [workbook, setWorkbook] = useState<XLSX.WorkBook | null>(null);
  const [sheetName, setSheetName] = useState<string>("");
  const [columns, setColumns] = useState<ColumnInfo[]>([]);
  const [headerRow, setHeaderRow] = useState<number>(0);
  const [selected, setSelected] = useState<Set<number>>(new Set());
  const [status, setStatus] = useState<Status>({ kind: "idle" });
  const [dragOver, setDragOver] = useState(false);

  const [system, setSystem] = useState<NumberSystem>("indian");
  const [caseStyle, setCaseStyle] = useState<CaseStyle>("title");
  const [suffixChoice, setSuffixChoice] = useState<"rupees only" | "only" | "custom">("rupees only");
  const [customSuffix, setCustomSuffix] = useState("");
  const [paise, setPaise] = useState(true);

  const fileInputRef = useRef<HTMLInputElement>(null);

  const opts: ConvertOptions = useMemo(
    () => ({
      system,
      caseStyle,
      suffix: suffixChoice === "custom" ? customSuffix.trim() : suffixChoice,
      paise,
    }),
    [system, caseStyle, suffixChoice, customSuffix, paise],
  );

  const reset = () => {
    setFile(null);
    setWorkbook(null);
    setSheetName("");
    setColumns([]);
    setSelected(new Set());
    setStatus({ kind: "idle" });
  };

  const loadFile = useCallback(async (f: File) => {
    if (!f.name.toLowerCase().endsWith(".xlsx")) {
      toast.error("Please upload a valid .xlsx Excel file.");
      return;
    }
    if (f.size > 25 * 1024 * 1024) {
      toast.warning("This file is large and may take longer to process.");
    }
    setFile(f);
    setStatus({ kind: "uploading" });
    try {
      const buf = await f.arrayBuffer();
      setStatus({ kind: "reading" });
      const wb = XLSX.read(buf, { type: "array", cellStyles: true, cellFormula: true, cellNF: true });
      if (!wb.SheetNames.length) {
        toast.error("This workbook format is not fully supported.");
        reset();
        return;
      }
      setWorkbook(wb);
      const first = wb.SheetNames[0];
      setSheetName(first);
      setStatus({ kind: "detecting" });
      const { headerRow: hr, columns: cols } = detectNumericColumns(wb.Sheets[first]);
      setHeaderRow(hr);
      setColumns(cols);
      setSelected(new Set(cols.filter((c) => c.isCurrencyLike).map((c) => c.index)));
      if (cols.length === 0) {
        toast.warning("No numeric columns were detected in this sheet.");
      } else {
        toast.success(`Detected header at row ${hr + 1} · ${cols.length} numeric column${cols.length > 1 ? "s" : ""}.`);
      }
      setStatus({ kind: "ready" });
    } catch (e) {
      console.error(e);
      toast.error("Something went wrong while processing the workbook.");
      reset();
    }
  }, []);

  const onSheetChange = (name: string) => {
    if (!workbook) return;
    setSheetName(name);
    const { headerRow: hr, columns: cols } = detectNumericColumns(workbook.Sheets[name]);
    setHeaderRow(hr);
    setColumns(cols);
    setSelected(new Set(cols.filter((c) => c.isCurrencyLike).map((c) => c.index)));
    if (cols.length === 0) toast.warning("No numeric columns were detected in this sheet.");
  };

  const toggleCol = (idx: number) => {
    setSelected((prev) => {
      const next = new Set(prev);
      if (next.has(idx)) next.delete(idx);
      else next.add(idx);
      return next;
    });
  };

  const onDrop = (e: React.DragEvent) => {
    e.preventDefault();
    setDragOver(false);
    const f = e.dataTransfer.files?.[0];
    if (f) loadFile(f);
  };

  const handleProcess = async () => {
    if (!workbook || !file) return;
    if (selected.size === 0) {
      toast.error("Please select at least one column to convert.");
      return;
    }
    setStatus({ kind: "processing", progress: 10 });
    try {
      // Process all sheets — but only convert selected columns on the active sheet.
      // Per spec: insert new column for each selected column on current sheet.
      const ws = workbook.Sheets[sheetName];
      setStatus({ kind: "processing", progress: 40 });
      const { skipped } = processSheet(ws, Array.from(selected), opts);
      setStatus({ kind: "processing", progress: 75 });

      const out = XLSX.write(workbook, {
        bookType: "xlsx",
        type: "array",
        cellStyles: true,
      });
      setStatus({ kind: "processing", progress: 95 });

      const baseName = file.name.replace(/\.xlsx$/i, "");
      const blob = new Blob([out], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });
      saveAs(blob, `${baseName}_processed.xlsx`);
      setStatus({ kind: "done" });
      toast.success("File ready — download started.");
      if (skipped > 0) {
        toast.message(`${skipped} cell${skipped > 1 ? "s were" : " was"} skipped (empty or invalid).`);
      }
    } catch (e) {
      console.error(e);
      toast.error("Unable to generate the processed file.");
      setStatus({ kind: "ready" });
    }
  };

  const statusLabel = (() => {
    switch (status.kind) {
      case "uploading": return "Uploading...";
      case "reading": return "Reading workbook...";
      case "detecting": return "Detecting numeric columns...";
      case "processing": return "Processing file...";
      case "done": return "File ready for download";
      default: return null;
    }
  })();

  return (
    <div className="min-h-screen bg-gradient-soft text-foreground">
      {/* Decorative blobs */}
      <div className="pointer-events-none absolute inset-x-0 top-0 -z-10 h-[500px] overflow-hidden">
        <div className="absolute -top-40 left-1/4 h-96 w-96 rounded-full bg-primary/20 blur-3xl" />
        <div className="absolute -top-20 right-1/4 h-80 w-80 rounded-full bg-primary-glow/20 blur-3xl" />
      </div>

      {/* Header */}
      <header className="container mx-auto px-6 pt-10 pb-6">
        <div className="flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="flex h-10 w-10 items-center justify-center rounded-xl bg-gradient-primary shadow-soft">
              <FileSpreadsheet className="h-5 w-5 text-primary-foreground" />
            </div>
            <span className="text-xl font-semibold tracking-tight">Excellon</span>
          </div>
          <PrivacyBadge />
        </div>
      </header>

      {/* Hero */}
      <section className="container mx-auto px-6 pb-10 text-center">
        <div className="mx-auto max-w-3xl animate-fade-in">
          <Badge variant="secondary" className="mb-4 rounded-full px-3 py-1">
            <Sparkles className="mr-1 h-3 w-3" /> 100% private · in-browser conversion
          </Badge>
          <h1 className="bg-gradient-primary bg-clip-text text-5xl font-bold tracking-tight text-transparent md:text-6xl">
            Excellon
          </h1>
          <p className="mt-4 text-lg text-muted-foreground">
            Convert Excel numeric amounts into Indian English words instantly while preserving your workbook formatting.
          </p>
        </div>
      </section>

      {/* Main grid */}
      <main className="container mx-auto grid grid-cols-1 gap-6 px-6 pb-16 lg:grid-cols-12">
        {/* Left: Options */}
        <aside className="lg:col-span-3">
          <Card className="rounded-2xl border-primary/10 bg-card/70 p-5 shadow-soft backdrop-blur-md">
            <div className="mb-4 flex items-center gap-2">
              <Settings2 className="h-4 w-4 text-primary" />
              <h2 className="text-sm font-semibold">Options</h2>
            </div>

            <div className="space-y-5">
              <div>
                <Label className="text-xs uppercase tracking-wide text-muted-foreground">Numbering</Label>
                <RadioGroup value={system} onValueChange={(v) => setSystem(v as NumberSystem)} className="mt-2 space-y-1.5">
                  <label className="flex items-center gap-2 text-sm">
                    <RadioGroupItem value="indian" /> Indian (Lakh / Crore)
                  </label>
                  <label className="flex items-center gap-2 text-sm">
                    <RadioGroupItem value="international" /> International (Million / Billion)
                  </label>
                </RadioGroup>
              </div>

              <div>
                <Label className="text-xs uppercase tracking-wide text-muted-foreground">Output Case</Label>
                <RadioGroup value={caseStyle} onValueChange={(v) => setCaseStyle(v as CaseStyle)} className="mt-2 space-y-1.5">
                  <label className="flex items-center gap-2 text-sm">
                    <RadioGroupItem value="title" /> Title Case
                  </label>
                  <label className="flex items-center gap-2 text-sm">
                    <RadioGroupItem value="upper" /> UPPERCASE
                  </label>
                  <label className="flex items-center gap-2 text-sm">
                    <RadioGroupItem value="lower" /> lowercase
                  </label>
                </RadioGroup>
              </div>

              <div>
                <Label className="text-xs uppercase tracking-wide text-muted-foreground">Suffix</Label>
                <RadioGroup value={suffixChoice} onValueChange={(v) => setSuffixChoice(v as typeof suffixChoice)} className="mt-2 space-y-1.5">
                  <label className="flex items-center gap-2 text-sm">
                    <RadioGroupItem value="rupees only" /> rupees only
                  </label>
                  <label className="flex items-center gap-2 text-sm">
                    <RadioGroupItem value="only" /> only
                  </label>
                  <label className="flex items-center gap-2 text-sm">
                    <RadioGroupItem value="custom" /> Custom
                  </label>
                </RadioGroup>
                {suffixChoice === "custom" && (
                  <Input
                    className="mt-2 h-9"
                    placeholder="e.g. INR only"
                    value={customSuffix}
                    onChange={(e) => setCustomSuffix(e.target.value)}
                  />
                )}
              </div>

              <label className="flex cursor-pointer items-start gap-2 rounded-lg bg-secondary/60 p-3 text-sm">
                <Checkbox checked={paise} onCheckedChange={(v) => setPaise(!!v)} className="mt-0.5" />
                <span>
                  <span className="font-medium">Convert decimals into paise</span>
                  <span className="mt-0.5 block text-xs text-muted-foreground">
                    e.g. 85432.75 → …rupees and seventy five paise only
                  </span>
                </span>
              </label>
            </div>
          </Card>
        </aside>

        {/* Center: Upload + Preview */}
        <section className="lg:col-span-6 space-y-6">
          <Card
            onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
            onDragLeave={() => setDragOver(false)}
            onDrop={onDrop}
            className={`relative overflow-hidden rounded-2xl border-2 border-dashed p-10 text-center transition-all ${
              dragOver ? "border-primary bg-primary/5 scale-[1.01]" : "border-primary/20 bg-card/70"
            } shadow-soft backdrop-blur-md`}
          >
            {!file ? (
              <div className="space-y-4 animate-fade-in">
                <div className="mx-auto flex h-16 w-16 items-center justify-center rounded-2xl bg-gradient-primary shadow-glow">
                  <Upload className="h-7 w-7 text-primary-foreground" />
                </div>
                <div>
                  <h3 className="text-lg font-semibold">Drop your .xlsx file here</h3>
                  <p className="mt-1 text-sm text-muted-foreground">or click to browse from your device</p>
                </div>
                <Button onClick={() => fileInputRef.current?.click()} size="lg" className="rounded-full px-6">
                  Browse File
                </Button>
                <input
                  ref={fileInputRef}
                  type="file"
                  accept=".xlsx"
                  className="hidden"
                  onChange={(e) => e.target.files?.[0] && loadFile(e.target.files[0])}
                />
              </div>
            ) : (
              <div className="space-y-3 animate-scale-in">
                <div className="mx-auto flex h-14 w-14 items-center justify-center rounded-2xl bg-primary/10">
                  <CheckCircle2 className="h-7 w-7 text-primary" />
                </div>
                <div>
                  <p className="text-sm text-muted-foreground">Loaded file</p>
                  <p className="text-base font-semibold">{file.name}</p>
                  <p className="text-xs text-muted-foreground">{(file.size / 1024).toFixed(1)} KB</p>
                </div>
                <Button variant="ghost" size="sm" onClick={reset} className="rounded-full">
                  <X className="mr-1 h-3 w-3" /> Remove
                </Button>
              </div>
            )}
          </Card>

          {workbook && workbook.SheetNames.length > 1 && (
            <Card className="rounded-2xl border-primary/10 bg-card/70 p-5 shadow-soft backdrop-blur-md animate-fade-in">
              <Label className="text-xs uppercase tracking-wide text-muted-foreground">Sheet</Label>
              <Select value={sheetName} onValueChange={onSheetChange}>
                <SelectTrigger className="mt-2 rounded-xl">
                  <SelectValue />
                </SelectTrigger>
                <SelectContent>
                  {workbook.SheetNames.map((n) => (
                    <SelectItem key={n} value={n}>{n}</SelectItem>
                  ))}
                </SelectContent>
              </Select>
            </Card>
          )}

          {statusLabel && (
            <Card className="rounded-2xl border-primary/10 bg-card/70 p-5 shadow-soft backdrop-blur-md">
              <div className="mb-2 flex items-center gap-2 text-sm font-medium">
                {status.kind === "done" ? (
                  <CheckCircle2 className="h-4 w-4 text-primary" />
                ) : (
                  <Loader2 className="h-4 w-4 animate-spin text-primary" />
                )}
                {statusLabel}
              </div>
              <Progress value={status.kind === "processing" ? status.progress : status.kind === "done" ? 100 : 60} className="h-2" />
            </Card>
          )}

          {file && (
            <div className="flex flex-wrap items-center justify-end gap-3">
              <Button variant="outline" onClick={reset} className="rounded-full">Start over</Button>
              <Button
                onClick={handleProcess}
                disabled={!workbook || selected.size === 0 || status.kind === "processing"}
                size="lg"
                className="rounded-full bg-gradient-primary px-6 shadow-soft hover:shadow-glow transition-shadow"
              >
                {status.kind === "processing" ? (
                  <><Loader2 className="mr-2 h-4 w-4 animate-spin" /> Processing</>
                ) : (
                  <><Download className="mr-2 h-4 w-4" /> Convert &amp; Download</>
                )}
              </Button>
            </div>
          )}
        </section>

        {/* Right: Detected columns */}
        <aside className="lg:col-span-3">
          <Card className="rounded-2xl border-primary/10 bg-card/70 p-5 shadow-soft backdrop-blur-md">
            <div className="mb-4 flex items-center gap-2">
              <ListChecks className="h-4 w-4 text-primary" />
              <h2 className="text-sm font-semibold">Detected Columns</h2>
            </div>

            {!workbook ? (
              <p className="text-sm text-muted-foreground">Upload a file to see numeric columns here.</p>
            ) : columns.length === 0 ? (
              <p className="text-sm text-muted-foreground">No numeric columns detected in this sheet.</p>
            ) : (
              <ul className="space-y-2 max-h-[460px] overflow-auto pr-1">
                {columns.map((c) => {
                  const checked = selected.has(c.index);
                  return (
                    <li key={c.index}>
                      <label
                        className={`flex cursor-pointer items-start gap-2 rounded-xl border p-3 text-sm transition-all hover:border-primary/40 ${
                          checked ? "border-primary/50 bg-primary/5" : "border-border bg-background"
                        }`}
                      >
                        <Checkbox checked={checked} onCheckedChange={() => toggleCol(c.index)} className="mt-0.5" />
                        <div className="min-w-0 flex-1">
                          <div className="truncate font-medium">{c.header}</div>
                          <div className="mt-1 flex items-center gap-1.5">
                            <Badge variant="secondary" className="rounded-full px-2 py-0 text-[10px]">
                              {Math.round(c.numericRatio * 100)}% numeric
                            </Badge>
                            {c.isCurrencyLike && (
                              <Badge className="rounded-full bg-primary/15 px-2 py-0 text-[10px] text-primary hover:bg-primary/15">
                                currency
                              </Badge>
                            )}
                          </div>
                        </div>
                      </label>
                    </li>
                  );
                })}
              </ul>
            )}

            {columns.length > 0 && (
              <div className="mt-3 flex justify-between text-xs text-muted-foreground">
                <button className="hover:text-primary" onClick={() => setSelected(new Set(columns.map((c) => c.index)))}>Select all</button>
                <button className="hover:text-primary" onClick={() => setSelected(new Set())}>Clear</button>
              </div>
            )}
          </Card>
        </aside>
      </main>

      <footer className="border-t border-primary/10 bg-card/40 py-6 backdrop-blur-md">
        <div className="container mx-auto px-6 text-center text-sm text-muted-foreground">
          Upload • Convert • Download
        </div>
      </footer>
    </div>
  );
}

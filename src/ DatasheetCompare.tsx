import { useState, useCallback, useRef } from "react";
import * as XLSX from "xlsx-js-style";
import {
  Upload,
  Loader2,
  Trash2,
  Download,
  AlertCircle,
  CheckCircle2,
  FileText,
  GitCompare,
  Sparkles,
} from "lucide-react";

interface Spec {
  key: string;
  value: string;
}

interface Sheet {
  id: string;
  fileName: string;
  imagePreview: string;
  mimeType: string;
  base64: string;
  status: "pending" | "processing" | "done" | "error";
  error?: string;
  productName: string;
  reference: string;
  manufacturer: string;
  specs: Spec[];
}

const NAVY = "#1E3A5F";
const BLUE = "#2563EB";
const AMBER = "#F59E0B";

function uid() {
  return Math.random().toString(36).slice(2, 10);
}

async function fileToBase64(file: File): Promise<string> {
  const buffer = await file.arrayBuffer();
  let bytes = new Uint8Array(buffer);

  // Corrige un défaut d'en-tête PDF observé sur certains exports (ex: "%%PDF-" au lieu de "%PDF-").
  if (file.type === "application/pdf") {
    const header = new TextDecoder().decode(bytes.slice(0, 8));
    if (header.startsWith("%%PDF-")) {
      bytes = bytes.slice(1);
    }
  }

  let binary = "";
  const chunkSize = 0x8000;
  for (let i = 0; i < bytes.length; i += chunkSize) {
    binary += String.fromCharCode(...bytes.subarray(i, i + chunkSize));
  }
  return btoa(binary);
}

async function callExtract(prompt: string, base64?: string, mimeType?: string) {
  const response = await fetch("/api/extract", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ base64, mimeType, prompt }),
  });

  if (!response.ok) {
    const errBody = await response.json().catch(() => ({}));
    const message =
      typeof errBody.error === "string"
        ? errBody.error
        : errBody.error?.message || errBody.message || `Erreur API (${response.status})`;
    throw new Error(message);
  }
  const data = await response.json();
  const textBlock = (data.content || []).find((b: any) => b.type === "text");
  if (!textBlock) throw new Error("Réponse vide du modèle");
  return textBlock.text as string;
}

async function extractSpecs(base64: string, mimeType: string) {
  const prompt = `Tu regardes une fiche technique produit (datasheet) — pompe, moteur, ou équipement industriel.
Extrait :
- le nom / la désignation du produit
- la référence ou le code produit
- le fabricant si visible
- TOUTES les caractéristiques techniques présentes sur la fiche, sous forme de paires clé/valeur, avec l'unité incluse
  dans la valeur (ex: "5.5 kW", "230V", "50 Hz", "IP55"). Utilise des libellés clairs et cohérents en français
  (ex: "Puissance nominale", "Débit nominal", "Hauteur manométrique totale (HMT)", "Tension d'alimentation",
  "Fréquence", "Vitesse de rotation", "Diamètre de raccordement", "Matériau", "Poids", "Dimensions",
  "Indice de protection", "Classe d'isolation", "Température de fonctionnement", "Pression maximale") —
  inclus tout ce qui figure réellement sur la fiche, même si ce n'est pas dans cette liste d'exemples.

Réponds STRICTEMENT en JSON, sans texte autour, sans balises markdown, selon ce schéma exact :

{
  "product_name": "string",
  "reference": "string",
  "manufacturer": "string",
  "specs": [{ "key": "string", "value": "string" }]
}

N'invente aucune donnée. Si un champ n'est pas lisible, mets une chaîne vide.`;

  const text = await callExtract(prompt, base64, mimeType);
  const cleaned = text.replace(/```json|```/g, "").trim();
  const parsed = JSON.parse(cleaned);

  const specs: Spec[] = (parsed.specs || []).map((s: any) => ({
    key: String(s.key ?? "").trim(),
    value: String(s.value ?? "").trim(),
  }));

  return {
    productName: String(parsed.product_name ?? ""),
    reference: String(parsed.reference ?? ""),
    manufacturer: String(parsed.manufacturer ?? ""),
    specs,
  };
}

interface ComparisonRow {
  key: string;
  values: string[];
  differs: boolean;
}

function buildComparisonRows(sheets: Sheet[]): ComparisonRow[] {
  const order: string[] = [];
  const displayLabel: Record<string, string> = {};
  const valuesByKey: Record<string, string[]> = {};

  sheets.forEach((sheet, sheetIndex) => {
    sheet.specs.forEach((spec) => {
      const norm = spec.key.toLowerCase().trim();
      if (!norm) return;
      if (!(norm in valuesByKey)) {
        valuesByKey[norm] = new Array(sheets.length).fill("");
        order.push(norm);
        displayLabel[norm] = spec.key;
      }
      valuesByKey[norm][sheetIndex] = spec.value;
    });
  });

  return order.map((norm) => {
    const values = valuesByKey[norm];
    const nonEmpty = values.filter((v) => v.trim() !== "");
    const distinct = new Set(nonEmpty.map((v) => v.toLowerCase().trim()));
    const differs = distinct.size > 1 || nonEmpty.length !== values.length;
    return { key: displayLabel[norm], values, differs };
  });
}

async function generateSummary(sheets: Sheet[]): Promise<string> {
  const payload = sheets.map((s) => ({
    product_name: s.productName,
    reference: s.reference,
    manufacturer: s.manufacturer,
    specs: s.specs,
  }));

  const prompt = `Voici les caractéristiques techniques extraites de ${sheets.length} fiches produit différentes, au format JSON :
${JSON.stringify(payload)}

Rédige une synthèse concise (5 à 10 puces maximum) des DIFFÉRENCES MAJEURES entre ces produits — privilégie les écarts
significatifs pouvant influencer un choix technique ou commercial (puissance, débit, dimensions, tension, compatibilité,
matériaux, prix implicite...). Ignore les différences purement cosmétiques ou de formulation. Réponds en français,
en liste à puces commençant par "-", sans préambule ni conclusion.`;

  return callExtract(prompt);
}

const thinBorder = (rgb: string) => {
  const side = { style: "thin", color: { rgb } };
  return { top: side, bottom: side, left: side, right: side };
};

function downloadComparison(sheets: Sheet[], rows: ComparisonRow[]) {
  const headerStyle = {
    font: { name: "Arial", sz: 9, bold: true, color: { rgb: "1E3A5F" } },
    fill: { fgColor: { rgb: "BFDBFE" } },
    alignment: { horizontal: "center", vertical: "center", wrapText: true },
    border: thinBorder("93C5FD"),
  };
  const keyStyle = {
    font: { name: "Arial", sz: 9, bold: true, color: { rgb: "1E3A5F" } },
    fill: { fgColor: { rgb: "F8FAFC" } },
    alignment: { vertical: "center" },
    border: thinBorder("E5EAF0"),
  };
  const cellStyle = {
    font: { name: "Arial", sz: 9 },
    alignment: { vertical: "center" },
    border: thinBorder("E5EAF0"),
  };
  const diffCellStyle = {
    font: { name: "Arial", sz: 9, bold: true, color: { rgb: "92400E" } },
    fill: { fgColor: { rgb: "FEF3C7" } },
    alignment: { vertical: "center" },
    border: thinBorder("FCD34D"),
  };

  const ws: XLSX.WorkSheet = {};
  const put = (row: number, col: number, v: string, s: any) => {
    const ref = XLSX.utils.encode_cell({ r: row - 1, c: col - 1 });
    ws[ref] = { t: "s", v, s };
  };

  put(1, 1, "Caractéristique", headerStyle);
  sheets.forEach((sheet, i) => {
    put(1, i + 2, `${sheet.productName || sheet.fileName}${sheet.reference ? " — " + sheet.reference : ""}`, headerStyle);
  });

  rows.forEach((row, i) => {
    const r = i + 2;
    put(r, 1, row.key, keyStyle);
    row.values.forEach((v, c) => {
      put(r, c + 2, v || "—", row.differs ? diffCellStyle : cellStyle);
    });
  });

  ws["!cols"] = [{ wch: 32 }, ...sheets.map(() => ({ wch: 24 }))];
  ws["!ref"] = XLSX.utils.encode_range({
    s: { r: 0, c: 0 },
    e: { r: rows.length, c: sheets.length },
  });

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Comparaison");
  XLSX.writeFile(wb, "Comparaison_fiches_techniques.xlsx");
}

export default function DatasheetCompare() {
  const [sheets, setSheets] = useState<Sheet[]>([]);
  const [isProcessing, setIsProcessing] = useState(false);
  const [summary, setSummary] = useState("");
  const [summaryLoading, setSummaryLoading] = useState(false);
  const [summaryError, setSummaryError] = useState("");
  const inputRef = useRef<HTMLInputElement>(null);

  const addFiles = useCallback(async (files: FileList | null) => {
    if (!files || files.length === 0) return;
    const newSheets: Sheet[] = [];
    for (const file of Array.from(files)) {
      const isImage = file.type.startsWith("image/");
      const isPdf = file.type === "application/pdf";
      if (!isImage && !isPdf) continue;
      const base64 = await fileToBase64(file);
      newSheets.push({
        id: uid(),
        fileName: file.name,
        imagePreview: isImage ? URL.createObjectURL(file) : "",
        mimeType: file.type,
        base64,
        status: "pending",
        productName: "",
        reference: "",
        manufacturer: "",
        specs: [],
      });
    }
    setSheets((prev) => [...prev, ...newSheets]);
    setSummary("");
  }, []);

  const removeSheet = (id: string) => {
    setSheets((prev) => prev.filter((s) => s.id !== id));
    setSummary("");
  };

  const processAll = async () => {
    setIsProcessing(true);
    const pending = sheets.filter((s) => s.status === "pending" || s.status === "error");
    for (const sheet of pending) {
      setSheets((prev) =>
        prev.map((s) => (s.id === sheet.id ? { ...s, status: "processing", error: undefined } : s))
      );
      try {
        const result = await extractSpecs(sheet.base64, sheet.mimeType);
        setSheets((prev) => prev.map((s) => (s.id === sheet.id ? { ...s, ...result, status: "done" } : s)));
      } catch (err: any) {
        setSheets((prev) =>
          prev.map((s) =>
            s.id === sheet.id ? { ...s, status: "error", error: err?.message || "Échec de l'analyse" } : s
          )
        );
      }
    }
    setIsProcessing(false);
  };

  const doneSheets = sheets.filter((s) => s.status === "done");
  const rows = doneSheets.length >= 2 ? buildComparisonRows(doneSheets) : [];
  const diffCount = rows.filter((r) => r.differs).length;
  const hasPending = sheets.some((s) => s.status === "pending" || s.status === "error");

  const handleSummary = async () => {
    setSummaryLoading(true);
    setSummaryError("");
    try {
      const text = await generateSummary(doneSheets);
      setSummary(text.trim());
    } catch (err: any) {
      setSummaryError(err?.message || "Échec de la synthèse");
    }
    setSummaryLoading(false);
  };

  return (
    <div className="space-y-6">
      <div className="flex items-center gap-3">
        <div
          className="w-10 h-10 rounded flex items-center justify-center flex-shrink-0"
          style={{ backgroundColor: NAVY }}
        >
          <GitCompare className="w-5 h-5 text-white" />
        </div>
        <div>
          <h1 className="text-lg font-semibold" style={{ color: NAVY }}>
            Comparateur de fiches techniques
          </h1>
          <p className="text-sm text-gray-500">
            Déposez 2 fiches techniques ou plus (PDF/images), l'appli aligne les caractéristiques et repère les écarts.
          </p>
        </div>
      </div>

      <div
        className="border-2 border-dashed rounded-lg p-8 text-center cursor-pointer transition-colors"
        style={{ borderColor: "#CBD5E1" }}
        onClick={() => inputRef.current?.click()}
        onDragOver={(e) => e.preventDefault()}
        onDrop={(e) => {
          e.preventDefault();
          addFiles(e.dataTransfer.files);
        }}
      >
        <Upload className="w-8 h-8 mx-auto mb-2 text-gray-400" />
        <p className="text-sm text-gray-600">
          Glissez plusieurs fiches techniques ici, ou{" "}
          <span style={{ color: BLUE }}>cliquez pour en choisir plusieurs</span>
        </p>
        <p className="text-xs text-gray-400 mt-1">PNG, JPG, PDF — au moins 2 fiches pour comparer</p>
        <input
          ref={inputRef}
          type="file"
          accept="image/*,application/pdf"
          multiple
          className="hidden"
          onChange={(e) => {
            addFiles(e.target.files);
            e.target.value = "";
          }}
        />
      </div>

      {sheets.length > 0 && (
        <div className="flex items-center gap-3">
          <span className="text-xs text-gray-500">
            {sheets.length} fiche{sheets.length > 1 ? "s" : ""} ajoutée{sheets.length > 1 ? "s" : ""}
          </span>
          <button
            onClick={processAll}
            disabled={!hasPending || isProcessing}
            className="flex items-center gap-2 px-4 py-2 rounded text-sm font-medium text-white disabled:opacity-40"
            style={{ backgroundColor: BLUE }}
          >
            {isProcessing ? <Loader2 className="w-4 h-4 animate-spin" /> : <GitCompare className="w-4 h-4" />}
            Comparer les fiches
          </button>
        </div>
      )}

      <div className="space-y-3">
        {sheets.map((sheet) => (
          <div
            key={sheet.id}
            className="flex items-center gap-3 px-4 py-3 border rounded-lg"
            style={{ borderColor: "#E2E8F0", backgroundColor: "#F8FAFC" }}
          >
            {sheet.imagePreview ? (
              <img
                src={sheet.imagePreview}
                alt={sheet.fileName}
                className="w-12 h-12 object-cover rounded border"
                style={{ borderColor: "#E2E8F0" }}
              />
            ) : (
              <div
                className="w-12 h-12 flex items-center justify-center rounded border flex-shrink-0"
                style={{ borderColor: "#E2E8F0", backgroundColor: "#F1F5F9" }}
              >
                <FileText className="w-5 h-5 text-gray-400" />
              </div>
            )}
            <div className="flex-1 min-w-0">
              <p className="text-sm font-medium truncate" style={{ color: NAVY }}>
                {sheet.fileName}
              </p>
              <p className="text-xs text-gray-500 truncate">
                {sheet.status === "pending" && "En attente d'analyse"}
                {sheet.status === "processing" && "Analyse en cours…"}
                {sheet.status === "done" &&
                  `${sheet.productName || "—"}${sheet.reference ? " · " + sheet.reference : ""} · ${sheet.specs.length} caractéristique(s)`}
                {sheet.status === "error" && (sheet.error || "Erreur")}
              </p>
            </div>
            {sheet.status === "processing" && <Loader2 className="w-4 h-4 animate-spin text-gray-400" />}
            {sheet.status === "done" && <CheckCircle2 className="w-4 h-4" style={{ color: "#16A34A" }} />}
            {sheet.status === "error" && <AlertCircle className="w-4 h-4 text-red-500" />}
            <button onClick={() => removeSheet(sheet.id)} aria-label="Supprimer" className="text-gray-400 hover:text-red-500">
              <Trash2 className="w-4 h-4" />
            </button>
          </div>
        ))}
      </div>

      {sheets.length === 0 && (
        <p className="text-xs text-gray-400 text-center">
          Aucune fiche pour le moment. Ajoutez au moins deux fiches techniques à comparer.
        </p>
      )}

      {doneSheets.length === 1 && (
        <p className="text-xs text-gray-400 text-center">Ajoutez une deuxième fiche pour lancer la comparaison.</p>
      )}

      {rows.length > 0 && (
        <div className="space-y-4">
          <div className="flex items-center justify-between flex-wrap gap-2">
            <div
              className="flex items-center gap-2 px-3 py-1.5 rounded text-xs font-medium"
              style={{ backgroundColor: "#FEF3C7", color: "#92400E" }}
            >
              <span className="w-2 h-2 rounded-full" style={{ backgroundColor: AMBER }} />
              {diffCount} caractéristique{diffCount > 1 ? "s différentes" : " différente"} sur {rows.length}
            </div>
            <div className="flex items-center gap-2">
              <button
                onClick={handleSummary}
                disabled={summaryLoading}
                className="flex items-center gap-2 px-3 py-1.5 rounded text-xs font-medium border disabled:opacity-40"
                style={{ borderColor: NAVY, color: NAVY }}
              >
                {summaryLoading ? <Loader2 className="w-3.5 h-3.5 animate-spin" /> : <Sparkles className="w-3.5 h-3.5" />}
                Synthèse des différences majeures
              </button>
              <button
                onClick={() => downloadComparison(doneSheets, rows)}
                className="flex items-center gap-2 px-3 py-1.5 rounded text-xs font-medium text-white"
                style={{ backgroundColor: NAVY }}
              >
                <Download className="w-3.5 h-3.5" />
                Exporter en Excel
              </button>
            </div>
          </div>

          {summaryError && <p className="text-xs text-red-500">{summaryError}</p>}

          {summary && (
            <div className="rounded-lg p-4 text-sm" style={{ backgroundColor: "#DBEAFE", color: NAVY }}>
              {summary.split("\n").map((line, i) => {
                const trimmed = line.trim();
                if (!trimmed) return null;
                const isBullet = trimmed.startsWith("-") || trimmed.startsWith("•");
                return (
                  <p key={i} className={isBullet ? "flex gap-2 mb-1" : "mb-1 font-medium"}>
                    {isBullet && <span>•</span>}
                    <span>{isBullet ? trimmed.replace(/^[-•]\s*/, "") : trimmed}</span>
                  </p>
                );
              })}
            </div>
          )}

          <div className="overflow-x-auto border rounded-lg" style={{ borderColor: "#E2E8F0" }}>
            <table className="w-full text-xs">
              <thead>
                <tr style={{ backgroundColor: "#BFDBFE" }}>
                  <th className="text-left p-2 sticky left-0" style={{ color: NAVY, backgroundColor: "#BFDBFE" }}>
                    Caractéristique
                  </th>
                  {doneSheets.map((sheet) => (
                    <th key={sheet.id} className="text-left p-2" style={{ color: NAVY }}>
                      {sheet.productName || sheet.fileName}
                      {sheet.reference && <div className="font-normal text-[11px] opacity-75">{sheet.reference}</div>}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {rows.map((row, i) => (
                  <tr
                    key={i}
                    style={{ backgroundColor: row.differs ? "#FEF3C7" : i % 2 === 0 ? "#FFFFFF" : "#F8FAFC" }}
                  >
                    <td
                      className="p-2 font-medium sticky left-0"
                      style={{
                        color: NAVY,
                        backgroundColor: row.differs ? "#FEF3C7" : i % 2 === 0 ? "#FFFFFF" : "#F8FAFC",
                      }}
                    >
                      {row.key}
                    </td>
                    {row.values.map((v, c) => (
                      <td key={c} className="p-2" style={{ color: row.differs ? "#92400E" : "#374151" }}>
                        {v || "—"}
                      </td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}
    </div>
  );
}

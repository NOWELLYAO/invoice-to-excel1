import { useState, useCallback, useRef } from "react";
import * as XLSX from "xlsx-js-style";
import {
  Upload,
  FileSpreadsheet,
  Loader2,
  Trash2,
  Download,
  AlertCircle,
  CheckCircle2,
  Plus,
  Minus,
  FileText,
} from "lucide-react";

interface Item {
  id: string;
  pn: string;
  desig: string;
  qty: number;
  total: number;
}

interface InvoiceEntry {
  id: string;
  fileName: string;
  imagePreview: string;
  mimeType: string;
  base64: string;
  status: "pending" | "processing" | "done" | "error";
  error?: string;
  invoiceNumber: string;
  billingDate: string;
  payerName: string;
  payerAddress: string;
  items: Item[];
}

const NAVY = "#1E3A5F";
const BLUE = "#2563EB";

function uid() {
  return Math.random().toString(36).slice(2, 10);
}

function fileToBase64(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve((reader.result as string).split(",")[1]);
    reader.onerror = () => reject(new Error("Lecture du fichier impossible"));
    reader.readAsDataURL(file);
  });
}

async function extractInvoiceData(base64: string, mimeType: string) {
  const prompt = `Tu regardes un document de facturation. Il peut s'agir soit :
(a) d'une capture d'écran SAP (F2 Sales Invoice - Overview of Billing Items), soit
(b) d'une facture commerciale multi-pages (Commercial Invoice, PDF, souvent Grundfos).

Identifie automatiquement le type de document et extrais les données réellement présentes.
Si c'est un PDF avec plusieurs pages, parcours TOUTES les pages et regroupe tous les articles
dans une seule liste, sans doublons, sans recopier les lignes de total de page intermédiaire.

Réponds STRICTEMENT en JSON, sans texte autour, sans balises markdown, selon ce schéma exact :

{
  "invoice_number": "string, numéro de facture (F2 Sales Invoice ou Commercial Invoice)",
  "billing_date": "string au format JJ/MM/AAAA",
  "payer_name": "string, nom du client/payeur (champ 'Payer' en SAP, ou société cliente / adresse de livraison sur une Commercial Invoice)",
  "payer_address": "string, adresse du client si visible sinon vide",
  "items": [
    {
      "material_code": "string, référence produit (colonne 'Material', ou code article sur 'Material / description')",
      "description": "string, désignation du produit uniquement — ignore les lignes secondaires type 'ECCN code' ou 'Country of origin'",
      "quantity": nombre (colonne Invoiced Quantity ou Quantity),
      "net_value": nombre (colonne Net Value pour SAP, ou NET Amount pour une Commercial Invoice — montant net de la ligne, sans séparateur de milliers, point décimal)
    }
  ]
}

N'invente aucune donnée. Si un champ n'est pas lisible, mets une chaîne vide ou 0.`;

  // Appelle notre propre fonction serveur (api/extract.ts) qui détient la clé API
  // Anthropic côté serveur. Le navigateur ne voit jamais la clé.
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

  const cleaned = textBlock.text.replace(/```json|```/g, "").trim();
  const parsed = JSON.parse(cleaned);

  const items: Item[] = (parsed.items || []).map((it: any) => ({
    id: uid(),
    pn: String(it.material_code ?? ""),
    desig: String(it.description ?? ""),
    qty: Number(it.quantity) || 0,
    total: Number(it.net_value) || 0,
  }));

  return {
    invoiceNumber: String(parsed.invoice_number ?? ""),
    billingDate: String(parsed.billing_date ?? ""),
    payerName: String(parsed.payer_name ?? ""),
    payerAddress: String(parsed.payer_address ?? ""),
    items,
  };
}

const thinBorder = (rgb: string) => {
  const side = { style: "thin", color: { rgb } };
  return { top: side, bottom: side, left: side, right: side };
};

const STYLE = {
  title: {
    font: { name: "Arial", sz: 14, bold: true, color: { rgb: "1D4ED8" } },
    alignment: { vertical: "center" },
  },
  ref: {
    font: { name: "Arial", sz: 9, color: { rgb: "6B7280" } },
    alignment: { vertical: "center" },
  },
  header: {
    font: { name: "Arial", sz: 9, bold: true, color: { rgb: "1E3A5F" } },
    fill: { fgColor: { rgb: "BFDBFE" } },
    alignment: { horizontal: "center", vertical: "center" },
    border: thinBorder("93C5FD"),
  },
  cell: {
    font: { name: "Arial", sz: 9 },
    alignment: { vertical: "center" },
    border: thinBorder("E5EAF0"),
  },
  cellCenter: {
    font: { name: "Arial", sz: 9 },
    alignment: { horizontal: "center", vertical: "center" },
    border: thinBorder("E5EAF0"),
  },
  pu: {
    font: { name: "Arial", sz: 9, color: { rgb: "374151" } },
    fill: { fgColor: { rgb: "F8FAFC" } },
    alignment: { horizontal: "right", vertical: "center" },
    border: thinBorder("E5EAF0"),
    numFmt: "#,##0.00",
  },
  total: {
    font: { name: "Arial", sz: 9 },
    fill: { fgColor: { rgb: "F0FDF4" } },
    alignment: { horizontal: "right", vertical: "center" },
    border: thinBorder("E5EAF0"),
    numFmt: "#,##0.00",
  },
  totalLabel: {
    font: { name: "Arial", sz: 10, bold: true, color: { rgb: "1E3A5F" } },
    fill: { fgColor: { rgb: "DBEAFE" } },
    alignment: { vertical: "center" },
  },
  totalValue: {
    font: { name: "Arial", sz: 10, bold: true, color: { rgb: "1E3A5F" } },
    fill: { fgColor: { rgb: "DBEAFE" } },
    alignment: { horizontal: "right", vertical: "center" },
    numFmt: "#,##0.00",
  },
  footer: {
    font: { name: "Arial", sz: 9, color: { rgb: "6B7280" } },
  },
};

function buildWorkbook(entry: InvoiceEntry): XLSX.WorkBook {
  const headerRow = 4;
  const firstItemRow = 5;
  const lastItemRow = firstItemRow + entry.items.length - 1;
  const totalRow = lastItemRow + 2;
  const footerRow = totalRow + 2;

  const ws: XLSX.WorkSheet = {};
  const put = (row: number, col: number, v: string | number | null, s: any, f?: string) => {
    const ref = XLSX.utils.encode_cell({ r: row - 1, c: col - 1 });
    const cell: any = { s };
    if (f) {
      cell.t = "n";
      cell.f = f;
    } else if (typeof v === "number") {
      cell.t = "n";
      cell.v = v;
    } else {
      cell.t = "s";
      cell.v = v ?? "";
    }
    ws[ref] = cell;
  };

  put(1, 1, `FACTURE N° ${entry.invoiceNumber}`, STYLE.title);
  put(
    2,
    1,
    `Réf : ${entry.invoiceNumber}   |   Date : ${entry.billingDate}   |   Payeur : ${entry.payerName}`,
    STYLE.ref
  );

  const headers = ["P/N", "Désignation", "P.U. (€)", "Qté", "Total (€)", "Statut"];
  headers.forEach((h, i) => put(headerRow, i + 1, h, STYLE.header));

  entry.items.forEach((item, i) => {
    const r = firstItemRow + i;
    const pu = item.qty ? Math.round((item.total / item.qty) * 100) / 100 : 0;
    put(r, 1, item.pn, STYLE.cell);
    put(r, 2, item.desig, STYLE.cell);
    put(r, 3, pu, STYLE.pu);
    put(r, 4, item.qty, STYLE.cellCenter);
    put(r, 5, null, STYLE.total, `C${r}*D${r}`);
    put(r, 6, "Facturé", STYLE.cell);
  });

  put(totalRow, 1, "TOTAL HT", STYLE.totalLabel);
  for (let c = 2; c <= 4; c++) put(totalRow, c, "", STYLE.totalLabel);
  put(totalRow, 5, null, STYLE.totalValue, `SUM(E${firstItemRow}:E${lastItemRow})`);

  put(
    footerRow,
    1,
    `Payeur : ${entry.payerName}${entry.payerAddress ? ", " + entry.payerAddress : ""}`,
    STYLE.footer
  );
  for (let c = 2; c <= 4; c++) put(footerRow, c, "", STYLE.footer);

  ws["!merges"] = [
    { s: { r: 0, c: 0 }, e: { r: 0, c: 5 } },
    { s: { r: 1, c: 0 }, e: { r: 1, c: 5 } },
    { s: { r: totalRow - 1, c: 0 }, e: { r: totalRow - 1, c: 3 } },
    { s: { r: footerRow - 1, c: 0 }, e: { r: footerRow - 1, c: 3 } },
  ];

  ws["!cols"] = [{ wch: 11 }, { wch: 38 }, { wch: 11 }, { wch: 5 }, { wch: 11 }, { wch: 17 }];
  ws["!rows"] = [{ hpt: 26 }, { hpt: 16 }];
  ws["!ref"] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: footerRow - 1, c: 5 } });

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Facture");
  return wb;
}

function invoiceTotal(entry: InvoiceEntry): number {
  return entry.items.reduce((sum, it) => sum + (Number(it.total) || 0), 0);
}

function formatEUR(value: number): string {
  return new Intl.NumberFormat("fr-FR", { style: "currency", currency: "EUR" }).format(value || 0);
}

function downloadWorkbook(entry: InvoiceEntry) {
  const wb = buildWorkbook(entry);
  const safeDate = (entry.billingDate || "").split("/").reverse().join("-");
  const fileName = `Facture_${entry.invoiceNumber || "sans_numero"}${safeDate ? "_" + safeDate : ""}.xlsx`;
  XLSX.writeFile(wb, fileName);
}

export default function InvoiceToExcel() {
  const [entries, setEntries] = useState<InvoiceEntry[]>([]);
  const [isProcessing, setIsProcessing] = useState(false);
  const inputRef = useRef<HTMLInputElement>(null);

  const addFiles = useCallback(async (files: FileList | null) => {
    if (!files || files.length === 0) return;
    const newEntries: InvoiceEntry[] = [];
    for (const file of Array.from(files)) {
      const isImage = file.type.startsWith("image/");
      const isPdf = file.type === "application/pdf";
      if (!isImage && !isPdf) continue;
      const base64 = await fileToBase64(file);
      newEntries.push({
        id: uid(),
        fileName: file.name,
        imagePreview: isImage ? URL.createObjectURL(file) : "",
        mimeType: file.type,
        base64,
        status: "pending",
        invoiceNumber: "",
        billingDate: "",
        payerName: "",
        payerAddress: "",
        items: [],
      });
    }
    setEntries((prev) => [...prev, ...newEntries]);
  }, []);

  const removeEntry = (id: string) => {
    setEntries((prev) => prev.filter((e) => e.id !== id));
  };

  const processAll = async () => {
    setIsProcessing(true);
    const pending = entries.filter((e) => e.status === "pending" || e.status === "error");
    for (const entry of pending) {
      setEntries((prev) =>
        prev.map((e) => (e.id === entry.id ? { ...e, status: "processing", error: undefined } : e))
      );
      try {
        const result = await extractInvoiceData(entry.base64, entry.mimeType);
        setEntries((prev) =>
          prev.map((e) => (e.id === entry.id ? { ...e, ...result, status: "done" } : e))
        );
      } catch (err: any) {
        setEntries((prev) =>
          prev.map((e) =>
            e.id === entry.id ? { ...e, status: "error", error: err?.message || "Échec de l'analyse" } : e
          )
        );
      }
    }
    setIsProcessing(false);
  };

  const updateEntry = (id: string, patch: Partial<InvoiceEntry>) => {
    setEntries((prev) => prev.map((e) => (e.id === id ? { ...e, ...patch } : e)));
  };

  const updateItem = (entryId: string, itemId: string, patch: Partial<Item>) => {
    setEntries((prev) =>
      prev.map((e) =>
        e.id === entryId
          ? { ...e, items: e.items.map((it) => (it.id === itemId ? { ...it, ...patch } : it)) }
          : e
      )
    );
  };

  const addItem = (entryId: string) => {
    setEntries((prev) =>
      prev.map((e) =>
        e.id === entryId
          ? { ...e, items: [...e.items, { id: uid(), pn: "", desig: "", qty: 1, total: 0 }] }
          : e
      )
    );
  };

  const removeItem = (entryId: string, itemId: string) => {
    setEntries((prev) =>
      prev.map((e) =>
        e.id === entryId ? { ...e, items: e.items.filter((it) => it.id !== itemId) } : e
      )
    );
  };

  const downloadAll = () => {
    const done = entries.filter((e) => e.status === "done");
    done.forEach((entry, i) => {
      setTimeout(() => downloadWorkbook(entry), i * 250);
    });
  };

  const hasPending = entries.some((e) => e.status === "pending" || e.status === "error");
  const hasDone = entries.some((e) => e.status === "done");

  return (
    <div className="w-full max-w-3xl mx-auto p-6 space-y-6 bg-white" style={{ fontFamily: "Arial, sans-serif" }}>
      <div className="flex items-center gap-3">
        <div
          className="w-10 h-10 rounded flex items-center justify-center flex-shrink-0"
          style={{ backgroundColor: NAVY }}
        >
          <FileSpreadsheet className="w-5 h-5 text-white" />
        </div>
        <div>
          <h1 className="text-lg font-semibold" style={{ color: NAVY }}>
            Captures SAP vers Excel
          </h1>
          <p className="text-sm text-gray-500">
            Déposez vos captures SAP ou vos factures PDF, l'extraction se fait automatiquement.
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
          Glissez plusieurs captures ou PDF de factures ici, ou{" "}
          <span style={{ color: BLUE }}>cliquez pour en choisir plusieurs</span>
        </p>
        <p className="text-xs text-gray-400 mt-1">
          PNG, JPG, PDF — sélectionnez plusieurs fichiers avec Ctrl (ou Cmd sur Mac) + clic
        </p>
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

      {entries.length > 0 && (
        <div className="flex items-center gap-3">
          <span className="text-xs text-gray-500">
            {entries.length} capture{entries.length > 1 ? "s" : ""} ajoutée{entries.length > 1 ? "s" : ""}
          </span>
          <button
            onClick={processAll}
            disabled={!hasPending || isProcessing}
            className="flex items-center gap-2 px-4 py-2 rounded text-sm font-medium text-white disabled:opacity-40"
            style={{ backgroundColor: BLUE }}
          >
            {isProcessing ? (
              <Loader2 className="w-4 h-4 animate-spin" />
            ) : (
              <FileSpreadsheet className="w-4 h-4" />
            )}
            Analyser les captures
          </button>
          {hasDone && (
            <button
              onClick={downloadAll}
              className="flex items-center gap-2 px-4 py-2 rounded text-sm font-medium border"
              style={{ borderColor: NAVY, color: NAVY }}
            >
              <Download className="w-4 h-4" />
              Tout télécharger
            </button>
          )}
        </div>
      )}

      <div className="space-y-4">
        {entries.map((entry) => (
          <div key={entry.id} className="border rounded-lg overflow-hidden" style={{ borderColor: "#E2E8F0" }}>
            <div className="flex items-center gap-3 px-4 py-3" style={{ backgroundColor: "#F8FAFC" }}>
              {entry.imagePreview ? (
                <img
                  src={entry.imagePreview}
                  alt={entry.fileName}
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
                  {entry.fileName}
                </p>
                <p className="text-xs text-gray-500">
                  {entry.status === "pending" && "En attente d'analyse"}
                  {entry.status === "processing" && "Analyse en cours…"}
                  {entry.status === "done" &&
                    `Facture N° ${entry.invoiceNumber || "—"} · ${entry.items.length} article(s) · Total ${formatEUR(
                      invoiceTotal(entry)
                    )}`}
                  {entry.status === "error" && (entry.error || "Erreur")}
                </p>
              </div>
              {entry.status === "processing" && <Loader2 className="w-4 h-4 animate-spin text-gray-400" />}
              {entry.status === "done" && <CheckCircle2 className="w-4 h-4" style={{ color: "#16A34A" }} />}
              {entry.status === "error" && <AlertCircle className="w-4 h-4 text-red-500" />}
              <button onClick={() => removeEntry(entry.id)} aria-label="Supprimer" className="text-gray-400 hover:text-red-500">
                <Trash2 className="w-4 h-4" />
              </button>
            </div>

            {entry.status === "done" && (
              <div className="p-4 space-y-3">
                <div className="grid grid-cols-2 gap-3">
                  <label className="text-xs text-gray-500 flex flex-col gap-1">
                    N° de facture
                    <input
                      value={entry.invoiceNumber}
                      onChange={(e) => updateEntry(entry.id, { invoiceNumber: e.target.value })}
                      className="border rounded px-2 py-1 text-sm"
                      style={{ borderColor: "#E2E8F0" }}
                    />
                  </label>
                  <label className="text-xs text-gray-500 flex flex-col gap-1">
                    Date
                    <input
                      value={entry.billingDate}
                      onChange={(e) => updateEntry(entry.id, { billingDate: e.target.value })}
                      className="border rounded px-2 py-1 text-sm"
                      style={{ borderColor: "#E2E8F0" }}
                    />
                  </label>
                  <label className="text-xs text-gray-500 flex flex-col gap-1 col-span-2">
                    Payeur
                    <input
                      value={entry.payerName}
                      onChange={(e) => updateEntry(entry.id, { payerName: e.target.value })}
                      className="border rounded px-2 py-1 text-sm"
                      style={{ borderColor: "#E2E8F0" }}
                    />
                  </label>
                </div>

                <div
                  className="flex items-center justify-between px-3 py-2 rounded"
                  style={{ backgroundColor: "#DBEAFE" }}
                >
                  <span className="text-xs font-medium" style={{ color: NAVY }}>
                    Total facture
                  </span>
                  <span className="text-sm font-semibold" style={{ color: NAVY }}>
                    {formatEUR(invoiceTotal(entry))}
                  </span>
                </div>

                <table className="w-full text-xs">
                  <thead>
                    <tr style={{ backgroundColor: "#BFDBFE" }}>
                      <th className="text-left p-1.5" style={{ color: NAVY }}>P/N</th>
                      <th className="text-left p-1.5" style={{ color: NAVY }}>Désignation</th>
                      <th className="text-right p-1.5" style={{ color: NAVY }}>Qté</th>
                      <th className="text-right p-1.5" style={{ color: NAVY }}>Net (€)</th>
                      <th className="p-1.5"></th>
                    </tr>
                  </thead>
                  <tbody>
                    {entry.items.map((item) => (
                      <tr key={item.id} className="border-b" style={{ borderColor: "#F1F5F9" }}>
                        <td className="p-1">
                          <input
                            value={item.pn}
                            onChange={(e) => updateItem(entry.id, item.id, { pn: e.target.value })}
                            className="w-full border rounded px-1.5 py-1"
                            style={{ borderColor: "#E2E8F0" }}
                          />
                        </td>
                        <td className="p-1">
                          <input
                            value={item.desig}
                            onChange={(e) => updateItem(entry.id, item.id, { desig: e.target.value })}
                            className="w-full border rounded px-1.5 py-1"
                            style={{ borderColor: "#E2E8F0" }}
                          />
                        </td>
                        <td className="p-1">
                          <input
                            type="number"
                            value={item.qty}
                            onChange={(e) => updateItem(entry.id, item.id, { qty: Number(e.target.value) })}
                            className="w-16 border rounded px-1.5 py-1 text-right"
                            style={{ borderColor: "#E2E8F0" }}
                          />
                        </td>
                        <td className="p-1">
                          <input
                            type="number"
                            value={item.total}
                            onChange={(e) => updateItem(entry.id, item.id, { total: Number(e.target.value) })}
                            className="w-20 border rounded px-1.5 py-1 text-right"
                            style={{ borderColor: "#E2E8F0" }}
                          />
                        </td>
                        <td className="p-1 text-center">
                          <button onClick={() => removeItem(entry.id, item.id)} aria-label="Retirer la ligne" className="text-gray-400 hover:text-red-500">
                            <Minus className="w-3.5 h-3.5" />
                          </button>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>

                <div className="flex items-center justify-between">
                  <button
                    onClick={() => addItem(entry.id)}
                    className="flex items-center gap-1 text-xs"
                    style={{ color: BLUE }}
                  >
                    <Plus className="w-3.5 h-3.5" /> Ajouter une ligne
                  </button>
                  <button
                    onClick={() => downloadWorkbook(entry)}
                    className="flex items-center gap-2 px-3 py-1.5 rounded text-xs font-medium text-white"
                    style={{ backgroundColor: NAVY }}
                  >
                    <Download className="w-3.5 h-3.5" />
                    Télécharger le fichier Excel
                  </button>
                </div>
              </div>
            )}
          </div>
        ))}
      </div>

      {entries.length === 0 && (
        <p className="text-xs text-gray-400 text-center">
          Aucune capture pour le moment. Ajoutez une ou plusieurs images ou PDF de factures ci-dessus.
        </p>
      )}
    </div>
  );
}

import { useState, useCallback, useRef } from "react";
import * as XLSX from "xlsx";
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
  const prompt = `Tu regardes une capture d'écran SAP d'une facture de vente (F2 Sales Invoice - Overview of Billing Items).
Extrait uniquement les données réellement visibles à l'écran et réponds STRICTEMENT en JSON, sans texte autour, sans balises markdown, selon ce schéma exact :

{
  "invoice_number": "string, ex 7293020306",
  "billing_date": "string au format JJ/MM/AAAA",
  "payer_name": "string, nom du payeur",
  "payer_address": "string, adresse du payeur si visible sinon vide",
  "items": [
    {
      "material_code": "string, colonne Material",
      "description": "string, colonne Item Description",
      "quantity": nombre (colonne Invoiced Quantity),
      "net_value": nombre (colonne Net Value de la ligne, sans séparateur de milliers, point décimal)
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
    throw new Error(errBody.error || `Erreur API (${response.status})`);
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

function buildWorkbook(entry: InvoiceEntry): XLSX.WorkBook {
  const headerRow = 4;
  const firstItemRow = 5;
  const lastItemRow = firstItemRow + entry.items.length - 1;
  const totalRow = lastItemRow + 2;
  const footerRow = totalRow + 2;

  const rows: any[][] = [];
  rows[0] = [`FACTURE N° ${entry.invoiceNumber}`];
  rows[1] = [
    `Réf : ${entry.invoiceNumber}   |   Date : ${entry.billingDate}   |   Payeur : ${entry.payerName}`,
  ];
  rows[2] = [];
  rows[headerRow - 1] = ["P/N", "Désignation", "P.U. (€)", "Qté", "Total (€)", "Statut"];

  entry.items.forEach((item, i) => {
    const r = firstItemRow + i - 1;
    const pu = item.qty ? Math.round((item.total / item.qty) * 100) / 100 : 0;
    rows[r] = [item.pn, item.desig, pu, item.qty, null, "Facturé"];
  });

  rows[totalRow - 1] = ["TOTAL HT", null, null, null, null];
  rows[footerRow - 1] = [`Payeur : ${entry.payerName}${entry.payerAddress ? ", " + entry.payerAddress : ""}`];

  const ws = XLSX.utils.aoa_to_sheet(rows);

  // Formulas
  for (let i = 0; i < entry.items.length; i++) {
    const r = firstItemRow + i;
    ws[`E${r}`] = { t: "n", f: `C${r}*D${r}` };
  }
  ws[`E${totalRow}`] = { t: "n", f: `SUM(E${firstItemRow}:E${lastItemRow})` };

  ws["!merges"] = [
    { s: { r: 0, c: 0 }, e: { r: 0, c: 5 } },
    { s: { r: 1, c: 0 }, e: { r: 1, c: 5 } },
    { s: { r: totalRow - 1, c: 0 }, e: { r: totalRow - 1, c: 3 } },
    { s: { r: footerRow - 1, c: 0 }, e: { r: footerRow - 1, c: 3 } },
  ];

  ws["!cols"] = [
    { wch: 11 },
    { wch: 38 },
    { wch: 11 },
    { wch: 5 },
    { wch: 11 },
    { wch: 17 },
  ];

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Facture");
  return wb;
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
      if (!file.type.startsWith("image/")) continue;
      const base64 = await fileToBase64(file);
      newEntries.push({
        id: uid(),
        fileName: file.name,
        imagePreview: URL.createObjectURL(file),
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
            Déposez vos captures de factures SAP (F2), l'extraction se fait automatiquement.
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
          Glissez vos captures ici, ou <span style={{ color: BLUE }}>cliquez pour parcourir</span>
        </p>
        <p className="text-xs text-gray-400 mt-1">PNG, JPG — plusieurs fichiers acceptés</p>
        <input
          ref={inputRef}
          type="file"
          accept="image/*"
          multiple
          className="hidden"
          onChange={(e) => addFiles(e.target.files)}
        />
      </div>

      {entries.length > 0 && (
        <div className="flex items-center gap-3">
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
              <img
                src={entry.imagePreview}
                alt={entry.fileName}
                className="w-12 h-12 object-cover rounded border"
                style={{ borderColor: "#E2E8F0" }}
              />
              <div className="flex-1 min-w-0">
                <p className="text-sm font-medium truncate" style={{ color: NAVY }}>
                  {entry.fileName}
                </p>
                <p className="text-xs text-gray-500">
                  {entry.status === "pending" && "En attente d'analyse"}
                  {entry.status === "processing" && "Analyse en cours…"}
                  {entry.status === "done" &&
                    `Facture N° ${entry.invoiceNumber || "—"} · ${entry.items.length} article(s)`}
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
          Aucune capture pour le moment. Ajoutez une ou plusieurs images de factures SAP ci-dessus.
        </p>
      )}
    </div>
  );
}

import React, { useMemo, useState } from "react";
import * as XLSX from "xlsx";

// --------------------------------------------------
// CAS / EC Checker – REACH SVHC & RoHS
// Ziel: Läuft sowohl in StackBlitz (React Fork) als auch lokal via npm
//
// WICHTIG:
// - Keine @/… Aliase, keine shadcn Imports
// - TE Logo wird aus /public geladen (CRA + StackBlitz kompatibel)
//   -> Lege die Datei in den Ordner "public" und referenziere sie über PUBLIC_URL.
//
// Excel-Zuordnung (FINAL):
// RoHS-Sheet:
//   Spalte C (Index 2): CAS
//   Spalte D (Index 3): Grenzwert in Gewichts-%
// SVHC-Sheet:
//   Spalte C (Index 2): CAS
//   Spalte D (Index 3): EC
//   Spalte E (Index 4): Grenzwert in Gewichts-%
//   Spalte G (Index 6): Reason for inclusion
// --------------------------------------------------

export default function App() {
  const [query, setQuery] = useState("");
  const [rohsList, setRohsList] = useState([]);
  const [svhcList, setSvhcList] = useState([]);
  const [fileLoaded, setFileLoaded] = useState(false);

  const [productWeight, setProductWeight] = useState("");
  const [substanceWeight, setSubstanceWeight] = useState("");
  const [logoOk, setLogoOk] = useState(true);
  // TE Logo aus public/ (CRA/StackBlitz)
  const teLogoUrl = useMemo(() => {
    const base = (process.env.PUBLIC_URL || "").replace(/\/$/, "");
    return `${base}/TE_Connectivity_logo.svg.png`;
  }, []);


  // -------- Excel Upload --------
  function handleFileUpload(e) {
    const file = e.target.files && e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const data = new Uint8Array(evt.target.result);
        const workbook = XLSX.read(data, { type: "array" });
        const rohsSheet = workbook.Sheets["RoHS"];
        const svhcSheet = workbook.Sheets["SVHC"];
        if (!rohsSheet || !svhcSheet) {
          alert("Excel muss die Sheets 'RoHS' und 'SVHC' enthalten.");
          return;
        }


        setRohsList(XLSX.utils.sheet_to_json(rohsSheet, { header: 1 }));
        setSvhcList(XLSX.utils.sheet_to_json(svhcSheet, { header: 1 }));
        setFileLoaded(true);
      } catch (err) {
        console.error(err);
        alert("Fehler beim Lesen der Excel-Datei. Bitte prüfen, ob es eine .xlsx ist.");
      }
    };


    reader.readAsArrayBuffer(file);
  }

  // -------- Trefferlogik --------
  function findRohsHit(list) {
    const q = query.trim();
    if (!q) return null;
    return list.find((row, idx) => idx !== 0 && String(row[2] || "").trim() === q);
  }


  function findSvhcHit(list) {
    const q = query.trim();
    if (!q) return null;
    return list.find(
      (row, idx) =>
        idx !== 0 &&
        (String(row[2] || "").trim() === q || String(row[3] || "").trim() === q)
    );
  }

  const rohsHit = findRohsHit(rohsList);
  const svhcHit = findSvhcHit(svhcList);

  // -------- Berechnung --------
  const pw = parseFloat(productWeight);
  const sw = parseFloat(substanceWeight);
  const percentage = pw > 0 && sw >= 0 ? (sw / pw) * 100 : null;

  function parseThreshold(cellValue) {
    // Beispiele: "> 0.1 % w/w", "0,1 %", "0.1%", "0.1"
    const raw = String(cellValue || "").toLowerCase();
    const normalized = raw.replace(/[^0-9,\.]/g, "").replace(",", ".").trim();
    return normalized === "" ? NaN : Number(normalized);
  }


  function renderThresholdResult(hit, type) {
    if (!hit || percentage === null) return null;


    const thresholdIndex = type === "RoHS" ? 3 : 4; // D bzw. E
    const threshold = parseThreshold(hit[thresholdIndex]);


    if (isNaN(threshold)) return <div>Kein Grenzwert hinterlegt</div>;


    const exceeded = percentage > threshold;
    return (
      <div style={{ marginTop: 10 }}>
        <div>
          <strong>Berechneter Anteil:</strong> {percentage.toFixed(3)} %
        </div>
        <div>
          <strong>Grenzwert:</strong> {threshold} %
        </div>
        <div
          style={{
            marginTop: 6,
            fontWeight: "bold",
            color: exceeded ? "red" : "green",
          }}
        >
          {exceeded ? "❌ Grenzwert überschritten" : "✅ Grenzwert eingehalten"}
        </div>
      </div>
    );
  }

  return (
    <div style={{ maxWidth: 760, margin: "40px auto", fontFamily: "Arial" }}>
      {/* Header */}
      <div style={{ display: "flex", alignItems: "center", gap: 16, marginBottom: 16 }}>
        {logoOk ? (
          <img
            src={teLogoUrl}
            alt="TE Connectivity"
            style={{ height: 48 }}
            onError={() => setLogoOk(false)}
          />
        ) : (
          <div style={{ width: 120, height: 48, display: "flex", alignItems: "center", color: "#888", fontSize: 12 }}>
            (TE Logo fehlt)
          </div>
        )}
        <h2 style={{ margin: 0 }}>CAS / EC‑Checker – REACH SVHC & RoHS</h2>
      </div>

      <input type="file" accept=".xlsx" onChange={handleFileUpload} />
      {fileLoaded && <p>✅ Excel geladen</p>}

      <input
        type="text"
        placeholder="CAS- oder EC-Nummer eingeben"
        value={query}
        onChange={(e) => setQuery(e.target.value)}
        disabled={!fileLoaded}
        style={{ width: "100%", marginTop: 12, padding: 8 }}
      />

      {/* Gewichte */}
      <div style={{ display: "flex", gap: 12, marginTop: 12 }}>
        <div style={{ flex: 1 }}>
          <input
            type="number"
            value={productWeight}
            onChange={(e) => setProductWeight(e.target.value)}
            style={{ width: "100%", padding: 8 }}
          />
          <div style={{ fontSize: 12, color: "#555", marginTop: 4 }}>
            Gewicht Fertigerzeugnis (g)
          </div>
        </div>
        <div style={{ flex: 1 }}>
          <input
            type="number"
            value={substanceWeight}
            onChange={(e) => setSubstanceWeight(e.target.value)}
            style={{ width: "100%", padding: 8 }}
          />
          <div style={{ fontSize: 12, color: "#555", marginTop: 4 }}>
            Gewicht Gefahrenstoff (g)
          </div>
        </div>
      </div>

      {fileLoaded && query.trim() !== "" && (
        <div style={{ marginTop: 20 }}>
          <p>
            <strong>RoHS:</strong> {rohsHit ? "JA" : "nein"}
          </p>
          {rohsHit && renderThresholdResult(rohsHit, "RoHS")}

          <p style={{ marginTop: 12 }}>
            <strong>REACH SVHC:</strong> {svhcHit ? "JA" : "nein"}
          </p>
          {svhcHit && renderThresholdResult(svhcHit, "SVHC")}

          {svhcHit && svhcHit[6] && (
            <div style={{ marginTop: 10 }}>
              <strong>Reason for inclusion:</strong>
              <div>{svhcHit[6]}</div>
            </div>
          )}
        </div>
      )}

      <p style={{ fontSize: 12, marginTop: 24, color: "#666" }}>
        Excel-Struktur:<br />
        • <b>RoHS</b>: CAS = Spalte C, Grenzwert (%) = Spalte D<br />
        • <b>SVHC</b>: CAS = Spalte C, EC = Spalte D, Grenzwert (%) = Spalte E, Reason = Spalte G
        <br />
        <br />
        Logo-Hinweis: Lege <b>TE_Connectivity_logo.svg.png</b> in den Ordner <b>public/</b>.
      </p>
    </div>
  );
}

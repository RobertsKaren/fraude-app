import { useState, useCallback, useMemo } from "react";
import Papa from "papaparse";
import * as XLSX from "xlsx";

const COMBUSTIBLES = new Set(["Diesel", "Quantium Diesel", "Super", "Quantium"]);

function parseDate(val) {
  if (!val) return null;
  const d = new Date(val);
  return isNaN(d) ? null : d;
}

function fmtDate(d) {
  if (!d) return "";
  return d.toISOString().slice(0, 10);
}

function computeSummary(rows) {
  const byDate = {};
  let totalAlertasCriticas = 0;

  // Agrupar por fecha para detectar cargas múltiples
  rows.forEach(r => {
    const d = parseDate(r.FechaSuscriptor);
    if (!d) return;
    const dateKey = fmtDate(d);
    if (!byDate[dateKey]) byDate[dateKey] = [];
    byDate[dateKey].push(r);
  });

  const isComb = r => COMBUSTIBLES.has((r.Producto || "").trim());
  const montos = rows.map(r => parseFloat(r.Monto) || 0).filter(m => m > 0);
  const medianaMonto = montos.sort((a, b) => a - b)[Math.floor(montos.length / 2)] || 0;

  const diasDetalle = [];
  let diasMultiCarga = 0;

  // Procesar cada transacción y generar alertas
  const rowsWithAlerts = rows.map(r => {
    const alerts = [];
    const fecha = fmtDate(parseDate(r.FechaSuscriptor));
    const cargasMismoDia = byDate[fecha] || [];
    
    // Regla 1: Cargas mismo día (Crítico)
    if (cargasMismoDia.length > 1 && isComb(r)) {
      alerts.push("🔴 Carga mismo día");
      totalAlertasCriticas++;
    }

    // Regla 2: Monto inusual (Advertencia)
    if (parseFloat(r.Monto) > medianaMonto * 3) {
      alerts.push("🟡 Monto alto (vs mediana)");
    }

    // Regla 3: Descuento aplicado
    const tieneDesc = parseFloat(r.Descuento) > 0;

    return { ...r, alertaStr: alerts.join(" | "), tieneDesc };
  });

  // Estadísticas para el Índice
  const totalTxn = rows.length;
  const totalConDescuento = rowsWithAlerts.filter(r => r.tieneDesc).length;
  const sitiosDistintos = new Set(rows.map(r => r.Sitio).filter(Boolean)).size;

  // Agrupar meses para promedios
  const mesesMap = {};
  rows.forEach(r => {
    const d = parseDate(r.FechaSuscriptor);
    if (!d) return;
    const mes = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
    if (!mesesMap[mes]) mesesMap[mes] = 0;
    mesesMap[mes]++;
  });
  const promTxnMes = Object.keys(mesesMap).length ? +(totalTxn / Object.keys(mesesMap).length).toFixed(1) : 0;

  // Nivel de Riesgo
  const riesgo = totalAlertasCriticas > 10 ? "CRÍTICO" : totalAlertasCriticas > 0 ? "ALTO" : "BAJO";

  return {
    totalTxn, 
    totalConDescuento, 
    totalAlertasCriticas,
    riesgo,
    promTxnMes,
    sitiosDistintos,
    rowsWithAlerts,
    fechaMin: new Date(Math.min(...rows.map(r => parseDate(r.FechaSuscriptor)).filter(Boolean))),
    fechaMax: new Date(Math.max(...rows.map(r => parseDate(r.FechaSuscriptor)).filter(Boolean))),
  };
}

function buildExcel(data, selectedDocs) {
  const wb = XLSX.utils.book_new();

  // Estilos
  const HEADER_STYLE = { font: { bold: true, color: { rgb: "FFFFFF" } }, fill: { fgColor: { rgb: "1F4E78" } }, alignment: { horizontal: "center" } };
  const CRITICO_STYLE = { font: { color: { rgb: "9C0006" }, bold: true }, fill: { fgColor: { rgb: "FFC7CE" } } };

  // 1. Hoja de RESUMEN (Índice)
  const wsIndexData = [
    ["AUDITORÍA DNI CLIENTES — RESUMEN POR DOCUMENTO"],
    ["Documento", "Nombre", "Transacciones", "Con Descuento", "Alertas Críticas", "Riesgo", "Prom. Trx/Mes", "Fecha Mín.", "Fecha Máx."]
  ];

  selectedDocs.forEach(doc => {
    const { stats: s, nombre } = data[doc];
    wsIndexData.push([
      doc, nombre, s.totalTxn, s.totalConDescuento, s.totalAlertasCriticas, s.riesgo, s.promTxnMes, fmtDate(s.fechaMin), fmtDate(s.fechaMax)
    ]);
  });

  const wsIndex = XLSX.utils.aoa_to_sheet(wsIndexData);
  XLSX.utils.book_append_sheet(wb, wsIndex, "RESUMEN");

  // 2. Hojas individuales por DNI
  selectedDocs.forEach(doc => {
    const { stats: s, nombre } = data[doc];
    const wsData = [
      [`DOCUMENTO: ${doc} | NOMBRE: ${nombre} | RIESGO: ${s.riesgo}`],
      ["Alertas Críticas:", s.totalAlertasCriticas, "Prom. Mensual:", s.promTxnMes],
      [],
      ["Fecha", "Producto", "Monto", "Descuento", "Monto Neto", "Comercio", "Alertas"]
    ];

    s.rowsWithAlerts.forEach(r => {
      wsData.push([
        r.FechaSuscriptor, r.Producto, parseFloat(r.Monto), parseFloat(r.Descuento), 
        parseFloat(r.MontoNeto), r.Comercio, r.alertaStr
      ]);
    });

    const ws = XLSX.utils.aoa_to_sheet(wsData);
    XLSX.utils.book_append_sheet(wb, ws, String(doc).slice(0, 31));
  });

  XLSX.writeFile(wb, "Auditoria_DNI_Resultados.xlsx");
}

export default function App() {
  // ... (El resto del componente App se mantiene igual que en tu código original)
  // Asegúrate de que el botón de descarga llame a buildExcel(dniData, selectedDocs)
}
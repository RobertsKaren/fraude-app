import { useState, useCallback, useMemo } from "react";
import Papa from "papaparse";
import XLSX from "xlsx-js-style";

const COMBUSTIBLES = new Set([
  "Diesel",
  "Quantium Diesel",
  "Super",
  "Quantium"
]);

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

  rows.forEach((r) => {
    const d = parseDate(r.FechaSuscriptor);
    if (!d) return;

    const dateKey = fmtDate(d);

    if (!byDate[dateKey]) byDate[dateKey] = [];

    byDate[dateKey].push(r);
  });

  const isComb = (r) =>
    COMBUSTIBLES.has((r.Producto || "").trim());

  const isStore = (r) =>
    !isComb(r) && r.Producto;

  const totalTrx = rows.length;
  const totalTienda = rows.filter(isStore).length;
  const totalComb = rows.filter(isComb).length;

  let diasDifProd = 0;
  let diasMultiComb = 0;
  let diasMismoProducto = 0;

  const diasDetalle = [];

  Object.entries(byDate).forEach(([date, dayRows]) => {
    const prods = [
      ...new Set(
        dayRows
          .map((r) => r.Producto)
          .filter(Boolean)
      )
    ];

    if (prods.length > 1) diasDifProd++;

    const combRows = dayRows.filter(isComb);

    if (combRows.length > 1) diasMultiComb++;

    const mismoProducto = {};

    combRows.forEach((r) => {
      const p = r.Producto || "Sin producto";

      if (!mismoProducto[p]) mismoProducto[p] = 0;

      mismoProducto[p]++;
    });

    if (
      Object.values(mismoProducto).some((x) => x > 1)
    ) {
      diasMismoProducto++;
    }

    const monto = dayRows.reduce(
      (s, r) => s + (parseFloat(r.Monto) || 0),
      0
    );

    const sitios = [
      ...new Set(
        dayRows
          .map((r) => r.Sitio)
          .filter(Boolean)
      )
    ]
      .sort()
      .join(", ");

    const productosStr = [
      ...new Set(
        dayRows
          .map((r) => r.Producto)
          .filter(Boolean)
      )
    ]
      .sort()
      .join(", ");

    diasDetalle.push({
      Fecha: date,
      Transacciones: dayRows.length,
      Combustibles: combRows.length,
      Tienda: dayRows.filter(isStore).length,
      Monto_Total: monto,
      Productos: productosStr,
      Sitios: sitios
    });
  });

  diasDetalle.sort((a, b) =>
    b.Fecha.localeCompare(a.Fecha)
  );

  const combDistintos = new Set(
    rows
      .filter(isComb)
      .map((r) => r.Producto)
  ).size;

  const sitiosDistintos = new Set(
    rows
      .map((r) => r.Sitio)
      .filter(Boolean)
  ).size;

  const totalConDescuento = rows.filter(
    (r) => parseFloat(r.Descuento) > 0
  ).length;

  const sumaDescuentos = rows.reduce(
  (s, r) => s + (parseFloat(r.Descuento) || 0),
  0
);

  const mesesMap = {};

  rows.forEach((r) => {
    const d = parseDate(r.FechaSuscriptor);

    if (!d) return;

    const mes = `${d.getFullYear()}-${String(
      d.getMonth() + 1
    ).padStart(2, "0")}`;

    if (!mesesMap[mes]) {
      mesesMap[mes] = {
        total: 0,
        desc: 0
      };
    }

    mesesMap[mes].total++;

    if (parseFloat(r.Descuento) > 0) {
      mesesMap[mes].desc++;
    }
  });

  const mesesArr = Object.values(mesesMap);

  const promTrxMes = mesesArr.length
    ? +(
        mesesArr.reduce((s, m) => s + m.total, 0) /
        mesesArr.length
      ).toFixed(1)
    : 0;

  const promDescMes = mesesArr.length
    ? +(
        mesesArr.reduce((s, m) => s + m.desc, 0) /
        mesesArr.length
      ).toFixed(1)
    : 0;

  const fechas = rows
    .map((r) => parseDate(r.FechaSuscriptor))
    .filter(Boolean);

  const fechaMin = fechas.length
    ? new Date(Math.min(...fechas))
    : null;

  const fechaMax = fechas.length
    ? new Date(Math.max(...fechas))
    : null;

  return {
    totalTrx,
    totalTienda,
    totalComb,
    diasDifProd,
    diasMultiComb,
    diasMismoProducto,
    combDistintos,
    sitiosDistintos,
    totalConDescuento,
    sumaDescuentos,
    promTrxMes,
    promDescMes,
    fechaMin,
    fechaMax,
    diasDetalle
  };
}

function buildExcel(data, selectedDocs) {
  const wb = XLSX.utils.book_new();

  const BORDER = {
    top: { style: "thin" },
    bottom: { style: "thin" },
    left: { style: "thin" },
    right: { style: "thin" }
  };

  const MAIN_TITLE = {
    font: {
      bold: true,
      sz: 16,
      color: { rgb: "FFFFFF" }
    },
    fill: {
      fgColor: { rgb: "16365C" }
    },
    alignment: {
      horizontal: "center",
      vertical: "center"
    },
    border: BORDER
  };

  const HEADER_STYLE = {
    font: {
      bold: true,
      color: { rgb: "FFFFFF" },
      sz: 10
    },
    fill: {
      fgColor: { rgb: "1F3864" }
    },
    alignment: {
      horizontal: "center",
      vertical: "center",
      wrapText: true
    },
    border: BORDER
  };

  const TITLE_STYLE = {
    font: {
      bold: true,
      color: { rgb: "FFFFFF" },
      sz: 10
    },
    fill: {
      fgColor: { rgb: "2E75B6" }
    },
    alignment: {
      horizontal: "center",
      vertical: "center",
      wrapText: true
    },
    border: BORDER
  };

  const BOLD_CENTER = {
    font: {
      bold: true,
      sz: 10
    },
    alignment: {
      horizontal: "center",
      vertical: "center"
    },
    border: BORDER
  };

  const NORMAL_LEFT = {
    font: {
      sz: 9
    },
    alignment: {
      horizontal: "left",
      vertical: "center",
      wrapText: true
    },
    border: BORDER
  };

  const MONEY_STYLE = {
    ...NORMAL_LEFT,
    numFmt: '"$"#,##0.00'
  };

  const MONEY_BOLD = {
  ...BOLD_CENTER,
  numFmt: '"$"#,##0.00'
};

  const ALT_FILL = {
    fgColor: { rgb: "F7F9FC" }
  };

  function makeCell(v, style) {
    return {
      v,
      s: style
    };
  }

  // ==========================
  // HOJA RESUMEN
  // ==========================

  const resumenData = [];

  resumenData.push([
    {
      v: "RESUMEN GENERAL — AUDITORÍA DNI",
      s: MAIN_TITLE
    }
  ]);

  const hdrs = [
    "Documento",
    "Nombre",
    "Total Trx",
    "Tienda",
    "Combustible",
    "Días dif. Producto",
    "Días múltiples cargas",
    "Días mismo producto",
    "Combustibles distintos",
    "Sitios distintos",
    "Con descuento",
    "Suma Descuentos",
    "Prom. Trx/Mes",
    "Prom. Desc./Mes",
    "Fecha mínima",
    "Fecha máxima"
  ];

  resumenData.push(
    hdrs.map((h) => makeCell(h, TITLE_STYLE))
  );

  selectedDocs.forEach((doc, i) => {
    const s = data[doc].stats;
    const nombre = data[doc].nombre;

    const fill =
      i % 2 === 0 ? ALT_FILL : null;

    const rowStyle = fill
      ? { ...NORMAL_LEFT, fill }
      : NORMAL_LEFT;

    resumenData.push(
      [
        doc,
        nombre,
        s.totalTrx,
        s.totalTienda,
        s.totalComb,
        s.diasDifProd,
        s.diasMultiComb,
        s.diasMismoProducto,
        s.combDistintos,
        s.sitiosDistintos,
        s.totalConDescuento,
        s.sumaDescuentos,
        s.promTrxMes,
        s.promDescMes,
        fmtDate(s.fechaMin),
        fmtDate(s.fechaMax)
      ].map((v) => makeCell(v, rowStyle))
    );
  });

  const wsResumen =
    XLSX.utils.aoa_to_sheet(resumenData);

  wsResumen["!merges"] = [
    {
      s: { r: 0, c: 0 },
      e: { r: 0, c: 14 }
    }
  ];

  wsResumen["!cols"] = [
    { wch: 14 },
    { wch: 30 },
    { wch: 12 },
    { wch: 10 },
    { wch: 12 },
    { wch: 18 },
    { wch: 20 },
    { wch: 20 },
    { wch: 20 },
    { wch: 18 },
    { wch: 16 },
    { wch: 16 },
    { wch: 16 },
    { wch: 14 },
    { wch: 14 }
  ];

  XLSX.utils.book_append_sheet(
    wb,
    wsResumen,
    "RESUMEN"
  );

  // ==========================
  // HOJAS POR DNI
  // ==========================

  selectedDocs.forEach((doc) => {
    const { stats: s, nombre } =
      data[doc];

    const wsData = [];

    const titulo = nombre
      ? `DNI: ${doc} — ${nombre}`
      : `DNI: ${doc}`;

    wsData.push([
      {
        v: titulo,
        s: MAIN_TITLE
      }
    ]);

    wsData.push([]);

    const kpiLabels1 = [
      "Total Trx",
      "Tienda",
      "Combustible",
      "Días dif. Producto",
      "Días múltiples cargas",
      "Días mismo producto",
      "Combustibles distintos",
      "Sitios distintos"
    ];

    const kpiVals1 = [
      s.totalTrx,
      s.totalTienda,
      s.totalComb,
      s.diasDifProd,
      s.diasMultiComb,
      s.diasMismoProducto,
      s.combDistintos,
      s.sitiosDistintos
    ];

    wsData.push(
      kpiLabels1.map((h) =>
        makeCell(h, TITLE_STYLE)
      )
    );

    wsData.push(
      kpiVals1.map((v) =>
        makeCell(v, BOLD_CENTER)
      )
    );

    wsData.push([]);

    const kpiLabels2 = [
      "Con descuento",
      "Suma descuentos",
      "Prom. Trx/Mes",
      "Prom. Desc./Mes",
      "Fecha mínima",
      "Fecha máxima",
    ];

    const kpiVals2 = [
      s.totalConDescuento,
      s.sumaDescuentos,
      s.promTrxMes,
      s.promDescMes,
      fmtDate(s.fechaMin),
      fmtDate(s.fechaMax)
    ];

    wsData.push(
      kpiLabels2.map((h) =>
        makeCell(h, TITLE_STYLE)
      )
    );

    wsData.push(
      kpiVals2.map((v) =>
        makeCell(v, BOLD_CENTER)
      )
    );

    wsData.push([]);

    const dayHdrs = [
      "Fecha",
      "Transacciones",
      "Combustibles",
      "Tienda",
      "Monto Total",
      "Productos",
      "Sitios"
    ];

    wsData.push(
      dayHdrs.map((h) =>
        makeCell(h, HEADER_STYLE)
      )
    );

    s.diasDetalle.forEach((row, i) => {
      const fill =
        i % 2 === 0 ? ALT_FILL : null;

      const st = fill
        ? { ...NORMAL_LEFT, fill }
        : NORMAL_LEFT;

      const moneyStyle = fill
        ? { ...MONEY_STYLE, fill }
        : MONEY_STYLE;

      wsData.push([
        makeCell(row.Fecha, st),
        makeCell(row.Transacciones, st),
        makeCell(row.Combustibles, st),
        makeCell(row.Tienda, st),
        makeCell(
          row.Monto_Total,
          moneyStyle
        ),
        makeCell(row.Productos, st),
        makeCell(row.Sitios, st)
      ]);
    });

    const lastRow =
      8 + s.diasDetalle.length+1;

wsData.push([
  makeCell("TOTAL", {
    ...BOLD_CENTER,
    fill: ALT_FILL
  }),
  {
    t: "n",
    f: `SUM(B9:B${lastRow})`,
    s: BOLD_CENTER
  },
  {
    t: "n",
    f: `SUM(C9:C${lastRow})`,
    s: BOLD_CENTER
  },
  {
    t: "n",
    f: `SUM(D9:D${lastRow})`,
    s: BOLD_CENTER
  },
  {
    t: "n",
    f: `SUM(E9:E${lastRow})`,
    s: MONEY_BOLD
  },
  makeCell("", NORMAL_LEFT),
  makeCell("", NORMAL_LEFT)
]);

    const ws =
      XLSX.utils.aoa_to_sheet(wsData);

    ws["!merges"] = [
      {
        s: { r: 0, c: 0 },
        e: { r: 0, c: 6 }
      }
    ];

    ws["!cols"] = [
      { wch: 14 },
      { wch: 14 },
      { wch: 14 },
      { wch: 12 },
      { wch: 18 },
      { wch: 50 },
      { wch: 40 }
    ];

    ws["!freeze"] = {
      xSplit: 0,
      ySplit: 8
    };

    ws["!autofilter"] = {
      ref: `A8:G${lastRow}`
    };

    XLSX.utils.book_append_sheet(
      wb,
      ws,
      String(doc).slice(0, 31)
    );
  });

  XLSX.writeFile(
    wb,
    "Auditoria_DNI.xlsx"
  );
}

export default function App() {
  const [csvData, setCsvData] =
    useState(null);

  const [dniData, setDniData] =
    useState({});

  const [selectedDocs, setSelectedDocs] =
    useState([]);

  const [search, setSearch] =
    useState("");

  const [loading, setLoading] =
    useState(false);

  const processFile = useCallback(
    (file) => {
      if (!file) return;

      setLoading(true);

      Papa.parse(file, {
        header: true,
        skipEmptyLines: true,

        complete: (results) => {
          const rows = results.data;

          const byDoc = {};

          rows.forEach((r) => {
            const doc =
              r.Documento?.trim();

            if (!doc) return;

            if (!byDoc[doc]) {
              byDoc[doc] = [];
            }

            byDoc[doc].push(r);
          });

          const computed = {};

          Object.entries(byDoc).forEach(
            ([doc, docRows]) => {
              const nombre =
                docRows.find((r) =>
                  r.Nombre?.trim()
                )?.Nombre || "";

              computed[doc] = {
                stats: computeSummary(
                  docRows
                ),
                nombre,
                count: docRows.length
              };
            }
          );

          setDniData(computed);
          setCsvData(rows);
          setSelectedDocs([]);
          setLoading(false);
        },

        error: () => {
          setLoading(false);
        }
      });
    },
    []
  );

  const allDocs = useMemo(
    () =>
      Object.keys(dniData).sort(),
    [dniData]
  );

  const filteredDocs = useMemo(() => {
    if (!search) return allDocs;

    const q = search.toLowerCase();

    return allDocs.filter((doc) => {
      const nombre = (
        dniData[doc]?.nombre || ""
      ).toLowerCase();

      return (
        doc.includes(q) ||
        nombre.includes(q)
      );
    });
  }, [allDocs, search, dniData]);

  const toggleDoc = (doc) => {
    setSelectedDocs((prev) =>
      prev.includes(doc)
        ? prev.filter((d) => d !== doc)
        : [...prev, doc]
    );
  };

  const generateReport = () => {
    if (!selectedDocs.length) return;

    buildExcel(
      dniData,
      selectedDocs
    );
  };

  return (
    <div
      style={{
        fontFamily: "Arial",
        padding: 24,
        maxWidth: 1000,
        margin: "0 auto"
      }}
    >
      <h1>
        Auditoría de Transacciones
      </h1>

      <input
        type="file"
        accept=".csv"
        onChange={(e) =>
          processFile(
            e.target.files[0]
          )
        }
      />

      {loading && (
        <p>Procesando archivo...</p>
      )}

      {csvData && (
        <>
          <div
            style={{
              marginTop: 20,
              marginBottom: 20
            }}
          >
            <input
              type="text"
              placeholder="Buscar DNI..."
              value={search}
              onChange={(e) =>
                setSearch(
                  e.target.value
                )
              }
              style={{
                padding: 10,
                width: "100%"
              }}
            />
          </div>
<div
  style={{
    display: "flex",
    gap: 10,
    marginBottom: 12
  }}
>
  <button
    onClick={() => {
      if (
        selectedDocs.length ===
        filteredDocs.length
      ) {
        setSelectedDocs([]);
      } else {
        setSelectedDocs(filteredDocs);
      }
    }}
    style={{
      padding: "10px 14px",
      fontSize: 14
    }}
  >
    {selectedDocs.length ===
      filteredDocs.length &&
    filteredDocs.length > 0
      ? "Deseleccionar todos"
      : "Seleccionar todos"}
  </button>

  <div
    style={{
      display: "flex",
      alignItems: "center",
      fontSize: 14,
      color: "#666"
    }}
  >
    {selectedDocs.length} seleccionados
  </div>
</div>
          <div
            style={{
              border: "1px solid #ccc",
              borderRadius: 8,
              maxHeight: 400,
              overflowY: "auto"
            }}
          >
            {filteredDocs.map((doc) => {
              const selected =
                selectedDocs.includes(
                  doc
                );

              const info =
                dniData[doc];

              return (
                <div
                  key={doc}
                  onClick={() =>
                    toggleDoc(doc)
                  }
                  style={{
                    padding: 12,
                    cursor: "pointer",
                    borderBottom:
                      "1px solid #eee",
                    background:
                      selected
                        ? "#dbeafe"
                        : "white"
                  }}
                >
                  <strong>{doc}</strong>

                  {info.nombre && (
                    <span>
                      {" "}
                      — {info.nombre}
                    </span>
                  )}

                  <div
                    style={{
                      marginTop: 4,
                      fontSize: 12,
                      color: "#666"
                    }}
                  >
                    {info.count} trx |{" "}
                    {
                      info.stats
                        .diasMultiComb
                    }{" "}
                    días múltiples
                    cargas
                  </div>
                </div>
              );
            })}
          </div>

          <button
            onClick={generateReport}
            disabled={
              !selectedDocs.length
            }
            style={{
              marginTop: 20,
              padding:
                "12px 18px",
              fontSize: 16
            }}
          >
            Generar Excel
          </button>
        </>
      )}
    </div>
  );
}
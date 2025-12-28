const xl = require("excel4node");

const COLUMN_MAP = {
  convocatoria: { header: "Convocatoria", path: "convocatoria", width: 18 },
  departamento: { header: "Departamento", path: "departamento", width: 18 },
  provincia: { header: "Provincia", path: "provincia", width: 18 },
  distrito: { header: "Distrito", path: "distrito", width: 18 },
  fecha_ingreso: { header: "Fecha Ingreso", path: "fecha_ingreso", width: 18 },
  fecha_registro: {
    header: "Fecha Registro",
    path: "fecha_registro",
    width: 18,
  },
  dia: { header: "Día", path: "dia", width: 12 },
  mes: { header: "Mes", path: "mes", width: 12 },
  anio: { header: "Año", path: "anio", width: 12 },
  fecha_caducidad: {
    header: "Fecha Caducidad",
    path: "fecha_caducidad",
    width: 18,
  },
  vigencia_dias: {
    header: "Vigencia (Días)",
    path: "vigencia_dias",
    width: 12,
  },
  estado: { header: "Estado", path: "estado", width: 18 },
  codigo_et: { header: "Código ET", path: "codigo_et", width: 18 },
  razon_social: { header: "Razón Social", path: "razon_social", width: 24 },
  personeria: { header: "Personería", path: "personeria", width: 18 },
  representante_legal: {
    header: "Representante Legal",
    path: "representante_legal",
    width: 24,
  },
  dni: { header: "Dni", path: "dni", width: 12 },
  ruc: { header: "Ruc", path: "ruc", width: 12 },
  direccion_razon_social: {
    header: "Dirección Razón Social",
    path: "direccion_razon_social",
    width: 24,
  },
  telefono: { header: "Teléfono", path: "telefono", width: 18 },
  celular: { header: "Celular", path: "celular", width: 18 },
  email: { header: "Email", path: "email", width: 24 },
  ingeniero: { header: "Ingeniero", path: "ingeniero", width: 24 },
  cip: { header: "Cip", path: "cip", width: 12 },
  arquitecto: { header: "Arquitecto", path: "arquitecto", width: 18 },
  cap: { header: "Cap", path: "cap", width: 12 },
  ingeniero2: { header: "Ingeniero2", path: "ingeniero2", width: 18 },
  cip2: { header: "Cip2", path: "cip2", width: 12 },
  arquitecto2: { header: "Arquitecto2", path: "arquitecto2", width: 18 },
  cap2: { header: "Cap2", path: "cap2", width: 12 },
  abogado: { header: "Abogado", path: "abogado", width: 18 },
  cal: { header: "Cal", path: "cal", width: 12 },
  resolucion: { header: "Resolucion", path: "resolucion", width: 24 },
  numero_total_de_bfhs_desembolsados1: {},
  numero_total_de_bfhs_desembolsados2: {},
  numero_total_de_bfhs_desembolsados3: {},
  numero_total_de_bfhs_desembolsados4: {},
  numero_total_de_bfhs_desembolsados: {
    header: "Total",
    path: "numero_total_de_bfhs_desembolsados",
    width: 12,
  },
};

const isProvided = (v) =>
  v !== undefined &&
  v !== null &&
  String(v).trim() !== "" &&
  String(v).toLowerCase() !== "all";

const forceZeroHourUTC = (iso) => {
  const d = new Date(iso);
  return new Date(
    Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate(), 0, 0, 0, 0)
  );
};

const escRegex = (s = "") => String(s).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");

const normalizeExpr = (fieldPath) => ({
  $let: {
    vars: {
      s: { $toLower: fieldPath },
      pairs: [
        ["á", "a"],
        ["é", "e"],
        ["í", "i"],
        ["ó", "o"],
        ["ú", "u"],
        ["ü", "u"],
        ["ñ", "n"],
      ],
    },
    in: {
      $reduce: {
        input: "$$pairs",
        initialValue: "$$s",
        in: {
          $replaceAll: {
            input: "$$value",
            find: { $arrayElemAt: ["$$this", 0] },
            replacement: { $arrayElemAt: ["$$this", 1] },
          },
        },
      },
    },
  },
});

const buildContainsAllTermsMatch = (fieldPath, rawQuery) => {
  const q = String(rawQuery || "")
    .trim()
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "");
  const terms = q.split(/\s+/).filter(Boolean);
  if (!terms.length) return null;
  return {
    $match: {
      $expr: {
        $and: terms.map((t) => ({
          $regexMatch: { input: normalizeExpr(fieldPath), regex: escRegex(t) },
        })),
      },
    },
  };
};

const buildDynamicProject = async (selectedKeys) => {
  const base = {
    _id: 0,
    _bhf: 0,
    _bhf4: 0,
    convocatoria: 1,
    departamento: 1,
    provincia: 1,
    distrito: 1,
    fecha_ingreso: 1,
    fecha_registro: 1,
    dia: 1,
    mes: 1,
    anio: 1,
    fecha_caducidad: 1,
    vigencia_dias: 1,
    estado: 1,
    codigo_et: 1,
    razon_social: 1,
    personeria: 1,
    representante_legal: 1,
    dni: 1,
    ruc: 1,
    direccion_razon_social: 1,
    telefono: 1,
    celular: 1,
    email: 1,
    ingeniero: 1,
    cip: 1,
    arquitecto: 1,
    cap: 1,
    ingeniero2: 1,
    cip2: 1,
    arquitecto2: 1,
    cap2: 1,
    abogado: 1,
    cal: 1,
    resolucion: 1,
    numero_total_de_bfhs_desembolsados1: 1,
    numero_total_de_bfhs_desembolsados2: 1,
    numero_total_de_bfhs_desembolsados3: 1,
    numero_total_de_bfhs_desembolsados4: 1,
    numero_total_de_bfhs_desembolsados: 1,
    bhf_desembolsados: 1,
  };

  const allowed = {};
  selectedKeys.forEach((k) => {
    if (COLUMN_MAP[k]) allowed[k] = base[k] ?? 1;
  });

  // Garantizamos fecha/ids si usuario los pidió; sino, se excluyen
  return Object.keys(allowed).length ? allowed : base; // fallback si enviaron algo raro
};

module.exports = {
  forceZeroHourUTC,
  buildContainsAllTermsMatch,
  escRegex,
  normalizeExpr,
  isProvided,
  buildDynamicProject,
  COLUMN_MAP,
};

const xl = require("excel4node");

const COLUMN_MAP = {
  id_convocatoria: {
    header: "ID Convocatoria",
    path: "id_convocatoria",
    width: 18,
  },
  items_count: { header: "Items", path: "items_count", width: 10 },
  departamento: { header: "Departamento", path: "departamento", width: 18 },
  nombre: { header: "Entidad", path: "nombre", width: 24 },
  fecha_publicacion: {
    header: "Fecha Publicación",
    path: "fecha_publicacion",
    width: 18,
  },
  nomenclatura: { header: "Nomenclatura", path: "nomenclatura", width: 18 },
  obj_contratacion: {
    header: "Obj. de Contratación",
    path: "obj_contratacion",
    width: 18,
  },
  descripcion: { header: "Descripción", path: "descripcion", width: 24 },
  presentacion_interval: {
    header: "Fecha Presentación",
    path: "presentacion_interval",
    width: 34,
  },
  buena_pro_interval: {
    header: "Fecha Buena Pro",
    path: "buena_pro_interval",
    width: 34,
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
    id_convocatoria: 1,
    parent_id: "$_id",
    fecha_publicacion_date: 1,
    fecha_publicacion: 1,
    nombre: 1,
    descripcion: 1,
    departamento: 1,
    nomenclatura: 1,
    obj_contratacion: 1,
    presentacion_interval: 1,
    buena_pro_interval: 1,
    items_count: { $size: { $ifNull: ["$items", []] } },
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

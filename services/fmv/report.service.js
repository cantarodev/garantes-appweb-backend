require("dotenv").config();
const mongoose = require("mongoose");
const ExcelJS = require("exceljs");

const User = require("../../models/user.model");
const {
  isProvided,
  buildDynamicProject,
  buildContainsAllTermsMatch,
  COLUMN_MAP,
} = require("../../utils/report-fmv.utils");

const getReportEntitiesFacets = async (userId, estado, departamento) => {
  try {
    const user = await User.findById(userId, "verified").lean();

    if (!user)
      return {
        success: false,
        data: [],
        message: `No se encontró el usuario con ID: ${userId}`,
      };

    if (user.verified === false)
      return {
        success: false,
        data: [],
        message: `El usuario con ID: ${userId} no ha sido verificado.`,
      };

    const col = mongoose.connection.collection("mivivienda");

    const estados = await col
      .aggregate(
        [
          { $match: { estado: { $type: "string", $nin: ["", "-"] } } },
          { $match: { departamento: { $type: "string", $nin: ["", "-"] } } },
          { $match: { personeria: { $type: "string", $nin: ["", "-"] } } },
          { $group: { _id: "$estado", count: { $sum: 1 } } },
          { $project: { _id: 0, value: "$_id", count: 1 } },
          { $sort: { value: 1 } },
          { $limit: 1000 },
        ],
        { allowDiskUse: true }
      )
      .toArray();

    const departamentos = await col
      .aggregate(
        [
          ...(isProvided(estado)
            ? [
                {
                  $match: {
                    estado: {
                      $type: "string",
                      $nin: ["", "-"],
                      $eq: String(estado).trim(),
                    },
                  },
                },
              ]
            : []),
          { $match: { departamento: { $type: "string", $nin: ["", "-"] } } },
          { $match: { personeria: { $type: "string", $nin: ["", "-"] } } },
          { $group: { _id: "$departamento", count: { $sum: 1 } } },
          { $project: { _id: 0, value: "$_id", count: 1 } },
          { $sort: { value: 1 } },
          { $limit: 1000 },
        ],
        { allowDiskUse: true }
      )
      .toArray();

    const personerias = await col
      .aggregate(
        [
          ...(isProvided(estado)
            ? [
                {
                  $match: {
                    estado: {
                      $type: "string",
                      $nin: ["", "-"],
                      $eq: String(estado).trim(),
                    },
                  },
                },
              ]
            : []),
          ...(isProvided(departamento)
            ? [
                {
                  $match: {
                    departamento: {
                      $type: "string",
                      $nin: ["", "-"],
                      $eq: String(departamento).trim(),
                    },
                  },
                },
              ]
            : []),
          { $match: { personeria: { $type: "string", $nin: ["", "-"] } } },
          { $group: { _id: "$personeria", count: { $sum: 1 } } },
          { $project: { _id: 0, value: "$_id", count: 1 } },
          { $sort: { value: 1 } },
          { $limit: 1000 },
        ],
        { allowDiskUse: true }
      )
      .toArray();

    return {
      success: true,
      data: { estados, departamentos, personerias },
      message: "El reporte se generó de forma exitosa.",
    };
  } catch (error) {
    console.error(
      "Error al obtener reporte de entidades. Detalles:",
      error.message
    );
    throw new Error(error.message);
  }
};

const getDeptExtremos = async (userId) => {
  try {
    const user = await User.findById(userId, "name lastname verified").lean();

    if (!user)
      return {
        success: false,
        data: [],
        message: `No se encontró el usuario con ID: ${userId}`,
      };

    if (user.verified === false)
      return {
        success: false,
        data: [],
        message: `El usuario con ID: ${userId} no ha sido verificado.`,
      };

    const pipeline = [
      { $match: { estado: { $type: "string", $nin: ["", "-"] } } },
      { $match: { departamento: { $type: "string", $nin: ["", "-"] } } },
      { $match: { personeria: { $type: "string", $nin: ["", "-"] } } },
      { $group: { _id: "$departamento", count: { $sum: 1 } } },
      {
        $facet: {
          top: [
            { $sort: { count: -1, _id: 1 } },
            { $limit: 1 },
            { $project: { _id: 0, departamento: "$_id", count: 1 } },
          ],
          bottom: [
            { $sort: { count: 1, _id: 1 } },
            { $limit: 1 },
            { $project: { _id: 0, departamento: "$_id", count: 1 } },
          ],
          all: [
            { $sort: { count: -1, _id: 1 } },
            { $project: { _id: 0, departamento: "$_id", count: 1 } },
          ],
        },
      },
      {
        $project: {
          top: 1,
          bottom: 1,
          ranking: "$all",
          totalDepartamentos: {
            $size: { $ifNull: ["$all", []] },
          },
          totalEntidades: {
            $reduce: {
              input: { $ifNull: ["$all", []] },
              initialValue: 0,
              in: { $add: ["$$value", { $ifNull: ["$$this.count", 0] }] },
            },
          },
        },
      },
    ];

    const col = mongoose.connection.collection("mivivienda");
    const [{ top, bottom, ranking, totalDepartamentos, totalEntidades }] =
      await col.aggregate(pipeline, { allowDiskUse: true }).toArray();

    return {
      success: true,
      top: top ?? [],
      bottom: bottom ?? [],
      ranking: ranking ?? [],
      totalDepartamentos: totalDepartamentos ?? 0,
      totalEntidades: totalEntidades ?? 0,
      message: "Cantidad de departamentos extraídos de forma exitosa.",
    };
  } catch (error) {
    console.error("LIST_PAGE_ERROR. Detalles:", error.message);
    throw new Error(error.message);
  }
};

const getReportEntities = async (
  userId,
  estado,
  departamento,
  personeria,
  ruc,
  page,
  limit
) => {
  try {
    const user = await User.findById(userId, "name lastname verified").lean();

    if (!user)
      return {
        success: false,
        data: [],
        message: `No se encontró el usuario con ID: ${userId}`,
      };

    if (user.verified === false)
      return {
        success: false,
        data: [],
        message: `El usuario con ID: ${userId} no ha sido verificado.`,
      };

    const pageNum = Math.max(parseInt(page, 10) || 1, 1);
    const perPage = Math.min(Math.max(parseInt(limit, 10) || 20, 1), 100);
    const skip = (pageNum - 1) * perPage;

    const filter = {};

    if (isProvided(estado)) {
      filter.estado = String(estado).trim();
    }

    if (isProvided(departamento)) {
      filter.departamento = String(departamento).trim();
    }

    if (isProvided(personeria)) {
      filter.personeria = String(personeria).trim();
    }

    const pipeline = [
      { $match: filter },
      { $match: { estado: { $type: "string", $nin: ["", "-"] } } },
      { $match: { departamento: { $type: "string", $nin: ["", "-"] } } },
      { $match: { personeria: { $type: "string", $nin: ["", "-"] } } },
      ...(isProvided(ruc) ? [buildContainsAllTermsMatch("$ruc", ruc)] : []),
      { $sort: { numero_total_de_bfhs_desembolsados: -1, _id: 1 } },
      {
        $facet: {
          data: [
            { $skip: skip },
            { $limit: perPage },
            {
              $project: {
                _id: 0,
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
                numero_total_de_bfhs_desembolsados: 1,
                bhf_desembolsados: 1,
              },
            },
          ],
          total: [{ $count: "count" }],
          top_departamento: [
            { $match: { departamento: { $type: "string", $nin: ["", "-"] } } },
            { $group: { _id: "$departamento", count: { $sum: 1 } } },
            { $sort: { count: -1, _id: 1 } },
            { $limit: 1 },
            { $project: { _id: 0, departamento: "$_id", count: 1 } },
          ],
          bottom_departamento: [
            { $match: { departamento: { $type: "string", $nin: ["", "-"] } } },
            { $group: { _id: "$departamento", count: { $sum: 1 } } },
            { $sort: { count: 1, _id: 1 } },
            { $limit: 1 },
            { $project: { _id: 0, departamento: "$_id", count: 1 } },
          ],
        },
      },
    ];

    const col = mongoose.connection.collection("mivivienda");
    const [{ data, total, top_departamento, bottom_departamento }] = await col
      .aggregate(pipeline, { allowDiskUse: true })
      .toArray();

    const totalRows = total?.[0]?.count || 0;
    const totalPages = Math.max(Math.ceil(totalRows / perPage), 1);
    const topDept = top_departamento?.[0] || null;
    const bottomDept = bottom_departamento?.[0] || null;

    return {
      success: true,
      data,
      page: pageNum,
      limit: perPage,
      total: totalRows,
      totalPages,
      hasPrev: pageNum > 1,
      hasNext: pageNum < totalPages,
      topDepartamento: topDept,
      bottomDepartamento: bottomDept,
      message: "El reporte se generó de forma exitosa.",
    };
  } catch (error) {
    console.error("LIST_PAGE_ERROR. Detalles:", error.message);
    throw new Error(error.message);
  }
};

const exportEntities = async ({
  userId,
  estado,
  departamento,
  personeria,
  ruc,
  columns,
  res,
}) => {
  // validar usuario
  const user = await User.findById(userId, "verified").lean();
  if (!user) throw new Error("Usuario no encontrado");
  if (!user.verified) throw new Error("Usuario no verificado");

  // filtro base
  const filter = {};
  if (isProvided(estado)) filter.estado = String(estado).trim();
  if (isProvided(departamento))
    filter.departamento = String(departamento).trim();
  if (isProvided(personeria)) filter.personeria = String(personeria).trim();

  const col = mongoose.connection.collection("mivivienda");

  const yearsAgg = await col
    .aggregate([
      { $unwind: "$bhf_desembolsados" },
      {
        $project: {
          anioNum: {
            $toInt: {
              $ifNull: ["$bhf_desembolsados.anio", null],
            },
          },
        },
      },
      { $match: { anioNum: { $ne: null } } },
      { $group: { _id: "$anioNum" } },
      { $sort: { _id: 1 } }, // ascendente
    ])
    .toArray();

  const years = yearsAgg.map((x) => x._id).slice(-4);

  // pipeline (mismo que el list, sin paginar)
  const pipeline = [
    { $match: filter },
    {
      $addFields: {
        // fuerza anio a número por si viene como string
        _bhf: {
          $map: {
            input: { $ifNull: ["$bhf_desembolsados", []] },
            as: "x",
            in: {
              anio: { $toInt: "$$x.anio" },
              csp: { $toDouble: { $ifNull: ["$$x.csp", 0] } },
            },
          },
        },
      },
    },
    {
      $addFields: {
        _bhf: {
          $sortArray: { input: "$_bhf", sortBy: { anio: 1 } },
        },
      },
    },
    { $addFields: { _bhf4: { $slice: ["$_bhf", -4] } } },
    {
      $addFields: {
        numero_total_de_bfhs_desembolsados1: {
          $ifNull: [{ $arrayElemAt: ["$_bhf4.csp", 0] }, ""],
        },
        numero_total_de_bfhs_desembolsados2: {
          $ifNull: [{ $arrayElemAt: ["$_bhf4.csp", 1] }, ""],
        },
        numero_total_de_bfhs_desembolsados3: {
          $ifNull: [{ $arrayElemAt: ["$_bhf4.csp", 2] }, ""],
        },
        numero_total_de_bfhs_desembolsados4: {
          $ifNull: [{ $arrayElemAt: ["$_bhf4.csp", 3] }, ""],
        },
      },
    },
    ...(isProvided(ruc) ? [buildContainsAllTermsMatch("$ruc", ruc)] : []),

    {
      $sort: {
        // fecha_ingreso_date: -1,
        numero_total_de_bfhs_desembolsados: -1,
        _id: 1,
      },
    },

    // Proyección dinámica según columnas elegidas
    { $project: await buildDynamicProject(columns) },
  ];

  // --- Excel streaming ---
  const workbook = new ExcelJS.stream.xlsx.WorkbookWriter({
    stream: res,
    useStyles: true,
  });
  const ws = workbook.addWorksheet("Entidades");

  const staticCols = columns
    .filter(
      (key) =>
        COLUMN_MAP[key] && !/numero_total_de_bfhs_desembolsados[1-4]$/.test(key)
    )
    .map((key) => ({
      header: COLUMN_MAP[key]?.header || key,
      key,
      width: COLUMN_MAP[key]?.width || 28,
    }));

  const dynamicCols = years
    .map((y, i) => ({
      header: `Total ${y}`,
      key: `numero_total_de_bfhs_desembolsados${i + 1}`,
      width: 14,
    }))
    .filter((col) => Array.isArray(columns) && columns.includes(col.key));

  // Une y limpia
  const columnsExcel = [...staticCols, ...dynamicCols].filter(Boolean);

  // Predicados
  const isYearTotalCol = (col) =>
    /^numero_total_de_bfhs_desembolsados[1-4]$/.test(col.key);
  const isGrandTotalCol = (col) =>
    col.key === "numero_total_de_bfhs_desembolsados";
  const yearIndex = (col) => {
    const m = col.key.match(/numero_total_de_bfhs_desembolsados(\d)$/);
    return m ? parseInt(m[1], 10) : 0;
  };

  // 3) Ordenar: normales -> totales por año (1..4) -> total general
  const weight = (col) =>
    isGrandTotalCol(col) ? 2 : isYearTotalCol(col) ? 1 : 0;

  const columnsOrdered = [...columnsExcel].sort((a, b) => {
    const wa = weight(a),
      wb = weight(b);
    if (wa !== wb) return wa - wb;
    if (wa === 1) return yearIndex(a) - yearIndex(b); // dentro de los totales por año
    return 0; // conserva orden original para el resto (o usa header si prefieres)
  });

  // 1) Asigna las columnas ya ordenadas
  ws.columns = columnsOrdered;

  // 2) Selecciona la fila de cabecera con los labels
  const headerRow = ws.getRow(1);

  // 3) Estilos: azul bebé de fondo, texto negro y bold
  headerRow.eachCell((cell) => {
    cell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "D6EAF8" }, // azul bebé (ajusta si quieres otro tono)
    };
    cell.font = {
      color: { argb: "000000" },
      bold: true, // ExcelJS no soporta 800, solo boolean
    };
    cell.alignment = {
      vertical: "middle",
      horizontal: "center",
      wrapText: true,
    };
    cell.border = {
      bottom: { style: "thin", color: { argb: "9E9E9E" } },
    };
  });

  ws.autoFilter = {
    from: { row: 1, column: 1 },
    to: { row: 1, column: columns.length },
  };

  // Cursor por lotes para no cargar todo en memoria
  const cursor = col.aggregate(pipeline, {
    allowDiskUse: true,
    batchSize: 500,
  });

  while (await cursor.hasNext()) {
    const doc = await cursor.next();

    // Asegura que sólo pasen las columnas seleccionadas
    const row = {};
    columns.forEach((k) => {
      // fecha como string legible (opcional); si prefieres ISO, usa doc[k]
      if (k === "fecha_ingreso_date" && doc[k]) {
        row[k] = new Date(doc[k]).toISOString(); // o formatea a tu gusto
      } else {
        row[k] = doc[k] ?? "";
      }
    });

    ws.addRow(row).commit();
  }

  await workbook.commit(); // cierra el stream y finaliza la respuesta
};

module.exports = {
  getDeptExtremos,
  getReportEntitiesFacets,
  getReportEntities,
  exportEntities,
};

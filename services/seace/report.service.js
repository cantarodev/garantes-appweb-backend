require("dotenv").config();
const mongoose = require("mongoose");
const ExcelJS = require("exceljs");

const User = require("../../models/user.model");
const {
  forceZeroHourUTC,
  isProvided,
  buildDynamicProject,
  buildContainsAllTermsMatch,
  COLUMN_MAP,
} = require("../../utils/report-seace.utils");

const getReportConvocatoriasFacets = async (userId, from, to, departamento) => {
  try {
    if (!from || !to)
      return {
        success: false,
        data: [],
        message: "from/to requeridos",
      };

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

    const fromDate = forceZeroHourUTC(from);
    const toDate = forceZeroHourUTC(to);
    if (isNaN(fromDate.getTime()) || isNaN(toDate.getTime())) {
      return {
        success: false,
        data: null,
        message: "Fechas inválidas.",
      };
    }

    const filter = {
      fecha_publicacion_date: { $gte: fromDate, $lt: toDate },
    };

    const col = mongoose.connection.collection("procedimientos");

    const departamentos = await col
      .aggregate(
        [
          { $match: filter },
          { $match: { departamento: { $type: "string", $nin: ["", "-"] } } },
          { $group: { _id: "$departamento", count: { $sum: 1 } } },
          { $project: { _id: 0, value: "$_id", count: 1 } },
          { $sort: { value: 1 } },
          { $limit: 1000 },
        ],
        { allowDiskUse: true }
      )
      .toArray();

    const dep = isProvided(departamento) ? String(departamento).trim() : null;
    const objetos = await col
      .aggregate(
        [
          { $match: filter },
          ...(dep
            ? [
                {
                  $match: {
                    $and: [
                      { departamento: dep },
                      { departamento: { $type: "string" } },
                      { departamento: { $nin: ["", "-"] } },
                    ],
                  },
                },
              ]
            : [
                {
                  $match: {
                    departamento: { $type: "string", $nin: ["", "-"] },
                  },
                },
              ]),
          {
            $match: { obj_contratacion: { $type: "string", $nin: ["", "-"] } },
          },
          { $group: { _id: "$obj_contratacion", count: { $sum: 1 } } },
          { $project: { _id: 0, value: "$_id", count: 1 } },
          { $sort: { value: 1 } },
          { $limit: 1000 },
        ],
        { allowDiskUse: true }
      )
      .toArray();

    return {
      success: true,
      data: { departamentos, objetos },
      message: "El reporte se generó de forma exitosa.",
    };
  } catch (error) {
    console.error(
      "Error al obtener reporte de convocatorias. Detalles:",
      error.message
    );
    throw new Error(error.message);
  }
};

const getDeptExtremos = async (userId, from, to) => {
  try {
    if (!from || !to)
      return {
        success: false,
        data: [],
        message: "from/to requeridos",
      };

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

    const fromDate = forceZeroHourUTC(from);
    const toDate = forceZeroHourUTC(to);
    if (isNaN(fromDate.getTime()) || isNaN(toDate.getTime())) {
      return {
        success: false,
        data: null,
        message: "Fechas inválidas.",
      };
    }

    const filter = {
      fecha_publicacion_date: { $gte: fromDate, $lt: toDate },
    };

    const pipeline = [
      { $match: filter },
      { $match: { departamento: { $type: "string", $nin: ["", "-"] } } },
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
          totalConvocatorias: {
            $reduce: {
              input: { $ifNull: ["$all", []] },
              initialValue: 0,
              in: { $add: ["$$value", { $ifNull: ["$$this.count", 0] }] },
            },
          },
        },
      },
    ];

    const col = mongoose.connection.collection("procedimientos");
    const [{ top, bottom, ranking, totalDepartamentos, totalConvocatorias }] =
      await col.aggregate(pipeline, { allowDiskUse: true }).toArray();

    return {
      success: true,
      top: top ?? [],
      bottom: bottom ?? [],
      ranking: ranking ?? [],
      totalDepartamentos: totalDepartamentos ?? 0,
      totalConvocatorias: totalConvocatorias ?? 0,
      message: "Cantidad de departamentos extraídos de forma exitosa.",
    };
  } catch (error) {
    console.error("LIST_PAGE_ERROR. Detalles:", error.message);
    throw new Error(error.message);
  }
};

const getReportConvocatorias = async (
  userId,
  from,
  to,
  departamento,
  objContratacion,
  q,
  page,
  limit
) => {
  try {
    if (!from || !to)
      return {
        success: false,
        data: [],
        message: "from/to requeridos",
      };

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

    const fromDate = forceZeroHourUTC(from);
    const toDate = forceZeroHourUTC(to);
    if (isNaN(fromDate.getTime()) || isNaN(toDate.getTime())) {
      return {
        success: false,
        data: null,
        message: "Fechas inválidas.",
      };
    }

    const pageNum = Math.max(parseInt(page, 10) || 1, 1);
    const perPage = Math.min(Math.max(parseInt(limit, 10) || 20, 1), 100);
    const skip = (pageNum - 1) * perPage;

    const match0 = {
      fecha_publicacion_date: { $gte: fromDate, $lt: toDate },
    };

    const addAnd = (cond) => {
      if (!match0.$and) match0.$and = [];
      match0.$and.push(cond);
    };

    if (isProvided(departamento)) {
      const dep = String(departamento).trim();
      addAnd({ departamento: dep });
      addAnd({ departamento: { $type: "string" } });
      addAnd({ departamento: { $nin: ["", "-"] } });
    } else {
      addAnd({ departamento: { $type: "string", $nin: ["", "-"] } });
    }

    if (isProvided(objContratacion)) {
      const obj = String(objContratacion).trim();
      addAnd({ obj_contratacion: obj });
      addAnd({ obj_contratacion: { $type: "string" } });
      addAnd({ obj_contratacion: { $nin: ["", "-"] } });
    }

    if (isProvided(q)) {
      const qEscaped = String(q).trim().replace(/"/g, '\\"');
      match0.$text = {
        $search: `"${qEscaped}"`,
        $caseSensitive: false,
        $diacriticSensitive: false,
      };
    }

    const pipeline = [
      { $match: match0 },
      {
        $facet: {
          data: [
            {
              $addFields: {
                cronograma_norm: {
                  $map: {
                    input: { $ifNull: ["$cronograma", []] },
                    as: "c",
                    in: {
                      etapaNorm: {
                        $let: {
                          vars: {
                            s: { $toLower: "$$c.etapa" },
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
                      },
                      fecha_inicio: "$$c.fecha_inicio",
                      fecha_fin: "$$c.fecha_fin",
                    },
                  },
                },
              },
            },
            {
              $addFields: {
                etapa_presentacion: {
                  $first: {
                    $filter: {
                      input: "$cronograma_norm",
                      as: "c",
                      cond: {
                        $regexMatch: {
                          input: "$$c.etapaNorm",
                          regex: "presentacion\\s*de\\s*(ofertas|propuestas)",
                        },
                      },
                    },
                  },
                },
                etapa_buena_pro: {
                  $first: {
                    $filter: {
                      input: "$cronograma_norm",
                      as: "c",
                      cond: {
                        $regexMatch: {
                          input: "$$c.etapaNorm",
                          regex: "buena\\s*pro",
                        },
                      },
                    },
                  },
                },
              },
            },
            {
              $addFields: {
                presentacion_interval: {
                  $cond: [
                    {
                      $and: [
                        {
                          $ifNull: ["$etapa_presentacion.fecha_inicio", false],
                        },
                        { $ifNull: ["$etapa_presentacion.fecha_fin", false] },
                      ],
                    },
                    {
                      $concat: [
                        "$etapa_presentacion.fecha_inicio",
                        " | ",
                        "$etapa_presentacion.fecha_fin",
                      ],
                    },
                    "",
                  ],
                },
                buena_pro_interval: {
                  $cond: [
                    {
                      $and: [
                        { $ifNull: ["$etapa_buena_pro.fecha_inicio", false] },
                        { $ifNull: ["$etapa_buena_pro.fecha_fin", false] },
                      ],
                    },
                    {
                      $concat: [
                        "$etapa_buena_pro.fecha_inicio",
                        " | ",
                        "$etapa_buena_pro.fecha_fin",
                      ],
                    },
                    "",
                  ],
                },
              },
            },

            {
              $addFields: {
                items_count: { $size: { $ifNull: ["$items", []] } },
              },
            },

            {
              $project: {
                items: 0,
                cronograma: 0,
                cronograma_norm: 0,
              },
            },

            { $sort: { fecha_publicacion_date: -1, _id: 1 } },
            { $skip: skip },
            { $limit: perPage },

            {
              $project: {
                _id: 0,
                id_convocatoria: 1,
                parent_id: "$_id",
                fecha_publicacion_date: 1,
                nombre: 1,
                descripcion: 1,
                departamento: 1,
                nomenclatura: 1,
                obj_contratacion: 1,
                presentacion_interval: 1,
                buena_pro_interval: 1,
                items_count: 1,
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

    const col = mongoose.connection.collection("procedimientos");
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

const exportConvocatorias = async ({
  userId,
  from,
  to,
  departamento,
  objContratacion,
  nomenclatura,
  descripcion,
  columns,
  res,
}) => {
  // validar usuario
  const user = await User.findById(userId, "verified").lean();
  if (!user) throw new Error("Usuario no encontrado");
  if (!user.verified) throw new Error("Usuario no verificado");

  const fromDate = forceZeroHourUTC(from);
  const toDate = forceZeroHourUTC(to);
  if (isNaN(fromDate.getTime()) || isNaN(toDate.getTime())) {
    throw new Error("Fechas inválidas");
  }
  // filtro base
  const filter = { fecha_publicacion_date: { $gte: fromDate, $lt: toDate } };
  if (isProvided(departamento))
    filter.departamento = String(departamento).trim();
  if (isProvided(objContratacion))
    filter.obj_contratacion = String(objContratacion).trim();

  // pipeline (mismo que el list, sin paginar)
  const pipeline = [
    { $match: filter },
    {
      $addFields: {
        cronograma_norm: {
          $map: {
            input: { $ifNull: ["$cronograma", []] },
            as: "c",
            in: {
              etapaNorm: {
                $let: {
                  vars: {
                    s: { $toLower: "$$c.etapa" },
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
              },
              fecha_inicio: "$$c.fecha_inicio",
              fecha_fin: "$$c.fecha_fin",
            },
          },
        },
      },
    },
    {
      $addFields: {
        etapa_presentacion: {
          $first: {
            $filter: {
              input: "$cronograma_norm",
              as: "c",
              cond: {
                $regexMatch: {
                  input: "$$c.etapaNorm",
                  regex: "presentacion\\s*de\\s*(ofertas|propuestas)",
                },
              },
            },
          },
        },
        etapa_buena_pro: {
          $first: {
            $filter: {
              input: "$cronograma_norm",
              as: "c",
              cond: {
                $regexMatch: { input: "$$c.etapaNorm", regex: "buena\\s*pro" },
              },
            },
          },
        },
      },
    },
    {
      $addFields: {
        presentacion_interval: {
          $cond: [
            {
              $and: [
                { $ifNull: ["$etapa_presentacion.fecha_inicio", false] },
                { $ifNull: ["$etapa_presentacion.fecha_fin", false] },
              ],
            },
            {
              $concat: [
                "$etapa_presentacion.fecha_inicio",
                " | ",
                "$etapa_presentacion.fecha_fin",
              ],
            },
            "",
          ],
        },
        buena_pro_interval: {
          $cond: [
            {
              $and: [
                { $ifNull: ["$etapa_buena_pro.fecha_inicio", false] },
                { $ifNull: ["$etapa_buena_pro.fecha_fin", false] },
              ],
            },
            {
              $concat: [
                "$etapa_buena_pro.fecha_inicio",
                " | ",
                "$etapa_buena_pro.fecha_fin",
              ],
            },
            "",
          ],
        },
      },
    },
    ...(isProvided(nomenclatura)
      ? [buildContainsAllTermsMatch("$nomenclatura", nomenclatura)]
      : []),
    ...(isProvided(descripcion)
      ? [buildContainsAllTermsMatch("$descripcion", descripcion)]
      : []),

    { $sort: { fecha_publicacion_date: -1, _id: 1 } },

    // Proyección dinámica según columnas elegidas
    { $project: await buildDynamicProject(columns) },
  ];

  const col = mongoose.connection.collection("procedimientos");

  // --- Excel streaming ---
  const workbook = new ExcelJS.stream.xlsx.WorkbookWriter({
    stream: res,
    useStyles: true,
  });
  const ws = workbook.addWorksheet("Convocatorias");

  // Definir columnas en Excel (sin header, solo key y width)
  ws.columns = columns.map((key) => ({
    key,
    width: COLUMN_MAP[key]?.width || 24,
  }));

  // Habilitar autofiltro (antes de escribir filas)
  ws.autoFilter = {
    from: { row: 1, column: 1 },
    to: { row: 1, column: columns.length },
  };

  // 3) Escribir cabecera manual con estilos
  const headerLabels = columns.map((key) => COLUMN_MAP[key]?.header || key);
  const headerRow = ws.addRow(headerLabels);

  // Estilos: azul bebé de fondo, texto negro y bold
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
  headerRow.commit?.();

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
      if (k === "fecha_publicacion_date" && doc[k]) {
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
  getReportConvocatoriasFacets,
  getReportConvocatorias,
  exportConvocatorias,
};

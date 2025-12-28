const reportService = require("../../services/fmv/report.service");

const getReportEntitiesFacets = async (req, res) => {
  const { userId, estado, departamento } = req.query;
  try {
    const results = await reportService.getReportEntitiesFacets(
      userId,
      estado,
      departamento
    );

    res.status(200).json(results);
  } catch (error) {
    console.error(error);
    return res.status(500).json({
      success: false,
      data: null,
      message: error.message,
    });
  }
};

const getDeptExtremos = async (req, res) => {
  const { userId } = req.query;

  try {
    const results = await reportService.getDeptExtremos(userId);

    res.status(200).json(results);
  } catch (error) {
    console.error(error);
    return res.status(500).json({
      success: false,
      data: null,
      message: error.message,
    });
  }
};

const getReportEntities = async (req, res) => {
  const {
    userId,
    estado,
    departamento,
    personeria,
    ruc,
    page = 1,
    limit = 20,
  } = req.query;

  try {
    const results = await reportService.getReportEntities(
      userId,
      estado,
      departamento,
      personeria,
      ruc,
      page,
      limit
    );

    res.status(200).json(results);
  } catch (error) {
    console.error(error);
    return res.status(500).json({
      success: false,
      data: null,
      message: error.message,
    });
  }
};

const exportEntitiesExcel = async (req, res) => {
  try {
    const { userId, estado, departamento, personeria, ruc, columns } =
      req.body || {};

    if (!userId) {
      return res
        .status(400)
        .json({ success: false, message: "userId es requerido" });
    }
    if (!Array.isArray(columns) || columns.length === 0) {
      return res
        .status(400)
        .json({ success: false, message: "columns (array) es requerido" });
    }

    const filename = `entidades_${new Date().toISOString().slice(0, 19).replace(/:/g, '-').replace('T', '_')}.xlsx`;


    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );
    res.setHeader("Content-Disposition", `attachment; filename="${filename}"`);
    res.setHeader("Cache-Control", "no-store");

    res.setHeader("Access-Control-Expose-Headers", "Content-Disposition");

    // genera y streamea el Excel directo a la respuesta
    await reportService.exportEntities({
      userId,
      estado,
      departamento,
      personeria,
      ruc,
      columns,
      res,
    });
  } catch (err) {
    console.error("EXPORT_ERROR:", err);
    if (!res.headersSent) {
      res
        .status(500)
        .json({ success: false, message: "Error al generar Excel" });
    } else {
      // si ya enviamos headers, cerramos la conexión
      try {
        res.end();
      } catch (_) {}
    }
  }
};

module.exports = {
  getDeptExtremos,
  getReportEntitiesFacets,
  getReportEntities,
  exportEntitiesExcel,
};

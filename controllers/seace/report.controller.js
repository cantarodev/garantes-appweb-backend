const reportService = require("../../services/seace/report.service");

const getReportConvocatoriasFacets = async (req, res) => {
  const { userId, fromDate, toDate, departamento } = req.query;
  try {
    const results = await reportService.getReportConvocatoriasFacets(
      userId,
      fromDate,
      toDate,
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
  const { userId, fromDate, toDate } = req.query;

  try {
    const results = await reportService.getDeptExtremos(
      userId,
      fromDate,
      toDate
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

const getReportConvocatorias = async (req, res) => {
  const {
    userId,
    fromDate,
    toDate,
    departamento,
    objContratacion,
    q,
    page = 1,
    limit = 20,
  } = req.query;

  try {
    const results = await reportService.getReportConvocatorias(
      userId,
      fromDate,
      toDate,
      departamento,
      objContratacion,
      q,
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

const exportConvocatoriasExcel = async (req, res) => {
  try {
    const {
      userId,
      fromDate,
      toDate,
      departamento,
      objContratacion,
      q,
      columns,
    } = req.body || {};

    if (!userId || !fromDate || !toDate) {
      return res
        .status(400)
        .json({ success: false, message: "userId, from y to son requeridos" });
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
    await reportService.exportConvocatorias({
      userId,
      from: fromDate,
      to: toDate,
      departamento,
      objContratacion,
      q,
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
  getReportConvocatoriasFacets,
  getReportConvocatorias,
  exportConvocatoriasExcel,
};

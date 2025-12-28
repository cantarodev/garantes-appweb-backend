const express = require("express");
const {
  getDeptExtremos,
  getReportConvocatorias,
  getReportConvocatoriasFacets,
  exportConvocatoriasExcel,
} = require("../../controllers/seace/report.controller");

const route = express.Router();

route.get("/convocatorias/departments", getDeptExtremos);

route.get("/convocatorias/show", getReportConvocatorias);

route.get("/convocatorias/facets", getReportConvocatoriasFacets);

route.post("/convocatorias/export", exportConvocatoriasExcel);

module.exports = route;

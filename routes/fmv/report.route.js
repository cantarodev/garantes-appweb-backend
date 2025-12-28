const express = require("express");
const {
  getDeptExtremos,
  getReportEntities,
  getReportEntitiesFacets,
  exportEntitiesExcel,
} = require("../../controllers/fmv/report.controller");

const route = express.Router();

route.get("/entities/departments", getDeptExtremos);

route.get("/entities/show", getReportEntities);

route.get("/entities/facets", getReportEntitiesFacets);

route.post("/entities/export", exportEntitiesExcel);

module.exports = route;

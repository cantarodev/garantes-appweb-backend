const express = require("express");
const routeUser = require("./user.route");
const reportSeaceManager = require("./seace/report.route");
const reportFmvManager = require("./fmv/report.route");
const auth = require("./auth.route");

const app = express();

app.use("/auth", auth);
app.use("/user", routeUser);
app.use("/seace/report", reportSeaceManager);
app.use("/fmv/report", reportFmvManager);

module.exports = app;

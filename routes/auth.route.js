const express = require("express");
const { refreshToken } = require("../controllers/auth.controller");

const route = express.Router();

route.post("/refresh-token", refreshToken);

module.exports = route;

const rateLimit = require("express-rate-limit");

const apiLimiter = rateLimit({
  windowMs: 15 * 60 * 1000,
  max: 15,
  message: {
    status: 429,
    message:
      "Demasiadas peticiones desde esta IP, por favor inténtalo de nuevo después de 15 minutos.",
  },
  headers: true,
});

module.exports = apiLimiter;

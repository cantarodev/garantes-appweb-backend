const jwt = require("jsonwebtoken");

const authenticateToken = (req, res, next) => {
  const authHeader = req.headers["authorization"];
  const token = authHeader && authHeader.split(" ")[1];

  if (!token) {
    return res
      .status(401)
      .json({ message: "Token no proporcionado", data: {} });
  }

  jwt.verify(token, process.env.ACCESS_TOKEN_SECRET, (err, user) => {
    if (err) {
      return res.status(403).json({ message: "Token inválido", data: {} });
    }

    user;
    req.user = user;
    next();
  });
};

const roleMiddleware = (requiredRole) => {
  return (req, res, next) => {
    console.log("role: ", req.user);

    if (req.user.role !== requiredRole) {
      return res.status(403).json({ message: "Acceso denegado" });
    }
    next();
  };
};

module.exports = { authenticateToken, roleMiddleware };

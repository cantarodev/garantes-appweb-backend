require("dotenv").config();
const express = require("express");
const http = require("http");
const cors = require("cors");
const path = require("path");
const helmet = require("helmet");
const connectToMongoDB = require("./database");
const indexRoutes = require("./routes");
require("./models");

const socketIo = require("socket.io");

const app = express();

const corsOptions = {
  origin: process.env.FRONT_URL,
  methods: ["GET", "POST", "PUT", "DELETE"],
  allowedHeaders: ["Content-Type", "Authorization"],
  exposedHeaders: ["Content-Disposition"],
  credentials: true,
};

app.use(helmet());
app.use(cors(corsOptions));
app.use(express.json({ limit: "1024mb" }));
app.use(express.urlencoded({ extended: true, limit: "1024mb" }));
app.use("/api/v1", indexRoutes);

const PORT = process.env.PORT;
const server = http.createServer(app);
const io = socketIo(server, {
  cors: {
    origin: process.env.FRONT_URL, // Permitir el origen de tu front-end
    methods: ["GET", "POST"], // Métodos permitidos
    credentials: true, // Permitir envío de cookies si es necesario
  },
});

const startServer = async () => {
  try {
    await connectToMongoDB();

    server.listen(PORT, () => {
      console.log(`Servidor corriendo en el puerto ${PORT}`);
    });
  } catch (err) {
    console.error("Error al iniciar el servidor:", err);
  }
};

startServer();

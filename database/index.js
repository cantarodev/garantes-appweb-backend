const mongoose = require("mongoose");

const uri = process.env.MONGODB_URI;

const connectToMongoDB = async () => {
  try {
    await mongoose.connect(uri, {
      serverSelectionTimeoutMS: 5000,
      socketTimeoutMS: 45000,
      connectTimeoutMS: 10000,
      maxPoolSize: 10,
    });
    console.log("Conexión a MongoDB exitosa");
  } catch (err) {
    console.error("Error al conectar a MongoDB:", err);
    setTimeout(connectToMongoDB, 5000);
  }
};

mongoose.connection.on("disconnected", () => {
  console.log("Conexión a MongoDB perdida. Intentando reconectar...");
  setTimeout(connectToMongoDB, 5000);
});

process.on("SIGINT", () => {
  mongoose.connection.close(() => {
    console.log(
      "Conexión a MongoDB cerrada debido a la terminación de la aplicación"
    );
    process.exit(0);
  });
});

module.exports = connectToMongoDB;

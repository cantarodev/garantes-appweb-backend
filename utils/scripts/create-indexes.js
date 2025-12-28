const mongoose = require("mongoose");
require("dotenv").config();

(async () => {
  if (!process.env.MONGODB_URI) {
    throw new Error(
      "MONGODB_URI no está definido en las variables de entorno."
    );
  }

  await mongoose.connect(process.env.MONGODB_URI);

  const db = mongoose.connection.db;
  const col = db.collection("procedimientos");

  await col.createIndex({ fecha_publicacion_date: -1, _id: 1 });

  await col.createIndex({
    fecha_publicacion_date: -1,
    departamento: 1,
  });
  await col.createIndex({
    fecha_publicacion_date: -1,
    obj_contratacion: 1,
  });

  // await col.createIndex({ nomenclatura: 1 });
  // await col.createIndex({ obj_contratacion: 1 });
  // await col.createIndex({
  //   fecha_publicacion_date: 1,
  //   nomenclatura: 1,
  //   obj_contratacion: 1,
  // });
  // await col.createIndex({ nomenclatura_norm: 1 });
  // await col.createIndex({ descripcion_norm: 1 });
  // await col.createIndex({
  //   nombre: "text",
  //   nomenclatura: "text",
  //   descripcion: "text",
  // });
  // await col.dropIndex("descripcion_text");

  // await col.createIndex(
  //   { descripcion: "text" },
  //   { default_language: "spanish" }
  // );

  console.log("Índices creados");
  await mongoose.disconnect();
})();

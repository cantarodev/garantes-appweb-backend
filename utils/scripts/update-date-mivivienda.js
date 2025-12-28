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
  const col = db.collection("mivivienda");

  await col.updateMany(
    { fecha_ingreso: { $type: "string", $regex: /^\d{2}\/\d{2}\/\d{4}$/ } },
    [
      {
        $set: {
          _tmp_dt: {
            $dateFromString: {
              dateString: "$fecha_ingreso",
              format: "%d/%m/%Y",
              onError: null,
              onNull: null,
            },
          },
        },
      },
      {
        // solo crea/actualiza cuando la conversión fue válida
        $set: {
          fecha_ingreso_date: {
            $cond: [{ $ne: ["$_tmp_dt", null] }, "$_tmp_dt", "$$REMOVE"],
          },
        },
      },
      { $unset: "_tmp_dt" },
    ]
  );

  await col.createIndex({ fecha_ingreso_date: 1 });
  await col.createIndex({ estado: 1 });
  await col.createIndex({ departamento: 1 });
  await col.createIndex({ personeria: 1 });
  // await col.createIndex({
  //   fecha_publicacion_date: 1,
  //   nomenclatura: 1,
  //   obj_contratacion: 1,
  // });

  console.log("Índices creados");
  await mongoose.disconnect();
})();

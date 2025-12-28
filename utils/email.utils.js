const nodemailer = require("nodemailer");
const bcrypt = require("bcrypt");
const { v4: uuidv4 } = require("uuid");
const User = require("../models/user.model");

const transporter = nodemailer.createTransport({
  host: process.env.SMTP_HOST,
  port: Number(process.env.SMTP_PORT),
  secure: true,
  auth: { user: process.env.SMTP_USER, pass: process.env.SMTP_PASS },
  pool: true,
  maxConnections: 5,
  maxMessages: 100,
});

async function sendEmail({ to, from, subject, html }) {
  console.log("EMAIL: ", { to, from, subject, html });

  const info = await transporter.sendMail({
    from: from || process.env.SEND_AUTH_EMAIL,
    to,
    subject,
    html,
  });
  if (!info?.messageId) throw new Error("SMTP_SEND_ERROR");
  return info;
}

const sendNotificationPleProcessed = async ({
  user_id,
  email,
  name,
  type,
  period,
}) => {
  const convertedType = type === "compras" ? "purchases" : "sales";
  const url = `${process.env.FRONT_URL}/dashboard/analytics/${convertedType}?user_id=${user_id}`;

  const mailOptions = {
    to: email,
    from: process.env.SEND_AUTH_EMAIL,
    subject: "¡Tu información ha sido procesada en Izitax!",
    html: `
    <div style="font-family: Arial, sans-serif; color: #333; line-height: 1.6;">
      <h2 style="color: #4CAF50;">¡Tu información ha sido procesada en Izitax!</h2>
      
      <p>Hola,</p>
      
      <p>Nos complace informarte que tu PLE <strong>${type}</strong> <strong>${name}</strong> del período <strong>${period}</strong> ha sido <strong>procesada correctamente</strong>. Ahora puedes acceder a <a href=${url} style="color: #4CAF50; text-decoration: none;">Izitax</a> y revisar los resultados.</p>
      
      <br />
      <p>Saludos,<br>El equipo de <strong>Izitax</strong></p>
      
      <hr style="border: 1px solid #ddd;" />
      <p style="font-size: 12px; color: #777;">Este es un mensaje automático. Por favor, no respondas a este correo.</p>
    </div>`,
  };

  try {
    await sendEmail(mailOptions);
    console.log("Confirmación de PLE procesado enviado.");
  } catch (error) {
    console.error(error);
    if (error.response) {
      console.error(error.response.body);
    }
  }
};

const sendNotificationPleError = async ({
  user_id,
  emails,
  name,
  type,
  period,
  errors,
}) => {
  // const url = `${process.env.FRONT_URL}/dashboard/file-manager?user_id=${user_id}`;
  const mailOptions = {
    to: emails,
    from: process.env.SEND_AUTH_EMAIL,
    subject: "¡Error inesperado en Izitax!",
    html: `
    <div style="font-family: Arial, sans-serif; color: #333; line-height: 1.6;">
      <h2 style="color: #4CAF50;">¡Un error inesperado al procesar el PLE!</h2>
      
      <p>Hola,</p>
      
      <p>Nos complace informarte que tu PLE <strong>${type}</strong> <strong>${name}</strong> del período <strong>${period}</strong> no ha sido <strong>procesada correctamente</strong>. A continuación la lista de las posibles causas:</p>
      <ul>
        ${errors.map((error) => `<li>${error}</li>`).join("")}
      </ul>
      <br />
      <p>Saludos,<br>El equipo de <strong>Izitax</strong></p>
      
      <hr style="border: 1px solid #ddd;" />
      <p style="font-size: 12px; color: #777;">Este es un mensaje automático. Por favor, no respondas a este correo.</p>
    </div>`,
  };

  try {
    await sendEmail(mailOptions);
    console.log("Confirmación de PLE procesado enviado.");
  } catch (error) {
    console.error(error);
    if (error.response) {
      console.error(error.response.body);
    }
  }
};

const sendVerificationEmail = async ({ user_id, email }, res) => {
  const currentUrl = process.env.FRONT_URL;
  const uniqueString = uuidv4() + user_id;
  const saltRounds = 10;

  const mailOptions = {
    to: email,
    from: process.env.SEND_AUTH_EMAIL,
    subject: "Verifica tu correo",
    html: `<p>
  Verifique su dirección de correo electrónico para completar el registro e iniciar sesión en su cuenta.</p><p>Este enlace <b>expira en 1 hora</b>.</p><p>Presiona <a href=${
    currentUrl + "/auth/verify/" + user_id + "/" + uniqueString
  }>aquí</a> para proceder.</p>`,
  };

  try {
    const hashedUniqueString = await bcrypt.hash(uniqueString, saltRounds);

    await User.updateOne(
      { _id: user_id },
      {
        $set: {
          verification_email: {
            unique_string: hashedUniqueString,
            expire_at: Date.now() + 3600000,
          },
        },
      }
    );

    await sendEmail(mailOptions);

    res.status(200).json({
      status: "success",
      data: { user_id, email },
      message: "¡El mensaje de verificación ha sido enviado!",
    });
  } catch (error) {
    console.error(error);
    let message;
    if (error.message.includes("hash")) {
      message =
        "¡Ocurrió un error al aplicar hash a los datos del correo electrónico!";
    } else if (error.message.includes("update")) {
      message =
        "¡No se pudieron guardar los datos del correo electrónico de verificación!";
    } else if (error.message.includes("sendMail")) {
      message = "¡El correo electrónico de verificación falló!";
    } else {
      message = "¡Ocurrió un error inesperado!";
    }
    res.status(500).json({
      status: "error",
      data: null,
      message,
    });
  }
};

const sendResetEmail = async ({ user_id, email }, redirectUrl, res) => {
  const resetString = uuidv4() + user_id;
  const saltRounds = 10;

  const mailOptions = {
    to: email,
    from: process.env.SEND_AUTH_EMAIL,
    subject: "Restablecimiento de contraseña",
    html: `<p>Hemos oído que perdiste la contraseña.</p> <p>No te preocupes, utiliza el siguiente enlace para restablecerlo.</p> <p>Este enlace <b>caduca en 60 minutos</b>.</p> <p>Presione <a href=${
      redirectUrl + "/" + user_id + "/" + resetString
    }>aquí</a> para continuar.</p>`,
  };

  try {
    const hashedResetString = await bcrypt.hash(resetString, saltRounds);

    await User.updateOne(
      { _id: user_id },
      {
        $set: {
          password_reset: {
            reset_string: hashedResetString,
            expire_at: Date.now() + 3600000, // 1 hora en milisegundos
          },
        },
      }
    );

    await sendEmail(mailOptions);

    res.status(200).json({
      status: "success",
      message: "Correo electrónico de restablecimiento de contraseña enviado.",
    });
  } catch (error) {
    console.error(error);

    let message;
    if (error.message.includes("hash")) {
      message =
        "¡Se produjo un error al procesar los datos de restablecimiento de contraseña!";
    } else if (error.message.includes("update")) {
      message =
        "¡No se pudieron guardar los datos para restablecer la contraseña!";
    } else if (error.message.includes("sendMail")) {
      message =
        "Error en el correo electrónico para restablecer la contraseña.";
    } else {
      message = "¡Ocurrió un error inesperado!";
    }

    res.status(500).json({
      status: "error",
      message,
    });
  }
};

module.exports = {
  sendVerificationEmail,
  sendResetEmail,
  sendNotificationPleProcessed,
  sendNotificationPleError,
};

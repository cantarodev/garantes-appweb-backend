const yup = require("yup");

const validatePassword = async (req, res, next) => {
  const schema = yup.object().shape({
    password: yup
      .string()
      .required("La contraseña es requerida")
      .min(8, "La contraseña debe tener al menos 8 caracteres")
      .test(
        "uppercase",
        "La contraseña debe contener al menos una letra mayúscula",
        (value) => /[A-Z]/.test(value)
      )
      .test(
        "number",
        "La contraseña debe contener al menos un número",
        (value) => /\d/.test(value)
      )
      .test(
        "specialChar",
        "La contraseña debe contener al menos un carácter especial",
        (value) => /[!@#$%^&*()_+]/.test(value)
      ),
  });

  try {
    await schema.validate(req.body, { abortEarly: false });
    next();
  } catch (error) {
    console.log("Validation Error:", error.errors);
    res
      .status(400)
      .json({ error: "Error de validación", details: error.errors });
  }
};

module.exports = {
  validatePassword,
};

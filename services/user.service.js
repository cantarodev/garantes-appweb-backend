const User = require("../models/user.model");
const jwt = require("jsonwebtoken");
const bcrypt = require("bcrypt");
const { v4: uuidv4 } = require("uuid");
const { generateAccessToken } = require("../utils/token.utils");

const getUsers = async () => {
  try {
    return await User.find();
  } catch (error) {
    throw new Error("Error al obtener usuarios. Detalles: " + error.message);
  }
};

const getUser = async (_id) => {
  try {
    const user = await User.findById(_id).select(
      "avatar email name lastname role_id"
    );

    if (!user) {
      throw new Error("Usuario no encontrado.");
    }

    return {
      user_id: user._id,
      ...user._doc,
    };
  } catch (error) {
    throw new Error("Error al obtener usuarios. Detalles: " + error.message);
  }
};

const infoUser = async (email, password) => {
  try {
    const user = await User.findOne({ email });

    if (!user) {
      throw new Error("Usuario no encontrado.");
    }

    if (!user.verified) {
      throw new Error(
        "El correo electrónico aún no ha sido verificado. Revisa tu correo."
      );
    }

    if (user.status == "pending") {
      throw new Error(
        "Acceso denegado: Tu solicitud está pendiente de aprobación por el administrador."
      );
    }

    if (user.status == "inactive") {
      throw new Error(
        "Acceso denegado: Tu cuenta ha sido desactivada. Contacta al administrador para más información."
      );
    }
    if (user.status != "active") {
      throw new Error(
        "Acceso denegado: Estado de cuenta inválido (${user.status}). Contacta al administrador."
      );
    }

    const passwordMatch = await bcrypt.compare(password, user.password);

    if (!passwordMatch) {
      throw new Error(
        "Las credenciales proporcionadas no son válidas. Por favor, inténtalo nuevamente."
      );
    }

    const token = generateAccessToken({
      user_id: user._id,
      email: user.email,
      name: user.name,
      role: user.role,
    });

    return {
      token,
    };
  } catch (error) {
    throw new Error(error.message);
  }
};

const verifyLink = async (userId, resetString) => {
  try {
    const user = await User.findOne({ _id: userId });
    if (!user) {
      console.log("Usuario no encontrado.");
      return false;
    }

    if (!user.password_reset || Object.keys(user.password_reset).length === 0) {
      console.log("No se encontraron datos para restablecer la contraseña.");
      return false;
    }

    if (user.password_reset.expire_at < Date.now()) {
      User.updateOne(
        { _id: userId },
        {
          $set: {
            password_reset: {},
          },
        }
      );
      console.log("El enlace de reinicio ha expirado.");
      return false;
    }

    const isResetStringValid = await bcrypt.compare(
      resetString,
      user.password_reset.reset_string
    );

    if (!isResetStringValid) {
      console.log("La cadena de reinicio no coincide.");
    }

    return isResetStringValid;
  } catch (error) {
    console.error(error);
    throw new Error("Error al verificar el enlace de restablecimiento.");
  }
};

const createUser = async (
  name,
  lastname,
  dni,
  phone,
  business_name,
  ruc,
  email,
  password
) => {
  try {
    const existingUser = await User.findOne({ email });
    console.log("El usuario: ", existingUser);

    if (existingUser) {
      throw new Error("El usuario ya existe.");
    }

    const business = { business_id: uuidv4(), name: business_name, ruc };

    const user = await User.create({
      name,
      lastname,
      dni,
      phone,
      business,
      accounts: [],
      verification_email: {},
      password_reset: {},
      email,
      password,
    });

    return {
      user_id: user._id,
      email: user.email,
    };
  } catch (error) {
    console.error(error);
    throw new Error(error.message);
  }
};

const updateUser = async (avatar, name, lastname, email, password) => {
  try {
    const existingUser = await User.findOne({ email });

    if (!existingUser) {
      throw new Error("No se encontró al usuario");
    }
    let updateUserObject = {
      email: email || existingUser.email,
      avatar: avatar || existingUser.avatar,
      name: name || existingUser.name,
      lastname: lastname || existingUser.lastname,
    };

    if (password !== "") {
      const salt = await bcrypt.genSalt(10);
      const hashedPassword = await bcrypt.hash(password, salt);
      updateUserObject.password = hashedPassword;
    }

    return await User.findOneAndUpdate(
      { email },
      {
        $set: updateUserObject,
      },
      {
        new: true,
        fields: "name lastname dni phone email",
      }
    );
  } catch (error) {
    throw new Error(
      "Error al actualizar la información del usuario. Detalles: " +
        error.message
    );
  }
};

const changeStatus = async (email, status) => {
  try {
    const existingUser = await User.findOne({ email });

    if (!existingUser) {
      throw new Error("No se encontró al usuario");
    }

    if (!existingUser.verified) {
      throw new Error("Ésta cuenta de usuario aún no ha sido verificado.");
    }

    return await User.findOneAndUpdate(
      { email },
      {
        $set: {
          status: status,
        },
      },
      {
        new: true,
        fields: ["name lastname dni phone email"],
      }
    );
  } catch (error) {
    throw new Error(
      "Error al actualizar la información del usuario. Detalles: " +
        error.message
    );
  }
};

const changeRole = async (email, role) => {
  try {
    const existingUser = await User.findOne({ email });

    if (!existingUser) {
      throw new Error("No se encontró al usuario");
    }

    if (!existingUser.verified) {
      throw new Error("Ésta cuenta de usuario aún no ha sido verificado.");
    }

    return await User.findOneAndUpdate(
      { email },
      {
        $set: {
          role,
        },
      },
      {
        new: true,
        fields: ["name lastname dni phone email"],
      }
    );
  } catch (error) {
    throw new Error(
      "Error al actualizar la información del usuario. Detalles: " +
        error.message
    );
  }
};

const deleteUser = async (email) => {
  try {
    const existingUser = await User.findOne({ email });

    if (!existingUser) {
      throw new Error("No existe un usuario con ese correo");
    }

    return await User.updateOne(
      { email },
      {
        $set: {
          status: "inactive",
        },
      },
      {
        new: true,
        fields: ["name lastname dni phone email"],
      }
    );
  } catch (error) {
    throw new Error("Error al eliminar el usuario. Detalles: " + error.message);
  }
};

const deleteAllUsers = async (userIds) => {
  try {
    await User.updateMany(
      { _id: { $in: userIds } },
      { $set: { status: "inactive" } }
    );
  } catch (error) {
    throw new Error(
      "Error al eliminar los usuarios. Detalles: " + error.message
    );
  }
};

module.exports = {
  getUsers,
  getUser,
  verifyLink,
  infoUser,
  createUser,
  updateUser,
  changeStatus,
  changeRole,
  deleteUser,
  deleteAllUsers,
};

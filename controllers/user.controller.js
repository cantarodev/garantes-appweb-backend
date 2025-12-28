const User = require("../models/user.model");
const bcrypt = require("bcrypt");
const userService = require("../services/user.service");
const {
  sendVerificationEmail,
  sendResetEmail,
} = require("../utils/email.utils");

const createUser = async (req, res) => {
  try {
    const { name, lastname, dni, phone, business_name, ruc, email, password } =
      req.body;

    const user = await userService.createUser(
      name,
      lastname,
      dni,
      phone,
      business_name,
      ruc,
      email,
      password
    );
    await sendVerificationEmail(user, res);
  } catch (error) {
    console.error(error);
    return res.status(500).json({
      status: "error",
      data: null,
      message: error.message,
    });
  }
};

const changeStatus = async (req, res) => {
  const { email, status } = req.params;
  try {
    const user = await userService.changeStatus(email, status);
    res.status(200).json({
      status: "success",
      data: user,
      message: "Estado del usuario actualizado.",
    });
  } catch (error) {
    console.error(error);
    res.status(500).json({
      status: "error",
      data: null,
      message: "No se pudo actualizar los datos del usuario.",
    });
  }
};

const changeRole = async (req, res) => {
  console.log("PARAMS: ", req.params);
  const { email, role } = req.params;
  try {
    const user = await userService.changeRole(email, role);
    res.status(200).json({
      status: "success",
      data: user,
      message: "Rol del usuario actualizado.",
    });
  } catch (error) {
    console.error(error);
    res.status(500).json({
      status: "error",
      data: null,
      message: "No se pudo actualizar los datos del usuario.",
    });
  }
};

const verifyLink = async (req, res) => {
  const { userId, resetString } = req.params;
  try {
    const user = await userService.verifyLink(userId, resetString);
    if (!user) throw new Error("El enlace ha expirado o ya fue utilizado.");

    return res.status(200).json({
      status: "success",
      message: "El enlace sigue vigente para ser utilizado.",
    });
  } catch (error) {
    console.error(error);
    return res.status(500).json({
      status: "error",
      message: error.message,
    });
  }
};

const verifyUser = async (req, res) => {
  const { userId, uniqueString } = req.params;
  User.findOne({ _id: userId })
    .then((user) => {
      if (
        !(
          Object.keys(user.verification_email).length === 0 &&
          user.verification_email.constructor === Object
        )
      ) {
        const expireAt = user.verification_email.expire_at;
        const hashedUniqueString = user.verification_email.unique_string;
        if (expireAt < Date.now()) {
          User.findByIdAndDelete({ _id: userId })
            .then(() => {
              res.json({
                status: "FAILED",
                message: "El link ha expirado. Por favor regístrate de nuevo.",
              });
            })
            .catch((error) => {
              console.log(error);
              res.json({
                status: "FAILED",
                message:
                  "Error al borrar el usuario con una cadena única caducada.",
              });
            });
        } else {
          bcrypt
            .compare(uniqueString, hashedUniqueString)
            .then((result) => {
              if (result) {
                User.updateOne(
                  { _id: userId },
                  {
                    $set: {
                      verified: true,
                      verification_email: {},
                    },
                  }
                )
                  .then(() => {
                    res.status(200).json({
                      status: "SUCCESS",
                      message: "La verificación se realizó de forma exitosa.",
                    });
                  })
                  .catch((error) => {
                    console.log(error);
                    res.json({
                      status: "FAILED",
                      message:
                        "Se produjo un error al actualizar el registro de usuario para mostrar verificado.",
                    });
                  });
              } else {
                res.json({
                  status: "FAILED",
                  message:
                    "Se pasaron detalles de verificación no válidos. Revisa tu correo.",
                });
              }
            })
            .catch((error) => {
              console.log(error);
              res.json({
                status: "FAILED",
                message: "Se produjo un error al comparar cadenas únicas.",
              });
            });
        }
      } else {
        res.json({
          status: "FAILED",
          message:
            "El registro de la cuenta no existe o ya se ha verificado. Por favor regístrate o inicia sesión.",
        });
      }
    })
    .catch((error) => {
      console.log(error);
      res.json({
        status: "FAILED",
        message:
          "Se produjo un error al verificar el registro de verificación de usuario existente.",
      });
    });
};

const requestResetPassword = async (req, res) => {
  const { email, redirectUrl } = req.body;

  try {
    const user = await User.findOne({ email });

    if (!user) {
      throw new Error(
        "No existe ninguna cuenta con el correo electrónico proporcionado."
      );
    }

    if (!user.verified) {
      throw new Error(
        "El correo electrónico aún no ha sido verificado. Revisa tu correo."
      );
    }

    await sendResetEmail(
      { user_id: user._id, email: user.email },
      redirectUrl,
      res
    );
  } catch (error) {
    console.error(error);
    return res.status(500).json({
      status: "error",
      message: error.message,
    });
  }
};

const resetPassword = async (req, res) => {
  const { userId, resetString, newPassword } = req.body;

  try {
    const user = await User.findOne({ _id: userId });

    if (
      !(
        Object.keys(user.password_reset).length === 0 &&
        user.password_reset.constructor === Object
      )
    ) {
      const hashedResetString = user.password_reset.reset_string;

      if (user.password_reset.expire_at < Date.now()) {
        await User.updateOne(
          { _id: userId },
          {
            $set: {
              password_reset: {},
            },
          }
        );
        res.json({
          status: "error",
          message: "El enlace para restablecer la contraseña ha caducado.",
        });
      } else {
        const result = await bcrypt.compare(resetString, hashedResetString);

        if (result) {
          const saltRounds = 10;
          const hashedNewPassword = await bcrypt.hash(newPassword, saltRounds);

          const updatedUser = await User.findByIdAndUpdate(
            { _id: userId },
            {
              $set: {
                password: hashedNewPassword,
                password_reset: {},
              },
            },
            { new: true }
          );

          res.json({
            status: "success",
            message: "La contraseña se ha restablecido correctamente.",
            data: { email: updatedUser.email },
          });
        } else {
          res.json({
            status: "error",
            message:
              "Se pasaron detalles de restablecimiento de contraseña no válidos.",
          });
        }
      }
    } else {
      res.json({
        status: "error",
        message: "Solicitud de restablecimiento de contraseña no encontrada.",
      });
    }
  } catch (error) {
    console.log(error);
    res.json({
      status: "error",
      message:
        "Error al procesar la solicitud de restablecimiento de contraseña.",
    });
  }
};

const getUsers = async (req, res) => {
  try {
    const users = await userService.getUsers();
    res.status(200).json({
      status: "success",
      data: users,
      message: "Usuarios recuperados exitosamente.",
    });
  } catch (error) {
    console.error("Error fetching users:", error);
    return res.status(500).json({
      status: "error",
      data: null,
      message: "Error al recuperar los usuarios. Por favor, intente más tarde.",
    });
  }
};

const getUser = async (req, res) => {
  const { user_id } = req.user;

  try {
    const user = await userService.getUser(user_id);
    res.status(200).json({
      status: "success",
      data: user,
      message: "Datos del usuario recuperados exitosamente.",
    });
  } catch (error) {
    console.error(error);
    return res.status(500).json({
      status: "error",
      data: null,
      message: "Error en la operación.",
    });
  }
};

const infoUser = async (req, res) => {
  const { email, password } = req.body;

  try {
    const user = await userService.infoUser(email, password);

    // res.cookie("authToken", user.token, {
    //   httpOnly: true,
    //   secure: process.env.NODE_ENV === "production",
    //   sameSite: "None",
    //   maxAge: 2 * 60 * 1000,
    // });

    res.status(200).json({
      status: "success",
      data: user,
      message: "Datos del usuario recuperados exitosamente.",
    });
  } catch (error) {
    console.error(error);
    res.status(500).json({
      status: "error",
      data: null,
      message: error.message,
    });
  }
};

const updateUser = async (req, res) => {
  const { avatar, name, lastname, email, password } = req.body;

  try {
    const user = await userService.updateUser(
      avatar,
      name,
      lastname,
      email,
      password
    );

    if (!user) {
      throw new Error("User not found");
    }

    res.status(200).json({
      status: "success",
      data: user,
      message: "Datos actualizados de forma exitosa.",
    });
  } catch (error) {
    console.error(error);
    res.status(500).json({
      status: "error",
      data: null,
      message: "No se pudo actualizar los datos del usuario.",
    });
  }
};

const deleteUser = async (req, res) => {
  try {
    const { email } = req.params;

    const user = await userService.deleteUser(email);

    res.status(200).json({
      status: "success",
      data: user,
      message: "Usuario eliminado de forma exitosa.",
    });
  } catch (error) {
    console.error(error);
    res.status(500).json({
      status: "error",
      data: null,
      message: "No se pudo eliminar el usuario.",
    });
  }
};

const deleteAllUsers = async (req, res) => {
  try {
    const { userIds } = req.params;

    await userService.deleteAllUsers(userIds);

    res.status(200).json({
      status: "success",
      message: "Usuarios eliminados de forma exitosa.",
    });
  } catch (error) {
    console.error(error);
    res.status(500).json({
      status: "error",
      data: null,
      message: "No se lograron eliminar los usuarios.",
    });
  }
};

module.exports = {
  getUsers,
  getUser,
  changeStatus,
  changeRole,
  verifyLink,
  verifyUser,
  requestResetPassword,
  resetPassword,
  infoUser,
  createUser,
  updateUser,
  deleteUser,
  deleteAllUsers,
};

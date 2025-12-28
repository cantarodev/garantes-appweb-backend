const { refreshTokenService } = require("../services/auth.service");

const refreshToken = async (req, res) => {
  const { lastToken } = req.body;

  try {
    const token = await refreshTokenService(lastToken);
    return res.status(200).json({ token });
  } catch (err) {
    return res
      .status(err.status || 500)
      .json({ message: err.message || "Something went wrong" });
  }
};

module.exports = { refreshToken };

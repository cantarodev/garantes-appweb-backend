const jwt = require("jsonwebtoken");
const { generateAccessToken } = require("../utils/token.utils");

const refreshTokenService = async (lastToken) => {
  if (!lastToken) {
    throw { status: 401, message: "Refresh token is required" };
  }

  try {
    const { user_id, email, name, role } = await jwt.verify(
      lastToken,
      process.env.ACCESS_TOKEN_SECRET
    );

    const token = generateAccessToken({ user_id, email, name, role });
    return token;
  } catch (err) {
    throw { staus: 403, message: "Invalid refresh token" };
  }
};

module.exports = { refreshTokenService };

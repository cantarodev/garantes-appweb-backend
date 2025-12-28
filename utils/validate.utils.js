const axios = require("axios");

const API_URL = process.env.API_LAMBDA_VALIDATE;

async function validateAccount(ruc, username, password) {
  try {
    console.log("account:", username, password);

    const response = await axios.post(API_URL, {
      ruc,
      username,
      password,
    });

    return { status: response.status, data: response.data };
  } catch (error) {
    const { response } = error;
    console.error("Error:", error.response);
    return { status: response.status, data: {} };
  }
}

module.exports = { validateAccount };

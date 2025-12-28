const axios = require("axios");

const instance = axios.create({
  timeout: 8000,
});

instance.interceptors.response.use(
  (res) => res,
  (err) => Promise.reject(err)
);

module.exports = instance;

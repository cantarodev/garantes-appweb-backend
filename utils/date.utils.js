const { format, parse } = require("date-fns");
const { es } = require("date-fns/locale");

const convertPeriodToMonth = (period) => {
  const date = parse(period, "yyyyMM", new Date());
  return format(date, "MMMM", { locale: es });
};

const currentDate = () => {
  return format(new Date(), "yyyyMMdd");
};

module.exports = { convertPeriodToMonth, currentDate };

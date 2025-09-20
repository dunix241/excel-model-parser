const XLSX = require("xlsx");
const fs = require("fs");

const workbook = XLSX.readFile("input.xlsx");

const models = [
  require("./models/tourist-count"),
  require("./models/revenue"),
  require("./models/grdp"),
  require("./models/flights"),
  require("./models/voyages"),
  require("./models/accommodation-facilities"),
];

models.forEach((model) => {
  const { map, includeKeys, excludeKeys, guestRanges } = model.config;
  let data = model.process(workbook);
  if (guestRanges && guestRanges.length) {
    data = data.map((obj, index) => {
      return {
        ...obj,
        Guests:
          (guestRanges[index] &&
            require("./models/guest-info").process(
              workbook,
              guestRanges[index],
            )) ||
          [],
      };
    });
  }
  if (model.config.map) {
    data = data.map(map);
  }
  if (excludeKeys && excludeKeys.length) {
    data = data.map((obj) =>
      Object.keys(obj)
        .filter((key) => !excludeKeys.includes(key))
        .reduce((res, key) => ((res[key] = obj[key]), res), {}),
    );
  }
  fs.writeFileSync(
    `./outs/${model.config.output}`,
    JSON.stringify(data, null, 2),
  );
});

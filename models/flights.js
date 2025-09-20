const XLSX = require("xlsx");
const fs = require("fs");

module.exports = {
  config: {
    sheet: "Dữ liệu Du lịch",
    range: "A30:I41",
    output: "flights.json",
    headers: [
      "Direction",
      "IdNumber",
      "Airline",
      "Route",
      "DepartureTime",
      "ArrivalTime",
      "_",
      "Status",
      "Location",
    ],
    handlers: [
      null,
      null,
      null,
      null,
      (cell) => new Date(cell),
      // (cell) => +`${cell}`.replaceAll(".", "") * 1_000_000_000,
      // (cell) => `${cell}`.replaceAll(",", "."),
      // (cell) => +`${cell}`.replaceAll(",", ".") + 0,
      // (cell) => `${cell}`.replaceAll(".", ""),
    ],
    map: (flight) => {
      const [lng, lat] = flight.Location.split(", ");
      return { ...flight, Lng: +lng, Lat: +lat };
    },
    includeKeys: [],
    excludeKeys: ["_", "Location"],
  },
  process(workbook) {
    const worksheet = workbook.Sheets[this.config.sheet];
    if (!worksheet) return;
    const rows = XLSX.utils.sheet_to_json(worksheet, {
      range: this.config.range,
      header: 1,
    });
    const jsonObjects = rows.map((row) =>
      Object.fromEntries(
        this.config.headers.map((header, i) => [
          header,
          (this.config.handlers[i] && this.config.handlers[i](row[i])) ||
            row[i],
        ]),
      ),
    );

    return jsonObjects;
  },
};

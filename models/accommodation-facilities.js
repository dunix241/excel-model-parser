const XLSX = require("xlsx");
const fs = require("fs");

module.exports = {
  config: {
    sheet: "Dữ liệu Du lịch",
    range: "A74:G83",
    output: "accommodation-facilities.json",
    headers: [
      "Name",
      "_",
      "CheckinCount",
      "CheckoutCount",
      "Status",
      "__",
      "Location",
    ],
    handlers: [
      null,
      null,
      null,
      null,
      (cell) =>
        typeof cell === "number"
          ? new Date(Math.round(cell * 86400000)).toISOString().substr(11, 8)
          : cell,
      (cell) =>
        typeof cell === "number"
          ? new Date(Math.round(cell * 86400000)).toISOString().substr(11, 8)
          : cell,
      // (cell) => new Date(cell),
      // (cell) => new Date(cell),
      // (cell) => +`${cell}`.replaceAll(".", "") * 1_000_000_000,
      // (cell) => `${cell}`.replaceAll(",", "."),
      // (cell) => +`${cell}`.replaceAll(",", ".") + 0,
      // (cell) => `${cell}`.replaceAll(".", ""),
    ],
    map: (flight) => {
      const [lng, lat] = flight.Location.split(", ");
      return { ...flight, Lng: +lng, Lat: +lat, StayingGuests: flight.Guests };
    },
    includeKeys: [],
    excludeKeys: ["_", "__", "Location", "Guests"],
    guestRanges: [
      "A5:D9",
      false,
      "A23:D27",
      false,
      "A41:D45",
      "A50:D54",
      false,
      "A63:D67",
      "A72:D76",
      false,
    ],
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

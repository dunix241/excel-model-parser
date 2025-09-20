const XLSX = require("xlsx");
const fs = require("fs");

module.exports = {
  config: {
    sheet: "Dữ liệu Du lịch",
    range: "A57:H64",
    output: "voyages.json",
    headers: [
      "Direction",
      "IdNumber",
      "DepartureTime",
      "ArrivalTime",
      "GuestCount",
      "_",
      "Status",
      "Location",
    ],
    handlers: [
      null,
      null,
      (cell) =>
        typeof cell === "number"
          ? new Date(Math.round(cell * 86400000)).toISOString().substr(11, 8)
          : cell,
      (cell) => {
        return cell.split(" ")[0] + ":00";
      },
      // (cell) => new Date(cell),
      // (cell) => new Date(cell),
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
    excludeKeys: ["_", "Location", "GuestCount"],
    guestRanges: [
      "I5:L9",
      "I14:L18",
      "I23:L27",
      "I32:L36",
      "I41:L45",
      "I50:L54",
      "I59:L63",
      "I68:L72",
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

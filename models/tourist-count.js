const XLSX = require("xlsx");
const fs = require("fs");

module.exports = {
  config: {
    sheet: "Dữ liệu Du lịch",
    range: "A6:E9",
    output: "tourist-count.json",
    headers: ["Year", "TouristCount", "Domestic", "International", "Growth"],
    handlers: [
      null,
      (cell) => +`${cell}`.replaceAll(".", ""),
      (cell) => `${cell}`.replaceAll(".", ""),
      (cell) => `${cell}`.replaceAll(".", ""),
    ],
    map: null,
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

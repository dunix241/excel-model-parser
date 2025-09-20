const XLSX = require("xlsx");
const fs = require("fs");

module.exports = {
  config: {
    sheet: "Chi tiết hành khách",
    output: "guest-info.json",
    headers: ["UpdatedDate", "FullName", "Gender", "YearOfBirth"],
    handlers: [
      (cell) => {
        const [day, month, year] = cell.split("/");
        return {
          $date: new Date(+year, +month - 1, +day).toISOString(),
        };
      },
      null,
      null,
      (cell) => `${cell}`,
    ],
    map: null,
  },
  process(workbook, range) {
    const worksheet = workbook.Sheets[this.config.sheet];
    if (!worksheet) return;
    const rows = XLSX.utils.sheet_to_json(worksheet, {
      range: range,
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

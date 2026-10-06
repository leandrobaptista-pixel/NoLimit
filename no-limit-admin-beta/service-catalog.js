(function (root) {
  "use strict";
  const columns = ["Product", "Service Full Name", "Type", "Memo/Description", "Quantity", "Sales Price"];
  const text = value => value == null ? "" : String(value).trim();
  const key = value => text(value).normalize("NFKC").toLocaleLowerCase("en-US");
  function price(value) {
    if (value == null || text(value) === "") return null;
    const number = typeof value === "number" ? value : Number(text(value).replace(/^\$/, "").replaceAll(",", ""));
    if (!Number.isFinite(number) || number < 0) throw new Error(`Invalid sales price: ${value}`);
    return number;
  }
  function fromRows(rows, sourceFile = "") {
    if (!Array.isArray(rows) || !rows.length) throw new Error("The import has no rows.");
    const header = rows[0].map(value => text(value).replace(/^\uFEFF/, ""));
    if (columns.some(column => !header.includes(column))) throw new Error(`Expected columns: ${columns.join(", ")}`);
    if (new Set(header).size !== header.length) throw new Error("Duplicate column headings.");
    return rows.slice(1).flatMap((row, index) => {
      if (!Array.isArray(row)) throw new Error(`Invalid row ${index + 2}.`);
      if (!row.some(value => text(value) !== "")) return [];
      const raw = Object.fromEntries(columns.map(column => [column, row[header.indexOf(column)] ?? null]));
      if (!text(raw["Service Full Name"])) throw new Error(`Row ${index + 2}: service name is missing.`);
      return [{ title: text(raw["Service Full Name"]), product: text(raw.Product), type: text(raw.Type),
        category: text(raw.Product) || "Custom / New Work", description: text(raw["Memo/Description"]),
        unit: text(raw.Quantity), unitPrice: price(raw["Sales Price"]), archived: false,
        importSource: { system: "QuickBooks", file: sourceFile, row: index + 2, values: raw } }];
    });
  }
  function csvRows(input) {
    const source = String(input).replace(/^\uFEFF/, "");
    const delimiter = source.split(/\r?\n/, 1)[0].includes(";") ? ";" : ",";
    const rows = []; let row = [], cell = "", quoted = false;
    for (let i = 0; i < source.length; i++) {
      const ch = source[i];
      if (ch === '"' && (quoted || cell === "")) {
        if (quoted && source[i + 1] === '"') { cell += '"'; i++; }
        else quoted = !quoted;
      } else if (!quoted && ch === delimiter) { row.push(cell); cell = ""; }
      else if (!quoted && (ch === "\n" || ch === "\r")) {
        if (ch === "\r" && source[i + 1] === "\n") i++;
        row.push(cell); rows.push(row); row = []; cell = "";
      } else cell += ch;
    }
    if (quoted) throw new Error("The CSV contains an unclosed quoted field.");
    if (cell !== "" || row.length) { row.push(cell); rows.push(row); }
    return rows;
  }
  function parse(input, fileName) {
    if (/\.json$/i.test(fileName)) {
      const bundle = JSON.parse(input);
      if (bundle.schemaVersion !== 1 || !Array.isArray(bundle.rows)) throw new Error("Unsupported service import file.");
      return fromRows(bundle.rows, text(bundle.sourceFile) || fileName);
    }
    if (!/\.csv$/i.test(fileName)) throw new Error("Use the prepared QuickBooks import file or a CSV with the same six columns.");
    return fromRows(csvRows(input), fileName);
  }
  function plan(existing, incoming) {
    const names = new Set(existing.filter(Boolean).map(item => key(item.title)));
    const additions = [], skipped = [];
    for (const item of incoming) {
      const name = key(item.title);
      if (!name) throw new Error("A service name is missing.");
      if (names.has(name)) { skipped.push(item); continue; }
      names.add(name); additions.push(item);
    }
    return { additions, skipped };
  }
  function billable(service) { return Boolean(service) && !service.archived && key(service.type) !== "location"; }
  root.NoLimitServiceCatalog = Object.freeze({ columns, fromRows, parse, plan, key, price, billable });
})(globalThis);

export type MaterialChargeTicketHeader = {
    district: string;
    department: string;
    requestNumber: string;
    requestDate: string;
    requisitioner: string;
    releaseDate: string;
    mctRelNumber: string;
    woNumber: string;
    joNumber: string;
    soNumber: string;
    purpose: string;
    notes: string;
};

export type MaterialChargeTicketItem = {
    id: string;
    item_code: string;
    particulars: string;
    unit: string;
    unit_cost: number | null;
    qty: number | null;
    total_cost: number | null;
    c2: number | null;
    deduct_from: "ending_qty" | "buffer_stock";
    purpose: string;
    remarks: string;
    notes: string;
};

export const EMPTY_HEADER: MaterialChargeTicketHeader = {
    district: "",
    department: "",
    requestNumber: "",
    requestDate: "",
    requisitioner: "",
    releaseDate: "",
    mctRelNumber: "",
    woNumber: "",
    joNumber: "",
    soNumber: "",
    purpose: "",
    notes: "",
};

const HEADER_KEY_MAP: Record<string, keyof MaterialChargeTicketHeader> = {
    district: "district",
    department: "department",
    "request#": "requestNumber",
    requestno: "requestNumber",
    reqno: "requestNumber",
    reqdate: "requestDate",
    requestdate: "requestDate",
    "req.date": "requestDate",
    requisitioner: "requisitioner",
    reldate: "releaseDate",
    releasedate: "releaseDate",
    "rel.date": "releaseDate",
    "mct/rel#": "mctRelNumber",
    "mctrel#": "mctRelNumber",
    mctrelno: "mctRelNumber",
    "mct/relno": "mctRelNumber",
    "wo#": "woNumber",
    "jo#": "joNumber",
    "so#": "soNumber",
    purpose: "purpose",
    notes: "notes",
    "notes/sr#": "notes",
    "sr#": "notes",
};

const ITEM_KEY_MAP: Record<string, keyof MaterialChargeTicketItem> = {
    itemcode: "item_code",
    "itemcode#": "item_code",
    item: "item_code",
    code: "item_code",
    particulars: "particulars",
    description: "particulars",
    unit: "unit",
    type: "unit",
    uom: "unit",
    unitcost: "unit_cost",
    "unitcost(php)": "unit_cost",
    qty: "qty",
    quantity: "qty",
    totalcost: "total_cost",
    amount: "total_cost",
    total: "total_cost",
    c2: "c2",
    purpose: "purpose",
    remarks: "notes",
    notes: "notes",
    "notes/sr#": "notes",
    "sr#": "notes",
};

const parseNumber = (value: string) => {
    if (!value) return null;
    const sanitized = value.replace(/[,\s]/g, "").replace(/[^0-9.+-]/g, "");
    if (!sanitized) return null;
    const parsed = Number(sanitized);
    return Number.isFinite(parsed) ? parsed : null;
};

const normalizeCell = (value: string) =>
    value
        .toLowerCase()
        .replace(/\s+/g, "")
        .replace(/[^a-z0-9#/.:_-]/g, "");

export const normalizeItemCode = (value: string) => value.trim().toUpperCase();

export const normalizeRows = (rows: Array<Array<string | number | null | undefined>>) =>
    rows
        .map((row) => row.map((cell) => String(cell ?? "").trim()))
        .filter((row) => row.some((cell) => cell.length > 0));

const parseLabelValuePairs = (rows: string[][]) => {
    const header: MaterialChargeTicketHeader = { ...EMPTY_HEADER };

    rows.forEach((row) => {
        row.forEach((cell, index) => {
            if (!cell) return;
            const trimmed = cell.trim();
            const normalized = normalizeCell(trimmed);

            if (normalized.includes(":") && !normalized.endsWith(":")) {
                const [label, ...valueParts] = trimmed.split(":");
                const mappedKey = HEADER_KEY_MAP[normalizeCell(label)];
                if (!mappedKey) return;
                const inlineValue = valueParts.join(":").trim();
                if (inlineValue && !header[mappedKey]) {
                    header[mappedKey] = inlineValue;
                }
                return;
            }

            if (!normalized.endsWith(":")) return;
            const mappedKey = HEADER_KEY_MAP[normalized.replace(/:$/, "")];
            if (!mappedKey) return;

            for (let next = index + 1; next < row.length; next += 1) {
                const value = row[next]?.trim();
                if (!value) continue;
                const normalizedValue = normalizeCell(value);
                if (
                    normalizedValue.endsWith(":") ||
                    HEADER_KEY_MAP[normalizedValue] ||
                    (normalizedValue.includes(":") &&
                        HEADER_KEY_MAP[normalizeCell(value.split(":")[0] ?? "")])
                ) {
                    continue;
                }
                if (!header[mappedKey]) {
                    header[mappedKey] = value;
                }
                break;
            }
        });
    });

    return header;
};

const findItemsHeaderRow = (rows: string[][]) => {
    for (let rowIndex = 0; rowIndex < rows.length; rowIndex += 1) {
        const row = rows[rowIndex];
        const normalizedRow = row.map((cell) => normalizeCell(cell));
        const hasItemCode = normalizedRow.some((cell) => cell === "itemcode");
        const hasParticulars = normalizedRow.some((cell) => cell === "particulars" || cell === "description");
        const hasUnitCost = normalizedRow.some((cell) => cell === "unitcost");
        const hasQty = normalizedRow.some((cell) => cell === "qty" || cell === "quantity");

        if (hasItemCode && hasParticulars && hasUnitCost && hasQty) {
            return { rowIndex, row };
        }
    }
    return null;
};

const parseItemsFromTable = (rows: string[][]) => {
    const headerInfo = findItemsHeaderRow(rows);
    if (!headerInfo) return [];

    const { rowIndex, row } = headerInfo;
    const headerMap = row.map((cell) => ITEM_KEY_MAP[normalizeCell(cell)] ?? null);
    const items: MaterialChargeTicketItem[] = [];

    for (let i = rowIndex + 1; i < rows.length; i += 1) {
        const currentRow = rows[i];
        const joined = currentRow.join(" ").toLowerCase();

        if (!currentRow.some((cell) => cell.trim().length > 0)) {
            continue;
        }

        if (joined.includes("total") || joined.includes("nothingfollows") || joined.includes("purpose")) {
            break;
        }

        const item: MaterialChargeTicketItem = {
            id: `mct-row-${items.length + 1}`,
            item_code: "",
            particulars: "",
            unit: "",
            unit_cost: null,
            qty: null,
            total_cost: null,
            c2: null,
            deduct_from: "ending_qty",
            purpose: "",
            remarks: "",
            notes: "",
        };

        let hasItemValue = false;

        currentRow.forEach((cell, colIndex) => {
            const key = headerMap[colIndex];
            if (!key) return;
            const value = cell.trim();
            if (!value) return;

            if (key === "deduct_from") {
                item.deduct_from = value.toLowerCase().includes("buffer") ? "buffer_stock" : "ending_qty";
            } else if (key === "unit_cost" || key === "qty" || key === "total_cost" || key === "c2") {
                item[key] = parseNumber(value);
            } else {
                item[key] = value;
            }
            hasItemValue = true;
        });

        if (hasItemValue && (item.item_code || item.particulars)) {
            items.push(item);
        }
    }

    return items;
};

export const parseCsvRows = (text: string) => {
    const rows: string[][] = [];
    let current: string[] = [];
    let field = "";
    let inQuotes = false;

    for (let i = 0; i < text.length; i += 1) {
        const char = text[i];
        const next = text[i + 1];

        if (inQuotes) {
            if (char === '"' && next === '"') {
                field += '"';
                i += 1;
            } else if (char === '"') {
                inQuotes = false;
            } else {
                field += char;
            }
            continue;
        }

        if (char === '"') {
            inQuotes = true;
            continue;
        }

        if (char === ",") {
            current.push(field);
            field = "";
            continue;
        }

        if (char === "\n") {
            current.push(field);
            rows.push(current);
            current = [];
            field = "";
            continue;
        }

        if (char === "\r") {
            continue;
        }

        field += char;
    }

    if (field.length > 0 || current.length > 0) {
        current.push(field);
        rows.push(current);
    }

    return rows;
};

export const parseRowsToTicket = (rows: string[][]) => {
    if (!rows.length) {
        return { header: { ...EMPTY_HEADER }, items: [] };
    }

    return {
        header: parseLabelValuePairs(rows),
        items: parseItemsFromTable(rows),
    };
};

export const formatDecimal = (value: number | null) => {
    if (value == null || Number.isNaN(value)) return "-";
    return value.toLocaleString("en-US", {
        minimumFractionDigits: 2,
        maximumFractionDigits: 2,
    });
};

export const formatC2 = (value: number | null) => {
    if (value == null || Number.isNaN(value)) return "-";
    return value.toLocaleString("en-US", {
        maximumFractionDigits: 0,
    });
};

export const sumNumbers = (values: Array<number | null>) =>
    values.reduce<number>((acc, value) => (value == null || Number.isNaN(value) ? acc : acc + value), 0);

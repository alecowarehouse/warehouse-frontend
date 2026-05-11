import { CreateView, CreateViewHeader } from "@/components/refine-ui/views/create-view";
import MultiUploadWidget from "@/components/multi-upload-widget";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardFooter, CardHeader, CardTitle } from "@/components/ui/card";
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from "@/components/ui/table";
import { Textarea } from "@/components/ui/textarea";
import { Tooltip, TooltipContent, TooltipTrigger } from "@/components/ui/tooltip";
import {
    Select,
    SelectContent,
    SelectItem,
    SelectTrigger,
    SelectValue,
} from "@/components/ui/select";
import {
    AlertDialog,
    AlertDialogAction,
    AlertDialogCancel,
    AlertDialogContent,
    AlertDialogDescription,
    AlertDialogFooter,
    AlertDialogHeader,
    AlertDialogTitle,
} from "@/components/ui/alert-dialog";
import {
    Dialog,
    DialogContent,
    DialogDescription,
    DialogFooter,
    DialogHeader,
    DialogTitle,
} from "@/components/ui/dialog";
import { useEffect, useMemo, useState } from "react";
import * as XLSX from "xlsx";
import { supabaseClient } from "@/providers/supabase-client";
import { useGetIdentity, useGo, useInvalidate, useNotification } from "@refinedev/core";
import { Badge } from "@/components/ui/badge";
import { ChevronDown, ChevronRight } from "lucide-react";

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

type AvailabilityStatus = "in_stock" | "insufficient" | "missing";

type AvailabilityInfo = {
    status: AvailabilityStatus;
    availableQty?: number | null;
    bufferStock?: number | null;
    endingQty?: number | null;
    deductFrom?: "ending_qty" | "buffer_stock";
};

type ParsedMctTicket = {
    id: string;
    fileName: string;
    header: MaterialChargeTicketHeader;
    items: MaterialChargeTicketItem[];
    status: "ready" | "saving" | "saved" | "error";
    error: string | null;
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
    "requestno": "requestNumber",
    reqno: "requestNumber",
    reqdate: "requestDate",
    "requestdate": "requestDate",
    "req.date": "requestDate",
    requisitioner: "requisitioner",
    reldate: "releaseDate",
    "releasedate": "releaseDate",
    "rel.date": "releaseDate",
    "mct/rel#": "mctRelNumber",
    "mctrel#": "mctRelNumber",
    "mctrelno": "mctRelNumber",
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
                const labelKey = normalizeCell(label);
                const mappedKey = HEADER_KEY_MAP[labelKey];
                if (!mappedKey) return;
                const inlineValue = valueParts.join(":").trim();
                if (inlineValue && !header[mappedKey]) {
                    header[mappedKey] = inlineValue;
                }
                return;
            }

            if (!normalized.endsWith(":")) return;
            const labelKey = normalized.replace(/:$/, "");
            const mappedKey = HEADER_KEY_MAP[labelKey];
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
            return { rowIndex, row, normalizedRow };
        }
    }
    return null;
};

const parseItemsFromTable = (rows: string[][]) => {
    const headerInfo = findItemsHeaderRow(rows);
    if (!headerInfo) return [];

    const { rowIndex, row } = headerInfo;
    const headerMap = row.map((cell) => {
        const normalized = normalizeCell(cell);
        return ITEM_KEY_MAP[normalized] ?? null;
    });

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

    const header = parseLabelValuePairs(rows);
    const items = parseItemsFromTable(rows);

    return { header, items };
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

const createTicketId = (file: File, index = 0) => `mct-upload-${file.name}-${file.size}-${file.lastModified}-${index}`;

const normalizeReportCell = (value: string) =>
    value
        .toLowerCase()
        .replace(/\s+/g, "")
        .replace(/[^a-z0-9#/.-]/g, "");

const isReportHeaderRow = (row: string[]) => {
    const normalized = row.map((cell) => normalizeReportCell(cell));
    return (
        normalized.some((cell) => cell.includes("ref#") || cell.includes("item#")) &&
        normalized.some((cell) => cell === "date") &&
        normalized.some((cell) => cell.includes("requisitioner")) &&
        normalized.some((cell) => cell.includes("purpose"))
    );
};

type ReportColumnMap = {
    refIndex: number;
    dateIndex: number;
    requisitionerIndex: number;
    supervisorIndex: number | null;
    preparedByIndex: number | null;
    purposeIndex: number;
    totalQtyIndex: number;
    totalAmountIndex: number;
};

const findColumnIndex = (normalizedRow: string[], matcher: (cell: string) => boolean) => {
    const index = normalizedRow.findIndex(matcher);
    return index >= 0 ? index : null;
};

const getReportColumnMap = (headerRow: string[]): ReportColumnMap | null => {
    const normalizedRow = headerRow.map((cell) => normalizeReportCell(cell));
    const refIndex = findColumnIndex(normalizedRow, (cell) => cell.includes("ref#") || cell.includes("item#"));
    const dateIndex = findColumnIndex(normalizedRow, (cell) => cell === "date");
    const requisitionerIndex = findColumnIndex(normalizedRow, (cell) => cell.includes("requisitioner"));
    const supervisorIndex = findColumnIndex(normalizedRow, (cell) => cell.includes("supervisor"));
    const preparedByIndex = findColumnIndex(normalizedRow, (cell) => cell.includes("preparedby"));
    const purposeIndex = findColumnIndex(normalizedRow, (cell) => cell.includes("purpose") || cell.includes("particulars"));
    const totalQtyIndex = findColumnIndex(normalizedRow, (cell) => cell.includes("totalqty"));
    const totalAmountIndex = findColumnIndex(normalizedRow, (cell) => cell.includes("totalamount"));

    if (
        refIndex == null ||
        dateIndex == null ||
        requisitionerIndex == null ||
        purposeIndex == null ||
        totalQtyIndex == null ||
        totalAmountIndex == null
    ) {
        return null;
    }

    return {
        refIndex,
        dateIndex,
        requisitionerIndex,
        supervisorIndex,
        preparedByIndex,
        purposeIndex,
        totalQtyIndex,
        totalAmountIndex,
    };
};

const isReportTicketRow = (row: string[]) => /^rel[-#/]\w+/i.test((row[0] ?? "").trim());
const isReportItemRow = (row: string[]) => /^\d+\.?$/.test((row[0] ?? "").trim());

const isReportMetadataRow = (row: string[]) => {
    const firstCell = (row[0] ?? "").trim().toLowerCase();
    return (
        firstCell.startsWith("released by district") ||
        firstCell.startsWith("notes/serial") ||
        firstCell.startsWith("ref mct") ||
        firstCell.startsWith("s#")
    );
};

const getTrailingNumericIndexes = (row: string[]) =>
    row
        .map((cell, index) => ({ index, value: parseNumber(cell) }))
        .filter((entry) => entry.value != null)
        .map((entry) => entry.index);

const compactJoin = (values: Array<string | null | undefined>) =>
    values
        .map((value) => value?.trim() ?? "")
        .filter(Boolean)
        .join(" ");

const joinReportRange = (row: string[], startIndex: number, endIndex: number) =>
    compactJoin(row.slice(startIndex, Math.max(startIndex + 1, endIndex)));

const looksLikeReportDate = (value: string) =>
    /^\d{1,2}[/-]\d{1,2}[/-]\d{2,4}$/.test(value.trim());

const getReportDate = (row: string[], columnMap: ReportColumnMap) => {
    const mappedDate = row[columnMap.dateIndex]?.trim() ?? "";
    if (looksLikeReportDate(mappedDate)) {
        return mappedDate;
    }

    return row.find((cell) => looksLikeReportDate(cell))?.trim() ?? mappedDate;
};

const parseReportHeader = (
    row: string[],
    pendingNotes: string[],
    columnMap: ReportColumnMap
): MaterialChargeTicketHeader => {
    const requisitionerEndIndex = columnMap.supervisorIndex ?? columnMap.preparedByIndex ?? columnMap.purposeIndex;
    const relDate = getReportDate(row, columnMap);

    return {
        ...EMPTY_HEADER,
        requestNumber: row[columnMap.refIndex]?.trim() ?? "",
        requestDate: "",
        requisitioner: joinReportRange(row, columnMap.requisitionerIndex, requisitionerEndIndex),
        releaseDate: relDate,
        mctRelNumber: row[columnMap.refIndex]?.trim() ?? "",
        purpose: joinReportRange(row, columnMap.purposeIndex, columnMap.totalQtyIndex),
        notes: pendingNotes.join("\n"),
    };
};

const parseReportItem = (row: string[], ticketIndex: number, itemIndex: number): MaterialChargeTicketItem | null => {
    const numericIndexes = getTrailingNumericIndexes(row);
    const totalCostIndex = numericIndexes[numericIndexes.length - 1];
    const qtyIndex = numericIndexes[numericIndexes.length - 2];
    const unitCostIndex = numericIndexes[numericIndexes.length - 3];

    const itemCode = row[1]?.trim() ?? "";
    const descriptionEndIndex = unitCostIndex != null ? Math.max(3, unitCostIndex - 1) : row.length;
    const particulars = compactJoin(row.slice(2, descriptionEndIndex));
    const unit = unitCostIndex != null ? row[unitCostIndex - 1]?.trim() ?? "" : "";

    if (!itemCode && !particulars) {
        return null;
    }

    return {
        id: `mct-report-${ticketIndex}-row-${itemIndex}`,
        item_code: itemCode,
        particulars,
        unit,
        unit_cost: unitCostIndex != null ? parseNumber(row[unitCostIndex]) : null,
        qty: qtyIndex != null ? parseNumber(row[qtyIndex]) : null,
        total_cost: totalCostIndex != null ? parseNumber(row[totalCostIndex]) : null,
        c2: null,
        deduct_from: "ending_qty",
        purpose: "",
        remarks: "",
        notes: "",
    };
};

const parseMultiTicketReport = (rows: string[][], file: File) => {
    const headerRow = rows.find(isReportHeaderRow);
    const columnMap = headerRow ? getReportColumnMap(headerRow) : null;

    if (!headerRow || !columnMap || !rows.some(isReportTicketRow)) {
        return [];
    }

    const tickets: ParsedMctTicket[] = [];
    let currentTicket: ParsedMctTicket | null = null;
    let pendingNotes: string[] = [];

    const flushTicket = () => {
        if (currentTicket) {
            tickets.push(currentTicket);
        }
        currentTicket = null;
    };

    rows.forEach((row) => {
        if (isReportHeaderRow(row) || !row.some(Boolean)) {
            return;
        }

        if (isReportMetadataRow(row)) {
            const note = compactJoin(row);
            if (currentTicket) {
                currentTicket.header.notes = compactJoin([currentTicket.header.notes, note]);
            } else {
                pendingNotes.push(note);
            }
            return;
        }

        if (isReportTicketRow(row)) {
            flushTicket();
            currentTicket = {
                id: createTicketId(file, tickets.length),
                fileName: file.name,
                header: parseReportHeader(row, pendingNotes, columnMap),
                items: [],
                status: "ready",
                error: null,
            };
            pendingNotes = [];
            return;
        }

        if (isReportItemRow(row) && currentTicket) {
            const item = parseReportItem(row, tickets.length, currentTicket.items.length + 1);
            if (item) {
                currentTicket.items.push(item);
            }
            return;
        }
    });

    flushTicket();
    return tickets;
};

const parseFileToTickets = async (file: File): Promise<ParsedMctTicket[]> => {
    const extension = file.name.split(".").pop()?.toLowerCase();
    let rows: string[][] = [];

    if (extension === "csv") {
        rows = normalizeRows(parseCsvRows(await file.text()));
    } else if (extension === "xlsx" || extension === "xls") {
        const buffer = await file.arrayBuffer();
        const workbook = XLSX.read(buffer, { type: "array", cellDates: true });
        const [sheetName] = workbook.SheetNames;
        if (!sheetName) {
            throw new Error(`${file.name}: no sheets found in the workbook.`);
        }
        const sheet = workbook.Sheets[sheetName];
        if (!sheet) {
            throw new Error(`${file.name}: unable to read the first sheet.`);
        }
        const sheetRows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", raw: false }) as Array<
            Array<string | number | null>
        >;
        rows = normalizeRows(sheetRows);
    } else {
        throw new Error(`${file.name}: unsupported file format.`);
    }

    const reportTickets = parseMultiTicketReport(rows, file);
    if (reportTickets.length > 0) {
        return reportTickets;
    }

    const { header, items } = parseRowsToTicket(rows);
    const ticketId = createTicketId(file);
    return [
        {
            id: ticketId,
            fileName: file.name,
            header,
            items: items.map((item) => ({ ...item, id: `${ticketId}-${item.id}` })),
            status: "ready",
            error: null,
        },
    ];
};

const getAvailabilityKey = (ticketId: string, itemId: string) => `${ticketId}:${itemId}`;

const IssueReturnCreatePage = () => {
    const [file, setFile] = useState<File | null>(null);
    const [uploadedFiles, setUploadedFiles] = useState<File[]>([]);
    const [parsedTickets, setParsedTickets] = useState<ParsedMctTicket[]>([]);
    const [expandedTicketIds, setExpandedTicketIds] = useState<Set<string>>(new Set());
    const [ticketHeader, setTicketHeader] = useState<MaterialChargeTicketHeader | null>(null);
    const [ticketItems, setTicketItems] = useState<MaterialChargeTicketItem[]>([]);
    const [parseError, setParseError] = useState<string | null>(null);
    const [parseStatus, setParseStatus] = useState<string | null>(null);
    const [validationErrors, setValidationErrors] = useState<string[]>([]);
    const [missingInventoryItems, setMissingInventoryItems] = useState<MaterialChargeTicketItem[]>([]);
    const [confirmOpen, setConfirmOpen] = useState(false);
    const [errorDialogOpen, setErrorDialogOpen] = useState(false);
    const [isSubmitting, setIsSubmitting] = useState(false);
    const [availabilityMap, setAvailabilityMap] = useState<Record<string, AvailabilityInfo>>({});
    const [availabilityStatus, setAvailabilityStatus] = useState<"idle" | "loading" | "error">("idle");
    const { data: identity } = useGetIdentity<{ id?: string | number }>();
    const { open } = useNotification();
    const invalidate = useInvalidate();
    const go = useGo();

    useEffect(() => {
        setFile(uploadedFiles[0] ?? null);
    }, [uploadedFiles]);

    useEffect(() => {
        if (uploadedFiles.length === 0) {
            setTicketHeader(null);
            setTicketItems([]);
            setParsedTickets([]);
            setExpandedTicketIds(new Set());
            setParseError(null);
            setParseStatus(null);
            return;
        }

        let isActive = true;

        const parseFiles = async () => {
            setParseError(null);
            setParseStatus(`Parsing ${uploadedFiles.length} file${uploadedFiles.length === 1 ? "" : "s"}...`);

            try {
                const tickets = (await Promise.all(uploadedFiles.map((currentFile) => parseFileToTickets(currentFile)))).flat();
                if (!isActive) return;

                setParsedTickets(tickets);
                setExpandedTicketIds(new Set(tickets.slice(0, 1).map((ticket) => ticket.id)));
                setTicketHeader(tickets[0]?.header ?? null);
                setTicketItems(tickets[0]?.items ?? []);
                setParseStatus(
                    tickets.length > 1
                        ? `Parsed ${tickets.length} MCT document${tickets.length === 1 ? "" : "s"}.`
                        : `Parsed ${tickets[0]?.items.length ?? 0} item${tickets[0]?.items.length === 1 ? "" : "s"}.`
                );
            } catch (error) {
                if (!isActive) return;
                const message = error instanceof Error ? error.message : "Unable to parse file.";
                setParseError(message);
                setTicketHeader(null);
                setTicketItems([]);
                setParsedTickets([]);
                setExpandedTicketIds(new Set());
                setParseStatus(null);
            }
        };

        void parseFiles();

        return () => {
            isActive = false;
        };
    }, [uploadedFiles]);

    const summaryHeader = useMemo(() => ticketHeader ?? EMPTY_HEADER, [ticketHeader]);
    const totalQty = useMemo(() => sumNumbers(ticketItems.map((item) => item.qty)), [ticketItems]);
    const totalCost = useMemo(() => sumNumbers(ticketItems.map((item) => item.total_cost)), [ticketItems]);

    const handleDeductFromChange = (itemId: string, value: "ending_qty" | "buffer_stock") => {
        setTicketItems((prev) =>
            prev.map((item) =>
                item.id === itemId ? { ...item, deduct_from: value } : item
            )
        );
    };

    useEffect(() => {
        const availabilityItems = parsedTickets.length > 1
            ? parsedTickets.flatMap((ticket) => ticket.items.map((item) => ({ ticketId: ticket.id, item })))
            : ticketItems.map((item) => ({ ticketId: "", item }));

        const uniqueCodes = Array.from(
            new Set(
                availabilityItems
                    .map(({ item }) => (item.item_code ? normalizeItemCode(item.item_code) : ""))
                    .filter(Boolean)
            )
        );

        if (uniqueCodes.length === 0) {
            setAvailabilityMap({});
            setAvailabilityStatus("idle");
            return;
        }

        let isActive = true;

        const fetchAvailability = async () => {
            setAvailabilityStatus("loading");
            const { data: serverTimestamp, error: timestampError } =
                await supabaseClient.rpc("get_server_timestamp");

            if (timestampError || !serverTimestamp) {
                if (isActive) {
                    setAvailabilityMap({});
                    setAvailabilityStatus("error");
                }
                return;
            }

            const serverDate = new Date(serverTimestamp);
            const month = serverDate.getMonth() + 1;
            const year = serverDate.getFullYear();

            const { data: itemsData } = await supabaseClient
                .from("items")
                .select("id,item_code")
                .in("item_code", uniqueCodes);

            const itemIdByCode = new Map<string, string>();
            (itemsData ?? []).forEach((row) => {
                if (row.item_code) {
                    itemIdByCode.set(normalizeItemCode(row.item_code), row.id);
                }
            });

            const itemIds = Array.from(new Set(itemIdByCode.values()));
            const inventoryByItemId = new Map<string, { endingQty: number; bufferStock: number }>();

            if (itemIds.length > 0) {
                const { data: inventoryRows } = await supabaseClient
                    .from("inventory_records")
                    .select("item_id, ending_qty, buffer_stock")
                    .in("item_id", itemIds)
                    .eq("month", month)
                    .eq("year", year);

                (inventoryRows ?? []).forEach((row) => {
                    inventoryByItemId.set(row.item_id, {
                        endingQty: row.ending_qty ?? 0,
                        bufferStock: row.buffer_stock ?? 0,
                    });
                });
            }

            const nextAvailability: Record<string, AvailabilityInfo> = {};

            availabilityItems.forEach(({ ticketId, item }) => {
                const code = item.item_code ? normalizeItemCode(item.item_code) : "";
                if (!code) return;

                const availabilityKey = ticketId ? getAvailabilityKey(ticketId, item.id) : item.id;
                const itemId = itemIdByCode.get(code);
                if (!itemId) {
                    nextAvailability[availabilityKey] = { status: "missing" };
                    return;
                }

                const inventory = inventoryByItemId.get(itemId);
                if (!inventory) {
                    nextAvailability[availabilityKey] = { status: "missing" };
                    return;
                }

                const requestedQty = item.qty ?? 0;
                const deductFrom = item.deduct_from ?? "ending_qty";
                const effectiveAvailable =
                    deductFrom === "buffer_stock" ? inventory.bufferStock : inventory.endingQty;
                if (effectiveAvailable - requestedQty < 0) {
                    nextAvailability[availabilityKey] = {
                        status: "insufficient",
                        availableQty: effectiveAvailable,
                        bufferStock: inventory.bufferStock,
                        endingQty: inventory.endingQty,
                        deductFrom,
                    };
                    return;
                }

                nextAvailability[availabilityKey] = {
                    status: "in_stock",
                    availableQty: effectiveAvailable,
                    bufferStock: inventory.bufferStock,
                    endingQty: inventory.endingQty,
                    deductFrom,
                };
            });

            if (isActive) {
                setAvailabilityMap(nextAvailability);
                setAvailabilityStatus("idle");
            }
        };

        fetchAvailability();

        return () => {
            isActive = false;
        };
    }, [ticketItems, parsedTickets]);

    const validateItems = () => {
        const errors: string[] = [];
        if (ticketItems.length === 0) {
            errors.push("No item rows detected. Upload a file with item entries before saving.");
        }

        ticketItems.forEach((item, index) => {
            if (!item.item_code?.trim()) {
                errors.push(`Row ${index + 1}: Missing item code.`);
            }
            if (!item.deduct_from) {
                errors.push(`Row ${index + 1}: Deduct from is required.`);
            }
            if (item.qty == null || Number.isNaN(item.qty) || item.qty <= 0) {
                errors.push(`Row ${index + 1}: Quantity must be greater than 0.`);
            }
        });

        setValidationErrors(errors);
        setErrorDialogOpen(errors.length > 0);
        return errors.length === 0;
    };

    const buildHeaderPayload = (header: MaterialChargeTicketHeader = summaryHeader) => ({
        district: header.district || null,
        department: header.department || null,
        request_number: header.requestNumber || null,
        request_date: header.requestDate || null,
        requisitioner: header.requisitioner || null,
        release_date: header.releaseDate || null,
        mct_rel_number: header.mctRelNumber || null,
        wo_number: header.woNumber || null,
        jo_number: header.joNumber || null,
        so_number: header.soNumber || null,
        purpose: header.purpose || null,
        notes: header.notes || null,
    });

    const buildItemsPayload = (items: MaterialChargeTicketItem[] = ticketItems) =>
        items.map((item) => ({
            item_code: item.item_code?.trim() || null,
            particulars: item.particulars || null,
            unit: item.unit || null,
            unit_cost: item.unit_cost ?? null,
            qty: item.qty ?? null,
            total_cost: item.total_cost ?? null,
            c2: item.c2 ?? null,
            deduct_from: item.deduct_from ?? "ending_qty",
            remarks: item.notes || null,
        }));

    const parseMissingCodes = (message: string, prefix: string) => {
        if (!message.startsWith(prefix)) return [];
        const raw = message.slice(prefix.length).trim();
        if (!raw) return [];
        return raw.split(",").map((code) => code.trim()).filter(Boolean);
    };

    const getSaveError = (message: string) => {
        const duplicate = parseMissingCodes(message, "duplicate_mct:");
        if (duplicate.length > 0) return `Duplicate MCT/Rel # detected: ${duplicate.join(", ")}`;

        const missingItems = parseMissingCodes(message, "missing_item_codes:");
        if (missingItems.length > 0) return `Item code not found: ${missingItems.join(", ")}`;

        const insufficientInventory = parseMissingCodes(message, "insufficient_inventory:");
        if (insufficientInventory.length > 0) return `Insufficient inventory for item code: ${insufficientInventory.join(", ")}`;

        return message || "Unable to save MCT.";
    };

    const updateParsedTicketStatus = (
        ticketId: string,
        status: ParsedMctTicket["status"],
        error: string | null = null
    ) => {
        setParsedTickets((currentTickets) =>
            currentTickets.map((ticket) =>
                ticket.id === ticketId ? { ...ticket, status, error } : ticket
            )
        );
    };

    const updateParsedTicketDeductFrom = (
        ticketId: string,
        itemId: string,
        value: "ending_qty" | "buffer_stock"
    ) => {
        setParsedTickets((currentTickets) =>
            currentTickets.map((ticket) =>
                ticket.id === ticketId
                    ? {
                        ...ticket,
                        items: ticket.items.map((item) =>
                            item.id === itemId ? { ...item, deduct_from: value } : item
                        ),
                    }
                    : ticket
            )
        );
    };

    const toggleExpandedTicket = (ticketId: string) => {
        setExpandedTicketIds((currentIds) => {
            const nextIds = new Set(currentIds);
            if (nextIds.has(ticketId)) {
                nextIds.delete(ticketId);
            } else {
                nextIds.add(ticketId);
            }
            return nextIds;
        });
    };

    const validateParsedTickets = () => {
        const errors: string[] = [];

        if (parsedTickets.length === 0) {
            errors.push("No MCT documents detected. Upload files before saving.");
        }

        parsedTickets.forEach((ticket, ticketIndex) => {
            const label = ticket.header.mctRelNumber || `MCT ${ticketIndex + 1}`;
            if (ticket.items.length === 0) {
                errors.push(`${label}: no item rows detected.`);
            }

            ticket.items.forEach((item, itemIndex) => {
                if (!item.item_code?.trim()) {
                    errors.push(`${label}, row ${itemIndex + 1}: missing item code.`);
                }
                if (!item.deduct_from) {
                    errors.push(`${label}, row ${itemIndex + 1}: deduct from is required.`);
                }
                if (item.qty == null || Number.isNaN(item.qty) || item.qty <= 0) {
                    errors.push(`${label}, row ${itemIndex + 1}: quantity must be greater than 0.`);
                }
            });
        });

        setValidationErrors(errors);
        setErrorDialogOpen(errors.length > 0);
        return errors.length === 0;
    };

    const saveParsedTickets = async (ticketIds: string[], createMissingInventory: boolean) => {
        const savedIds: string[] = [];
        const missingInventoryIds: string[] = [];
        const errors: string[] = [];

        for (const ticket of parsedTickets.filter((currentTicket) => ticketIds.includes(currentTicket.id))) {
            updateParsedTicketStatus(ticket.id, "saving");

            const { error } = await supabaseClient.rpc("create_mct_transaction", {
                p_header: buildHeaderPayload(ticket.header),
                p_items: buildItemsPayload(ticket.items),
                p_create_missing_inventory: createMissingInventory,
                p_created_by: identity?.id ? String(identity.id) : null,
            });

            if (!error) {
                savedIds.push(ticket.id);
                updateParsedTicketStatus(ticket.id, "saved");
                continue;
            }

            const message = error.message ?? "Unable to save MCT.";
            const missingInventory = parseMissingCodes(message, "missing_inventory:");

            if (!createMissingInventory && missingInventory.length > 0) {
                missingInventoryIds.push(ticket.id);
                updateParsedTicketStatus(
                    ticket.id,
                    "error",
                    `Missing inventory record for item code: ${missingInventory.join(", ")}`
                );
                continue;
            }

            const displayError = getSaveError(message);
            errors.push(`${ticket.header.mctRelNumber || ticket.fileName}: ${displayError}`);
            updateParsedTicketStatus(ticket.id, "error", displayError);
        }

        return { savedIds, missingInventoryIds, errors };
    };

    const handleSaveParsedTickets = async () => {
        if (isSubmitting || !validateParsedTickets()) return;

        setIsSubmitting(true);
        setValidationErrors([]);
        setMissingInventoryItems([]);

        try {
            const ticketIds = parsedTickets
                .filter((ticket) => ticket.status === "ready" || ticket.status === "error")
                .map((ticket) => ticket.id);
            const result = await saveParsedTickets(ticketIds, false);

            if (result.missingInventoryIds.length > 0) {
                const missingItems = parsedTickets
                    .filter((ticket) => result.missingInventoryIds.includes(ticket.id))
                    .flatMap((ticket) => ticket.items);
                setMissingInventoryItems(missingItems);
                setConfirmOpen(true);
            }

            if (result.errors.length > 0) {
                setValidationErrors(result.errors);
                setErrorDialogOpen(true);
            }

            if (result.savedIds.length > 0) {
                invalidate({ resource: "mcts", invalidates: ["list"] });
                invalidate({ resource: "mct_items", invalidates: ["list"] });
            }

            if (result.savedIds.length > 0 && result.errors.length === 0 && result.missingInventoryIds.length === 0) {
                open?.({
                    type: "success",
                    message: "MCTs saved",
                    description: `${result.savedIds.length} MCT document${result.savedIds.length === 1 ? "" : "s"} saved.`,
                });
                go({ to: "/mct", type: "replace" });
            }
        } catch (error) {
            setValidationErrors([error instanceof Error ? error.message : "Unable to save MCTs."]);
            setErrorDialogOpen(true);
        } finally {
            setIsSubmitting(false);
        }
    };

    const handleAddMct = async () => {
        if (isSubmitting) return;
        setValidationErrors([]);
        setMissingInventoryItems([]);

        if (!validateItems()) {
            return;
        }

        setIsSubmitting(true);
        try {
            const { data, error } = await supabaseClient.rpc("create_mct_transaction", {
                p_header: buildHeaderPayload(),
                p_items: buildItemsPayload(),
                p_create_missing_inventory: false,
                p_created_by: identity?.id ? String(identity.id) : null,
            });

            if (error) {
                const message = error.message ?? "Unable to save MCT.";
                const duplicate = parseMissingCodes(message, "duplicate_mct:");
                if (duplicate.length > 0) {
                    const errors = [`Duplicate MCT/Rel # detected: ${duplicate.join(", ")}`];
                    setValidationErrors(errors);
                    setErrorDialogOpen(true);
                    return;
                }
                const missingItems = parseMissingCodes(message, "missing_item_codes:");
                if (missingItems.length > 0) {
                    const errors = missingItems.map((code) => `Item code not found: ${code}`);
                    setValidationErrors(errors);
                    setErrorDialogOpen(true);
                    return;
                }
                const insufficientInventory = parseMissingCodes(message, "insufficient_inventory:");
                if (insufficientInventory.length > 0) {
                    const errors = insufficientInventory.map(
                        (code) => `Insufficient inventory for item code: ${code}`
                    );
                    setValidationErrors(errors);
                    setErrorDialogOpen(true);
                    return;
                }
                const missingInventory = parseMissingCodes(message, "missing_inventory:");
                if (missingInventory.length > 0) {
                    setMissingInventoryItems(
                        ticketItems.filter((item) =>
                            missingInventory.includes(item.item_code.trim().toUpperCase())
                        )
                    );
                    setConfirmOpen(true);
                    return;
                }
                setValidationErrors([message]);
                setErrorDialogOpen(true);
                return;
            }

            open?.({
                type: "success",
                message: "MCT saved",
                description: "Material charge ticket saved.",
            });

            invalidate({ resource: "mcts", invalidates: ["list"] });
            invalidate({ resource: "mct_items", invalidates: ["list"] });

            go({ to: "/mct", type: "replace" });
        } catch (error) {
            const description = error instanceof Error ? error.message : "Unable to save MCT.";
            setValidationErrors([description]);
            setErrorDialogOpen(true);
        } finally {
            setIsSubmitting(false);
        }
    };

    const handleConfirmCreateInventory = async () => {
        if (isSubmitting) {
            setConfirmOpen(false);
            return;
        }

        setIsSubmitting(true);
        try {
            if (parsedTickets.length > 1) {
                const ticketIds = parsedTickets
                    .filter((ticket) => ticket.status === "error")
                    .map((ticket) => ticket.id);
                const result = await saveParsedTickets(ticketIds, true);

                if (result.errors.length > 0) {
                    setValidationErrors(result.errors);
                    setErrorDialogOpen(true);
                    return;
                }

                open?.({
                    type: "success",
                    message: "MCTs saved",
                    description: "MCT documents saved.",
                });

                invalidate({ resource: "mcts", invalidates: ["list"] });
                invalidate({ resource: "mct_items", invalidates: ["list"] });
                go({ to: "/mct", type: "replace" });
                return;
            }

            const { error } = await supabaseClient.rpc("create_mct_transaction", {
                p_header: buildHeaderPayload(),
                p_items: buildItemsPayload(),
                p_create_missing_inventory: true,
                p_created_by: identity?.id ? String(identity.id) : null,
            });

            if (error) {
                const message = error.message ?? "Unable to save MCT.";
                const duplicate = parseMissingCodes(message, "duplicate_mct:");
                if (duplicate.length > 0) {
                    const errors = [`Duplicate MCT/Rel # detected: ${duplicate.join(", ")}`];
                    setValidationErrors(errors);
                    setErrorDialogOpen(true);
                    return;
                }
                setValidationErrors([message]);
                setErrorDialogOpen(true);
                return;
            }

            open?.({
                type: "success",
                message: "MCT saved",
                description: "Material charge ticket saved.",
            });

            invalidate({ resource: "mcts", invalidates: ["list"] });
            invalidate({ resource: "mct_items", invalidates: ["list"] });

            go({ to: "/mct", type: "replace" });
        } catch (error) {
            const description = error instanceof Error ? error.message : "Unable to save MCT.";
            setValidationErrors([description]);
            setErrorDialogOpen(true);
        } finally {
            setIsSubmitting(false);
            setConfirmOpen(false);
        }
    };

    if (parsedTickets.length > 1) {
        return (
            <CreateView className="item-view">
                <CreateViewHeader title="MCT" />
                <div className="my-4 flex items-center">
                    <Card className="w-full max-w-7xl mx-auto item-form-card gap-0 overflow-hidden border-border/80 shadow-sm py-0">
                        <CardHeader className="border-b pt-6">
                            <CardTitle>Material Charge Ticket</CardTitle>
                            <CardDescription>
                                Upload one or more Excel/CSV files to extract Material Charge Ticket details.
                            </CardDescription>
                        </CardHeader>
                        <CardContent className="pt-5 space-y-5">
                            <MultiUploadWidget value={uploadedFiles} onFilesChange={setUploadedFiles} disabled={isSubmitting} />
                            {parseStatus ? <p className="text-sm text-muted-foreground">{parseStatus}</p> : null}
                            {parseError ? <p className="text-sm text-destructive">{parseError}</p> : null}

                            <div className="grid gap-2 mb-5">
                                {parsedTickets.map((ticket, ticketIndex) => {
                                    const header = ticket.header;
                                    const isExpanded = expandedTicketIds.has(ticket.id);
                                    const ticketTotalQty = sumNumbers(ticket.items.map((item) => item.qty));
                                    const ticketTotalCost = sumNumbers(ticket.items.map((item) => item.total_cost));

                                    return (
                                        <div key={ticket.id} className="rounded-lg border bg-muted/10">
                                            <button
                                                type="button"
                                                onClick={() => toggleExpandedTicket(ticket.id)}
                                                className="flex w-full items-center justify-between gap-3 px-3 py-3 text-left"
                                            >
                                                <div className="flex min-w-0 items-center gap-2">
                                                    {isExpanded ? (
                                                        <ChevronDown className="h-4 w-4 shrink-0" />
                                                    ) : (
                                                        <ChevronRight className="h-4 w-4 shrink-0" />
                                                    )}
                                                    <div className="min-w-0">
                                                        <p className="truncate text-sm font-semibold">
                                                            {header.mctRelNumber || `MCT ${ticketIndex + 1}`}
                                                        </p>
                                                        <p className="truncate text-xs text-muted-foreground">
                                                            {ticket.fileName} · {header.requisitioner || "No requisitioner"}
                                                        </p>
                                                    </div>
                                                </div>
                                                <div className="flex shrink-0 items-center gap-3 text-xs text-muted-foreground">
                                                    <span>{ticket.items.length} item{ticket.items.length === 1 ? "" : "s"}</span>
                                                    <span>{formatDecimal(ticketTotalCost)}</span>
                                                    {ticket.status === "saving" ? (
                                                        <Badge variant="secondary">Saving</Badge>
                                                    ) : ticket.status === "saved" ? (
                                                        <Badge className="bg-emerald-600 hover:bg-emerald-600">Saved</Badge>
                                                    ) : ticket.status === "error" ? (
                                                        <Badge variant="destructive">Error</Badge>
                                                    ) : (
                                                        <Badge variant="outline">Ready</Badge>
                                                    )}
                                                </div>
                                            </button>

                                            {isExpanded ? (
                                                <div className="border-t p-3 space-y-4">
                                                    {ticket.error ? (
                                                        <p className="text-sm text-destructive">{ticket.error}</p>
                                                    ) : null}
                                                    <p className="text-xs font-semibold uppercase tracking-wide text-muted-foreground">
                                                        Ticket Details
                                                    </p>
                                                    <div className="grid gap-4 lg:grid-cols-[1.1fr_0.9fr]">
                                                        <div className="grid gap-3">
                                                            <div className="flex items-center justify-between gap-3 rounded-md border bg-background px-3 py-2 text-sm">
                                                                <span className="text-muted-foreground">District</span>
                                                                <span className="font-medium">{header.district || "-"}</span>
                                                            </div>
                                                            <div className="flex items-center justify-between gap-3 rounded-md border bg-background px-3 py-2 text-sm">
                                                                <span className="text-muted-foreground">Department</span>
                                                                <span className="font-medium">{header.department || "-"}</span>
                                                            </div>
                                                            <div className="flex items-center justify-between gap-3 rounded-md border bg-background px-3 py-2 text-sm">
                                                                <span className="text-muted-foreground">Request #</span>
                                                                <span className="font-medium">{header.requestNumber || "-"}</span>
                                                            </div>
                                                            <div className="flex items-center justify-between gap-3 rounded-md border bg-background px-3 py-2 text-sm">
                                                                <span className="text-muted-foreground">Req. Date</span>
                                                                <span className="font-medium">{header.requestDate || "-"}</span>
                                                            </div>
                                                            <div className="flex items-center justify-between gap-3 rounded-md border bg-background px-3 py-2 text-sm">
                                                                <span className="text-muted-foreground">Requisitioner</span>
                                                                <span className="font-medium">{header.requisitioner || "-"}</span>
                                                            </div>
                                                        </div>

                                                        <div className="grid gap-3">
                                                            <div className="flex items-center justify-between gap-3 rounded-md border bg-background px-3 py-2 text-sm">
                                                                <span className="text-muted-foreground">Rel. Date</span>
                                                                <span className="font-medium">{header.releaseDate || "-"}</span>
                                                            </div>
                                                            <div className="flex items-center justify-between gap-3 rounded-md border bg-background px-3 py-2 text-sm">
                                                                <span className="text-muted-foreground">MCT/Rel #</span>
                                                                <span className="font-medium">{header.mctRelNumber || "-"}</span>
                                                            </div>
                                                            <div className="flex items-center justify-between gap-3 rounded-md border bg-background px-3 py-2 text-sm">
                                                                <span className="text-muted-foreground">WO #</span>
                                                                <span className="font-medium">{header.woNumber || "-"}</span>
                                                            </div>
                                                            <div className="flex items-center justify-between gap-3 rounded-md border bg-background px-3 py-2 text-sm">
                                                                <span className="text-muted-foreground">JO #</span>
                                                                <span className="font-medium">{header.joNumber || "-"}</span>
                                                            </div>
                                                            <div className="flex items-center justify-between gap-3 rounded-md border bg-background px-3 py-2 text-sm">
                                                                <span className="text-muted-foreground">SO #</span>
                                                                <span className="font-medium">{header.soNumber || "-"}</span>
                                                            </div>
                                                        </div>
                                                    </div>

                                                    <p className="text-xs font-semibold uppercase tracking-wide text-muted-foreground">
                                                        Ticket Items
                                                    </p>
                                                    <div className="overflow-x-auto rounded-md border bg-background">
                                                        <Table>
                                                            <TableHeader>
                                                                <TableRow>
                                                                    <TableHead className="w-12 text-center">#</TableHead>
                                                                    <TableHead>Item Code</TableHead>
                                                                    <TableHead>Particulars</TableHead>
                                                                    <TableHead>Unit</TableHead>
                                                                    <TableHead className="text-right">Unit Cost</TableHead>
                                                                    <TableHead className="text-right">Qty</TableHead>
                                                                    <TableHead className="text-right">Total Cost</TableHead>
                                                                    <TableHead className="text-right">C2</TableHead>
                                                                    <TableHead>Remarks</TableHead>
                                                                    <TableHead>Deduct From</TableHead>
                                                                </TableRow>
                                                            </TableHeader>
                                                            <TableBody>
                                                                {ticket.items.length === 0 ? (
                                                                    <TableRow>
                                                                        <TableCell colSpan={10} className="py-8 text-center text-sm text-muted-foreground">
                                                                            No item rows detected yet.
                                                                        </TableCell>
                                                                    </TableRow>
                                                                ) : (
                                                                    ticket.items.map((item, index) => {
                                                                        const availability = availabilityMap[getAvailabilityKey(ticket.id, item.id)];
                                                                        const status = availability?.status;
                                                                        const label = !availability
                                                                            ? availabilityStatus === "loading"
                                                                                ? "Checking..."
                                                                                : "No record"
                                                                            : status === "in_stock"
                                                                                ? "In stock"
                                                                                : status === "insufficient"
                                                                                    ? "Insufficient"
                                                                                    : "No record";
                                                                        const colorClass =
                                                                            status === "in_stock"
                                                                                ? "text-emerald-600"
                                                                                : status === "insufficient"
                                                                                    ? "text-destructive"
                                                                                    : "text-amber-600";

                                                                        return (
                                                                            <TableRow key={item.id}>
                                                                                <TableCell className="text-center text-muted-foreground">{index + 1}</TableCell>
                                                                                <TableCell>
                                                                                    <Tooltip>
                                                                                        <TooltipTrigger asChild>
                                                                                            <span className={`font-medium ${colorClass}`}>{item.item_code || "-"}</span>
                                                                                        </TooltipTrigger>
                                                                                        <TooltipContent side="top" align="start">
                                                                                            <div className="grid gap-0.5">
                                                                                                <span className="text-xs font-semibold">{label}</span>
                                                                                                {availability?.endingQty != null ? (
                                                                                                    <span className="text-xs text-primary-foreground/80">
                                                                                                        Available (Ending): {availability.endingQty}
                                                                                                    </span>
                                                                                                ) : null}
                                                                                                {availability?.bufferStock != null ? (
                                                                                                    <span className="text-xs text-primary-foreground/80">
                                                                                                        Buffer stock: {availability.bufferStock}
                                                                                                    </span>
                                                                                                ) : null}
                                                                                            </div>
                                                                                        </TooltipContent>
                                                                                    </Tooltip>
                                                                                </TableCell>
                                                                                <TableCell className="min-w-[220px] whitespace-normal">{item.particulars || "-"}</TableCell>
                                                                                <TableCell>{item.unit || "-"}</TableCell>
                                                                                <TableCell className="text-right">{formatDecimal(item.unit_cost)}</TableCell>
                                                                                <TableCell className="text-right">{item.qty ?? "-"}</TableCell>
                                                                                <TableCell className="text-right">{formatDecimal(item.total_cost)}</TableCell>
                                                                                <TableCell className="text-right">{formatC2(item.c2)}</TableCell>
                                                                                <TableCell className="min-w-[160px] whitespace-normal">{item.notes || "-"}</TableCell>
                                                                                <TableCell className="whitespace-nowrap py-1">
                                                                                    <Select
                                                                                        value={item.deduct_from}
                                                                                        disabled={isSubmitting || ticket.status === "saved"}
                                                                                        onValueChange={(value) =>
                                                                                            updateParsedTicketDeductFrom(
                                                                                                ticket.id,
                                                                                                item.id,
                                                                                                value as "ending_qty" | "buffer_stock"
                                                                                            )
                                                                                        }
                                                                                    >
                                                                                        <SelectTrigger className="h-8 w-30 px-2">
                                                                                            <SelectValue placeholder="Deduct from" />
                                                                                        </SelectTrigger>
                                                                                        <SelectContent>
                                                                                            <SelectItem value="ending_qty">Ending Qty</SelectItem>
                                                                                            <SelectItem value="buffer_stock">Buffer Stock</SelectItem>
                                                                                        </SelectContent>
                                                                                    </Select>
                                                                                </TableCell>
                                                                            </TableRow>
                                                                        );
                                                                    })
                                                                )}
                                                                {ticket.items.length > 0 ? (
                                                                    <TableRow>
                                                                        <TableCell className="text-center text-sm font-semibold text-muted-foreground">-</TableCell>
                                                                        <TableCell />
                                                                        <TableCell />
                                                                        <TableCell />
                                                                        <TableCell className="text-right text-sm font-semibold">Total:</TableCell>
                                                                        <TableCell className="text-right text-sm font-semibold">
                                                                            {Number.isFinite(ticketTotalQty) ? ticketTotalQty : "-"}
                                                                        </TableCell>
                                                                        <TableCell className="text-right text-sm font-semibold">
                                                                            {formatDecimal(ticketTotalCost)}
                                                                        </TableCell>
                                                                        <TableCell />
                                                                        <TableCell />
                                                                        <TableCell />
                                                                    </TableRow>
                                                                ) : null}
                                                            </TableBody>
                                                        </Table>
                                                    </div>

                                                    <div className="grid gap-4">
                                                        <div className="grid gap-1.5">
                                                            <p className="text-xs font-semibold uppercase tracking-wide text-muted-foreground">Purpose</p>
                                                            <Textarea value={header.purpose} readOnly className="min-h-24 bg-background" />
                                                        </div>
                                                        <div className="grid gap-1.5">
                                                            <p className="text-xs font-semibold uppercase tracking-wide text-muted-foreground">Notes / SR #</p>
                                                            <Textarea value={header.notes} readOnly className="h-24 resize-y overflow-auto bg-background" />
                                                        </div>
                                                    </div>
                                                </div>
                                            ) : null}
                                        </div>
                                    );
                                })}
                            </div>
                        </CardContent>
                        <CardFooter className="border-t px-0 py-4 !pt-4">
                            <div className="flex w-full items-center justify-between px-6">
                                <p className="text-xs text-muted-foreground">
                                    {isSubmitting ? "Saving MCTs..." : "Review parsed values before saving."}
                                </p>
                                <div className="flex justify-end gap-2">
                                    <Button
                                        type="button"
                                        variant="outline"
                                        onClick={() => go({ to: "/mct", type: "replace" })}
                                        disabled={isSubmitting}
                                    >
                                        Cancel
                                    </Button>
                                    <Button type="button" onClick={handleSaveParsedTickets} disabled={isSubmitting}>
                                        {isSubmitting ? "Saving..." : "Add MCTs"}
                                    </Button>
                                </div>
                            </div>
                        </CardFooter>
                    </Card>
                </div>
                <AlertDialog open={confirmOpen} onOpenChange={setConfirmOpen}>
                    <AlertDialogContent className="sm:max-w-xl overflow-hidden p-0 border-border/80 shadow-sm">
                        <AlertDialogHeader className="border-b px-6 py-5">
                            <AlertDialogTitle className="text-2xl">Missing inventory records</AlertDialogTitle>
                            <AlertDialogDescription>
                                Some items do not have an inventory record for the current month. Add records now?
                            </AlertDialogDescription>
                        </AlertDialogHeader>
                        <div className="grid gap-3 px-6 py-6 text-sm">
                            {missingInventoryItems.map((item) => (
                                <div key={item.id} className="flex items-center justify-between gap-4">
                                    <span className="font-medium">{item.item_code}</span>
                                    <span className="text-muted-foreground">{item.particulars || "-"}</span>
                                </div>
                            ))}
                        </div>
                        <AlertDialogFooter className="border-t px-6 py-4 sm:justify-end">
                            <AlertDialogCancel disabled={isSubmitting}>Cancel</AlertDialogCancel>
                            <AlertDialogAction onClick={handleConfirmCreateInventory} disabled={isSubmitting}>
                                Add Records &amp; Continue
                            </AlertDialogAction>
                        </AlertDialogFooter>
                    </AlertDialogContent>
                </AlertDialog>
                <Dialog open={errorDialogOpen} onOpenChange={setErrorDialogOpen}>
                    <DialogContent className="sm:max-w-md">
                        <DialogHeader>
                            <DialogTitle>Unable to save MCT</DialogTitle>
                            <DialogDescription>
                                {validationErrors.length > 0 ? (
                                    <ul className="list-disc pl-4 space-y-1 text-destructive">
                                        {validationErrors.map((error) => (
                                            <li key={error}>{error}</li>
                                        ))}
                                    </ul>
                                ) : (
                                    <span className="text-destructive">Please review the errors and try again.</span>
                                )}
                            </DialogDescription>
                        </DialogHeader>
                        <DialogFooter>
                            <Button type="button" onClick={() => setErrorDialogOpen(false)}>
                                Okay
                            </Button>
                        </DialogFooter>
                    </DialogContent>
                </Dialog>
            </CreateView>
        );
    }

    return (
        <CreateView className="item-view">
            <CreateViewHeader title="MCT" />
            <div className="my-4 flex items-center">
                <Card className="w-full max-w-7xl mx-auto item-form-card gap-0 overflow-hidden border-border/80 shadow-sm py-0">
                    <CardHeader className="border-b pt-6">
                        <CardTitle>Material Charge Ticket</CardTitle>
                        <CardDescription>
                            Upload an Excel or CSV file to extract Material Charge Ticket details.
                        </CardDescription>
                    </CardHeader>
                    <CardContent className="pt-5 space-y-5">
                        <MultiUploadWidget value={uploadedFiles} onFilesChange={setUploadedFiles} />
                        {parseStatus ? <p className="text-sm text-muted-foreground">{parseStatus}</p> : null}
                        {parseError ? <p className="text-sm text-destructive">{parseError}</p> : null}
                        <div className="rounded-lg border bg-muted/10 p-3 space-y-4 mb-4">
                            <p className="text-xs font-semibold uppercase tracking-wide text-muted-foreground">
                                Ticket Details
                            </p>
                            <div className="grid gap-4 lg:grid-cols-[1.1fr_0.9fr]">
                                <div className="grid gap-3">
                                    <div className="flex items-center justify-between gap-3 rounded-md border bg-background px-3 py-2 text-sm">
                                        <span className="text-muted-foreground">District</span>
                                        <span className="font-medium">{summaryHeader.district || "-"}</span>
                                    </div>
                                    <div className="flex items-center justify-between gap-3 rounded-md border bg-background px-3 py-2 text-sm">
                                        <span className="text-muted-foreground">Department</span>
                                        <span className="font-medium">{summaryHeader.department || "-"}</span>
                                    </div>
                                    <div className="flex items-center justify-between gap-3 rounded-md border bg-background px-3 py-2 text-sm">
                                        <span className="text-muted-foreground">Request #</span>
                                        <span className="font-medium">{summaryHeader.requestNumber || "-"}</span>
                                    </div>
                                    <div className="flex items-center justify-between gap-3 rounded-md border bg-background px-3 py-2 text-sm">
                                        <span className="text-muted-foreground">Req. Date</span>
                                        <span className="font-medium">{summaryHeader.requestDate || "-"}</span>
                                    </div>
                                    <div className="flex items-center justify-between gap-3 rounded-md border bg-background px-3 py-2 text-sm">
                                        <span className="text-muted-foreground">Requisitioner</span>
                                        <span className="font-medium">{summaryHeader.requisitioner || "-"}</span>
                                    </div>
                                </div>

                                <div className="grid gap-3">
                                    <div className="flex items-center justify-between gap-3 rounded-md border bg-background px-3 py-2 text-sm">
                                        <span className="text-muted-foreground">Rel. Date</span>
                                        <span className="font-medium">{summaryHeader.releaseDate || "-"}</span>
                                    </div>
                                    <div className="flex items-center justify-between gap-3 rounded-md border bg-background px-3 py-2 text-sm">
                                        <span className="text-muted-foreground">MCT/Rel #</span>
                                        <span className="font-medium">{summaryHeader.mctRelNumber || "-"}</span>
                                    </div>
                                    <div className="flex items-center justify-between gap-3 rounded-md border bg-background px-3 py-2 text-sm">
                                        <span className="text-muted-foreground">WO #</span>
                                        <span className="font-medium">{summaryHeader.woNumber || "-"}</span>
                                    </div>
                                    <div className="flex items-center justify-between gap-3 rounded-md border bg-background px-3 py-2 text-sm">
                                        <span className="text-muted-foreground">JO #</span>
                                        <span className="font-medium">{summaryHeader.joNumber || "-"}</span>
                                    </div>
                                    <div className="flex items-center justify-between gap-3 rounded-md border bg-background px-3 py-2 text-sm">
                                        <span className="text-muted-foreground">SO #</span>
                                        <span className="font-medium">{summaryHeader.soNumber || "-"}</span>
                                    </div>
                                </div>
                            </div>
                            <p className="text-xs font-semibold uppercase tracking-wide text-muted-foreground">
                                Ticket Items
                            </p>
                            <div className="overflow-x-auto rounded-md border bg-background">
                                <Table>
                                    <TableHeader>
                                        <TableRow>
                                            <TableHead className="w-12 text-center">#</TableHead>
                                            <TableHead>Item Code</TableHead>
                                            <TableHead>Particulars</TableHead>
                                            <TableHead>Unit</TableHead>
                                            <TableHead className="text-right">Unit Cost</TableHead>
                                            <TableHead className="text-right">Qty</TableHead>
                                            <TableHead className="text-right">Total Cost</TableHead>
                                            <TableHead className="text-right">C2</TableHead>
                                            <TableHead>Remarks</TableHead>
                                            <TableHead>Deduct From</TableHead>
                                        </TableRow>
                                    </TableHeader>
                                    <TableBody>
                                        {ticketItems.length === 0 ? (
                                            <TableRow>
                                                <TableCell colSpan={10} className="py-8 text-center text-sm text-muted-foreground">
                                                    No item rows detected yet.
                                                </TableCell>
                                            </TableRow>
                                        ) : (
                                            ticketItems.map((item, index) => (
                                                <TableRow key={item.id}>
                                                    <TableCell className="text-center text-muted-foreground">{index + 1}</TableCell>
                                                    <TableCell>
                                                        {(() => {
                                                            const availability = availabilityMap[item.id];
                                                            const status = availability?.status;
                                                            const label = !availability
                                                                ? availabilityStatus === "loading"
                                                                    ? "Checking..."
                                                                    : "No record"
                                                                : availability.status === "in_stock"
                                                                    ? "In stock"
                                                                    : availability.status === "insufficient"
                                                                        ? "Insufficient"
                                                                        : "No record";
                                                            const availableText =
                                                                availability?.endingQty != null
                                                                    ? `Available (Ending): ${availability.endingQty}`
                                                                    : null;
                                                            const bufferText =
                                                                availability?.bufferStock != null
                                                                    ? `Buffer stock: ${availability.bufferStock}`
                                                                    : null;
                                                            const colorClass =
                                                                status === "in_stock"
                                                                    ? "text-emerald-600"
                                                                    : status === "insufficient"
                                                                        ? "text-destructive"
                                                                        : "text-amber-600";

                                                            return (
                                                                <Tooltip>
                                                                    <TooltipTrigger asChild>
                                                                        <span className={`font-medium ${colorClass}`}>
                                                                            {item.item_code || "-"}
                                                                        </span>
                                                                    </TooltipTrigger>
                                                                    <TooltipContent side="top" align="start">
                                                                        <div className="grid gap-0.5">
                                                                            <span className="text-xs font-semibold">{label}</span>
                                                                            {availableText ? (
                                                                                <span className="text-xs text-primary-foreground/80">
                                                                                    {availableText}
                                                                                </span>
                                                                            ) : null}
                                                                            {bufferText ? (
                                                                                <span className="text-xs text-primary-foreground/80">
                                                                                    {bufferText}
                                                                                </span>
                                                                            ) : null}
                                                                        </div>
                                                                    </TooltipContent>
                                                                </Tooltip>
                                                            );
                                                        })()}
                                                    </TableCell>
                                                    <TableCell className="min-w-[220px] whitespace-normal">{item.particulars || "-"}</TableCell>
                                                    <TableCell>{item.unit || "-"}</TableCell>
                                                    <TableCell className="text-right">{formatDecimal(item.unit_cost)}</TableCell>
                                                    <TableCell className="text-right">{item.qty ?? "-"}</TableCell>
                                                    <TableCell className="text-right">{formatDecimal(item.total_cost)}</TableCell>
                                                    <TableCell className="text-right">{formatC2(item.c2)}</TableCell>
                                                    <TableCell className="min-w-[160px] whitespace-normal">{item.notes || "-"}</TableCell>
                                                    <TableCell className="whitespace-nowrap py-1">
                                                        <Select
                                                            value={item.deduct_from}
                                                            onValueChange={(value) =>
                                                                handleDeductFromChange(
                                                                    item.id,
                                                                    value as "ending_qty" | "buffer_stock"
                                                                )
                                                            }
                                                        >
                                                            <SelectTrigger className="h-8 w-30 px-2">
                                                                <SelectValue placeholder="Deduct from" />
                                                            </SelectTrigger>
                                                            <SelectContent>
                                                                <SelectItem value="ending_qty">Ending Qty</SelectItem>
                                                                <SelectItem value="buffer_stock">Buffer Stock</SelectItem>
                                                            </SelectContent>
                                                        </Select>
                                                    </TableCell>
                                                </TableRow>
                                            ))
                                        )}
                                        {ticketItems.length > 0 ? (
                                            <TableRow>
                                                <TableCell className="text-center text-sm font-semibold text-muted-foreground">-</TableCell>
                                                <TableCell />
                                                <TableCell />
                                                <TableCell />
                                                <TableCell className="text-right text-sm font-semibold">
                                                    Total:
                                                </TableCell>
                                                <TableCell className="text-right text-sm font-semibold">
                                                    {Number.isFinite(totalQty) ? totalQty : "-"}
                                                </TableCell>
                                                <TableCell className="text-right text-sm font-semibold">
                                                    {formatDecimal(totalCost)}
                                                </TableCell>
                                                <TableCell />
                                                <TableCell />
                                                <TableCell />
                                            </TableRow>
                                        ) : null}
                                    </TableBody>
                                </Table>
                            </div>
                            <div className="grid gap-4">
                                <div className="grid gap-1.5">
                                    <p className="text-xs font-semibold uppercase tracking-wide text-muted-foreground">Purpose</p>
                                    <Textarea value={summaryHeader.purpose} readOnly className="min-h-24 bg-background" />
                                </div>
                                <div className="grid gap-1.5">
                                    <p className="text-xs font-semibold uppercase tracking-wide text-muted-foreground">Notes / SR #</p>
                                    <Textarea value={summaryHeader.notes} readOnly className="h-24 resize-y overflow-auto bg-background" />
                                </div>
                            </div>
                        </div>
                    </CardContent>
                    <CardFooter className="border-t px-0 py-4 !pt-4">
                        <div className="flex w-full items-center justify-between px-6">
                            <p className="text-xs text-muted-foreground">
                                {isSubmitting ? "Saving MCT..." : "Review parsed values before saving."}
                            </p>
                            <div className="flex justify-end gap-2">
                                <Button
                                    type="button"
                                    variant="outline"
                                    onClick={() => go({ to: "/mct", type: "replace" })}
                                    disabled={isSubmitting}
                                >
                                    Cancel
                                </Button>
                                <Button type="button" onClick={handleAddMct} disabled={isSubmitting}>
                                    {isSubmitting ? "Saving..." : "Add MCT"}
                                </Button>
                            </div>
                        </div>
                    </CardFooter>
                </Card>
            </div>
            <AlertDialog open={confirmOpen} onOpenChange={setConfirmOpen}>
                <AlertDialogContent className="sm:max-w-xl overflow-hidden p-0 border-border/80 shadow-sm">
                    <AlertDialogHeader className="border-b px-6 py-5">
                        <AlertDialogTitle className="text-2xl">Missing inventory records</AlertDialogTitle>
                        <AlertDialogDescription>
                            Some items do not have an inventory record for the current month. Add records now?
                        </AlertDialogDescription>
                    </AlertDialogHeader>
                    <div className="grid gap-3 px-6 py-6 text-sm">
                        {missingInventoryItems.map((item) => (
                            <div key={item.id} className="flex items-center justify-between gap-4">
                                <span className="font-medium">{item.item_code}</span>
                                <span className="text-muted-foreground">{item.particulars || "-"}</span>
                            </div>
                        ))}
                    </div>
                    <AlertDialogFooter className="border-t px-6 py-4 sm:justify-end">
                        <AlertDialogCancel disabled={isSubmitting}>Cancel</AlertDialogCancel>
                        <AlertDialogAction onClick={handleConfirmCreateInventory} disabled={isSubmitting}>
                            Add Records &amp; Continue
                        </AlertDialogAction>
                    </AlertDialogFooter>
                </AlertDialogContent>
            </AlertDialog>
            <Dialog open={errorDialogOpen} onOpenChange={setErrorDialogOpen}>
                <DialogContent className="sm:max-w-md">
                    <DialogHeader>
                        <DialogTitle>Unable to save MCT</DialogTitle>
                        <DialogDescription>
                            {validationErrors.length > 0 ? (
                                <ul className="list-disc pl-4 space-y-1 text-destructive">
                                    {validationErrors.map((error) => (
                                        <li key={error}>{error}</li>
                                    ))}
                                </ul>
                            ) : (
                                <span className="text-destructive">Please review the errors and try again.</span>
                            )}
                        </DialogDescription>
                    </DialogHeader>
                    <DialogFooter>
                        <Button type="button" onClick={() => setErrorDialogOpen(false)}>
                            Okay
                        </Button>
                    </DialogFooter>
                </DialogContent>
            </Dialog>
        </CreateView>
    );
};

export default IssueReturnCreatePage;

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
import BatchUploadMctPage from "./batch-upload";

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

const IssueReturnCreatePage = () => {
    const [file, setFile] = useState<File | null>(null);
    const [uploadedFiles, setUploadedFiles] = useState<File[]>([]);
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
        if (!file) {
            setTicketHeader(null);
            setTicketItems([]);
            setParseError(null);
            setParseStatus(null);
            return;
        }

        const parseFile = async () => {
            setParseError(null);
            setParseStatus("Parsing file...");

            try {
                const extension = file.name.split(".").pop()?.toLowerCase();
                if (extension === "csv") {
                    const text = await file.text();
                    const rows = parseCsvRows(text);
                    const normalizedRows = normalizeRows(rows);
                    const { header, items } = parseRowsToTicket(normalizedRows);
                    setTicketHeader(header);
                    setTicketItems(items);
                    setParseStatus(`Parsed ${items.length} item${items.length === 1 ? "" : "s"}.`);
                    return;
                }

                if (extension === "xlsx" || extension === "xls") {
                    const buffer = await file.arrayBuffer();
                    const workbook = XLSX.read(buffer, { type: "array", cellDates: true });
                    const [sheetName] = workbook.SheetNames;
                    if (!sheetName) {
                        throw new Error("No sheets found in the workbook.");
                    }
                    const sheet = workbook.Sheets[sheetName];
                    const sheetRows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", raw: false }) as Array<
                        Array<string | number | null>
                    >;
                    const normalizedRows = normalizeRows(sheetRows);
                    const { header, items } = parseRowsToTicket(normalizedRows);
                    setTicketHeader(header);
                    setTicketItems(items);
                    setParseStatus(`Parsed ${items.length} item${items.length === 1 ? "" : "s"}.`);
                    return;
                }

                throw new Error("Unsupported file format. Please upload a CSV or Excel file.");
            } catch (error) {
                const message = error instanceof Error ? error.message : "Unable to parse file.";
                setParseError(message);
                setTicketHeader(null);
                setTicketItems([]);
                setParseStatus(null);
            }
        };

        void parseFile();
    }, [file]);

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
        const uniqueCodes = Array.from(
            new Set(
                ticketItems
                    .map((item) => (item.item_code ? normalizeItemCode(item.item_code) : ""))
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

            ticketItems.forEach((item) => {
                const code = item.item_code ? normalizeItemCode(item.item_code) : "";
                if (!code) return;

                const itemId = itemIdByCode.get(code);
                if (!itemId) {
                    nextAvailability[item.id] = { status: "missing" };
                    return;
                }

                const inventory = inventoryByItemId.get(itemId);
                if (!inventory) {
                    nextAvailability[item.id] = { status: "missing" };
                    return;
                }

                const requestedQty = item.qty ?? 0;
                const deductFrom = item.deduct_from ?? "ending_qty";
                const effectiveAvailable =
                    deductFrom === "buffer_stock" ? inventory.bufferStock : inventory.endingQty;
                if (effectiveAvailable - requestedQty < 0) {
                    nextAvailability[item.id] = {
                        status: "insufficient",
                        availableQty: effectiveAvailable,
                        bufferStock: inventory.bufferStock,
                        endingQty: inventory.endingQty,
                        deductFrom,
                    };
                    return;
                }

                nextAvailability[item.id] = {
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
    }, [ticketItems]);

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

    const buildHeaderPayload = () => ({
        district: summaryHeader.district || null,
        department: summaryHeader.department || null,
        request_number: summaryHeader.requestNumber || null,
        request_date: summaryHeader.requestDate || null,
        requisitioner: summaryHeader.requisitioner || null,
        release_date: summaryHeader.releaseDate || null,
        mct_rel_number: summaryHeader.mctRelNumber || null,
        wo_number: summaryHeader.woNumber || null,
        jo_number: summaryHeader.joNumber || null,
        so_number: summaryHeader.soNumber || null,
        purpose: summaryHeader.purpose || null,
        notes: summaryHeader.notes || null,
    });

    const buildItemsPayload = () =>
        ticketItems.map((item) => ({
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

    if (uploadedFiles.length > 1) {
        return <BatchUploadMctPage files={uploadedFiles} onFilesChange={setUploadedFiles} />;
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

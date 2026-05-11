import { ListView, ListViewHeader } from "@/components/refine-ui/views/list-view";
import { Search, Plus, FileSpreadsheet, Pencil, Loader2, RefreshCw, History, Trash, Printer } from "lucide-react";
import { Input } from "@/components/ui/input";
import { useCallback, useEffect, useMemo, useState } from "react";
import {
    Select,
    SelectContent,
    SelectItem,
    SelectTrigger,
    SelectValue,
} from "@/components/ui/select";
import { MONTHS_OPTIONS } from "@/constants";
import { CreateButton } from "@/components/refine-ui/buttons/create";
import { DataTable } from "@/components/refine-ui/data-table/data-table";
import { DataTableFilterCombobox } from "@/components/refine-ui/data-table/data-table-filter";
import { DataTableSorter } from "@/components/refine-ui/data-table/data-table-sorter";
import { useTable } from "@refinedev/react-table";
import { ColumnDef } from "@tanstack/react-table";
import { Badge } from "@/components/ui/badge";
import { CrudFilters, useGetIdentity, useInvalidate, useList, useNotification, useOne } from "@refinedev/core";
import { ItemInventoryRow, UserRow } from "@/types";
import { Button } from "@/components/ui/button";
import {
    Dialog,
    DialogContent,
    DialogDescription,
    DialogFooter,
    DialogHeader,
    DialogTitle,
    DialogTrigger,
} from "@/components/ui/dialog";
import { useItemImport } from "@/hooks/use-item-import";
import { ItemImportPanel } from "@/components/items/item-import-panel";
import { Tooltip, TooltipContent, TooltipTrigger } from "@/components/ui/tooltip";
import { Textarea } from "@/components/ui/textarea";
import { useUpdate } from "@refinedev/core";
import { supabaseClient } from "@/providers/supabase-client";
import * as XLSX from "xlsx";

const MONTH_TO_NUMBER: Record<string, number> = {
    January: 1,
    February: 2,
    March: 3,
    April: 4,
    May: 5,
    June: 6,
    July: 7,
    August: 8,
    September: 9,
    October: 10,
    November: 11,
    December: 12,
};

const MONTH_NUMBER_TO_NAME = Object.fromEntries(
    Object.entries(MONTH_TO_NUMBER).map(([month, number]) => [number, month])
) as Record<number, string>;

const ITEMS_LIST_FILTERS_STORAGE_KEY = "items-list-filters";
const ITEMS_LIST_FILTERS_USER_KEY = "items-list-filters-user";

type ItemsListPersistedFilters = {
    selectedMonth: string;
    selectedYear: string;
    searchQuery: string;
};

type ItemInventoryRowWithId = ItemInventoryRow & {
    item_id?: number | string | null;
    inventory_item_id?: number | string | null;
};

type MctHistoryHeader = {
    id: string;
    mct_rel_number: string | null;
    request_number: string | null;
    release_date: string | null;
    requisitioner: string | null;
    purpose: string | null;
    created_at: string | null;
    created_by: string | null;
    status?: string | null;
    rolled_back_at?: string | null;
    rolled_back_by?: string | null;
};

type ItemHistoryMctRow = {
    id: string;
    item_id: string | null;
    item_code: string | null;
    qty: number | null;
    unit_cost: number | null;
    total_cost: number | null;
    c2: number | null;
    deduct_from: string | null;
    remarks: string | null;
    created_at: string | null;
    mcts: MctHistoryHeader | null;
};

type ItemHistoryMctRowDb = Omit<ItemHistoryMctRow, "mcts"> & {
    mcts: MctHistoryHeader | MctHistoryHeader[] | null;
};

type EmergencyHistoryHeader = {
    id: string;
    emergency_date: string | null;
    requisitioner: string | null;
    rel_number: string | null;
    purpose: string | null;
    created_at: string | null;
    created_by: string | null;
    status?: string | null;
    rolled_back_at?: string | null;
    rolled_back_by?: string | null;
};

type ItemHistoryEmergencyRow = {
    id: string;
    emergency_id: string;
    item_id: string | null;
    item_code: string | null;
    qty: number | null;
    unit_cost: number | null;
    total_cost: number | null;
    c2: number | null;
    deduct_from: string | null;
    remarks: string | null;
    created_at: string | null;
    emergencies: EmergencyHistoryHeader | null;
};

type ItemHistoryEmergencyRowDb = Omit<ItemHistoryEmergencyRow, "emergencies"> & {
    emergencies: EmergencyHistoryHeader | EmergencyHistoryHeader[] | null;
};

type ItemHistoryChangeEntry = {
    key: string;
    transactionType: "MCT" | "Emergency";
    label: string;
    delta: number;
    date: string | null | undefined;
    actor: string | null | undefined;
    isRollback?: boolean;
};

type ImportedInventoryItem = {
    item_code: string;
    description: string;
    type: string;
    unit_cost: number | null;
    starting_qty: number | null;
    rowNumber: number;
};

type ImportParseError = {
    rowNumber: number;
    message: string;
};

type ItemImportPreview = {
    items: ImportedInventoryItem[];
    errors: ImportParseError[];
};

type ExportFormat = "pdf" | "excel";
type EditableInventoryField = "unit_cost" | "buffer_stock" | "starting_qty" | "ending_qty";

type EditableInventoryCellProps = {
    ariaLabel: string;
    field: EditableInventoryField;
    isCurrency?: boolean;
    item: ItemInventoryRowWithId;
    value: number | null | undefined;
    onSave: (item: ItemInventoryRowWithId, field: EditableInventoryField, value: number | null) => Promise<boolean>;
};

const EditableInventoryCell = ({
    ariaLabel,
    field,
    isCurrency = false,
    item,
    value,
    onSave,
}: EditableInventoryCellProps) => {
    const [draft, setDraft] = useState(value == null ? "" : String(value));
    const [isSaving, setIsSaving] = useState(false);

    useEffect(() => {
        setDraft(value == null ? "" : String(value));
    }, [value]);

    const commit = useCallback(async () => {
        const trimmedDraft = draft.trim();
        const nextValue = trimmedDraft === "" ? null : Number(trimmedDraft);

        if (trimmedDraft !== "" && !Number.isFinite(nextValue)) {
            setDraft(value == null ? "" : String(value));
            return;
        }

        if (nextValue === (value ?? null)) return;

        setIsSaving(true);
        let didSave = false;
        try {
            didSave = await onSave(item, field, nextValue);
        } catch {
            didSave = false;
        } finally {
            setIsSaving(false);
        }

        if (!didSave) {
            setDraft(value == null ? "" : String(value));
        }
    }, [draft, field, item, onSave, value]);

    return (
        <Input
            aria-label={ariaLabel}
            className="h-8 min-w-20 border-transparent bg-transparent px-2 text-right tabular-nums shadow-none hover:border-input hover:bg-background focus:border-input focus:bg-background"
            disabled={isSaving}
            inputMode={isCurrency ? "decimal" : "numeric"}
            min={0}
            onBlur={() => void commit()}
            onChange={(event) => setDraft(event.target.value)}
            onKeyDown={(event) => {
                if (event.key === "Enter") {
                    event.currentTarget.blur();
                }
                if (event.key === "Escape") {
                    setDraft(value == null ? "" : String(value));
                    event.currentTarget.blur();
                }
            }}
            step={isCurrency ? 0.01 : 1}
            type="number"
            value={draft}
        />
    );
};

const isUuidLike = (value: string) => /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(value);

const getDefaultMonth = () => new Date().toLocaleString("en-US", { month: "long" });
const getDefaultYear = () => String(new Date().getFullYear());

const formatDate = (value?: string | null) => {
    if (!value) return "-";
    const date = new Date(value);
    if (Number.isNaN(date.getTime())) return value;
    return date.toLocaleDateString("en-US", {
        year: "numeric",
        month: "short",
        day: "2-digit",
    });
};


const formatDelta = (value: number, isCurrency = false) => {
    const formatted = isCurrency
        ? value.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 })
        : value.toLocaleString("en-US");
    return value > 0 ? `+${formatted}` : formatted;
};

const readPersistedItemsFilters = (storageKey: string): ItemsListPersistedFilters | null => {
    if (typeof window === "undefined") return null;

    const raw = window.sessionStorage.getItem(storageKey);
    if (!raw) return null;

    try {
        const parsed = JSON.parse(raw) as Partial<ItemsListPersistedFilters>;
        const monthValues = new Set(["all", ...MONTHS_OPTIONS.map((month) => month.value)]);
        const selectedMonth = monthValues.has(parsed.selectedMonth ?? "")
            ? (parsed.selectedMonth as string)
            : getDefaultMonth();
        const selectedYear =
            typeof parsed.selectedYear === "string" && parsed.selectedYear.trim()
                ? parsed.selectedYear
                : getDefaultYear();
        const searchQuery = typeof parsed.searchQuery === "string" ? parsed.searchQuery : "";

        return { selectedMonth, selectedYear, searchQuery };
    } catch {
        return null;
    }
};

const getErrorMessage = (error: unknown) => {
    if (error instanceof Error) return error.message;
    if (typeof error === "string") return error;
    if (error && typeof error === "object" && "message" in error) {
        return String((error as { message?: unknown }).message);
    }
    return "Unknown error";
};

const formatNumber = (value: number | null | undefined) => {
    if (value == null || Number.isNaN(value)) return "-";
    return value.toLocaleString("en-US", { maximumFractionDigits: 0 });
};

const formatMoney = (value: number | null | undefined) => {
    if (value == null || Number.isNaN(value)) return "-";
    return value.toLocaleString("en-US", {
        minimumFractionDigits: 2,
        maximumFractionDigits: 2,
    });
};

const formatExportNumber = (value: number | null | undefined) => {
    if (value == null || Number.isNaN(value)) return "";
    return value;
};

const escapeHtml = (value: string | number | null | undefined) =>
    String(value ?? "")
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#039;");

const sanitizeFileName = (value: string) =>
    value
        .toLowerCase()
        .replace(/[^a-z0-9]+/g, "-")
        .replace(/^-+|-+$/g, "");

const getInventoryReportTitle = (selectedYear: string, selectedMonth: string, selectedType: string) => {
    const typeLabel = selectedType === "all" ? "" : ` - ${selectedType}`;

    if (selectedYear === "all" && selectedMonth === "all") {
        return `Items - All Dates${typeLabel}`;
    }

    if (selectedYear === "all") {
        return `Items for ${selectedMonth} (All Years)${typeLabel}`;
    }

    if (selectedMonth === "all") {
        return `Items for ${selectedYear}${typeLabel}`;
    }

    return `Items as of ${selectedMonth} ${selectedYear}${typeLabel}`;
};

const normalizeImportHeader = (value: string) =>
    value
        .toLowerCase()
        .replace(/\s+/g, "")
        .replace(/[^a-z0-9]/g, "");

const parseImportNumber = (value: string) => {
    if (!value) return null;
    const sanitized = value.replace(/[,\s]/g, "").replace(/[^0-9.+-]/g, "");
    if (!sanitized) return null;
    const parsed = Number(sanitized);
    return Number.isFinite(parsed) ? parsed : null;
};

const ITEM_IMPORT_HEADER_MAP: Record<string, keyof ImportedInventoryItem> = {
    itemcode: "item_code",
    assetnameparticulars: "description",
    assetname: "description",
    particulars: "description",
    description: "description",
    unit: "type",
    type: "type",
    unitcost: "unit_cost",
    unitcostphp: "unit_cost",
    district2onhand: "starting_qty",
};

const parseImportedInventoryItems = (rows: Array<Array<unknown>>): ItemImportPreview => {
    const normalizedRows = rows
        .map((row, index) => ({
            rowNumber: index + 1,
            cells: row.map((cell) => String(cell ?? "").trim()),
        }))
        .filter(({ cells }) => cells.some((cell) => cell.length > 0));

    const headerInfo = normalizedRows
        .slice(0, 25)
        .map(({ cells }, rowIndex) => {
            const headerMap = cells.map((cell) => ITEM_IMPORT_HEADER_MAP[normalizeImportHeader(cell)] ?? null);
            const mappedHeaders = new Set(headerMap.filter(Boolean));
            return {
                rowIndex,
                headerMap,
                hasRequiredHeaders:
                    mappedHeaders.has("item_code") &&
                    mappedHeaders.has("description") &&
                    mappedHeaders.has("type"),
            };
        })
        .find((entry) => entry.hasRequiredHeaders);

    if (!headerInfo) {
        throw new Error("Could not find required headers: Item Code, Asset Name/Particulars, and Unit.");
    }

    const itemsByCode = new Map<string, ImportedInventoryItem>();
    const errors: ImportParseError[] = [];

    for (let i = headerInfo.rowIndex + 1; i < normalizedRows.length; i += 1) {
        const { rowNumber, cells } = normalizedRows[i];
        const importedItem: ImportedInventoryItem = {
            item_code: "",
            description: "",
            type: "",
            unit_cost: null,
            starting_qty: null,
            rowNumber,
        };
        let hasMappedValue = false;
        let unitCostParseError = false;
        let startingQtyParseError = false;

        cells.forEach((cell, colIndex) => {
            const key = headerInfo.headerMap[colIndex];
            if (!key || !cell) return;
            hasMappedValue = true;
            if (key === "unit_cost") {
                importedItem.unit_cost = parseImportNumber(cell);
                unitCostParseError = importedItem.unit_cost == null;
            } else if (key === "starting_qty") {
                importedItem.starting_qty = parseImportNumber(cell);
                startingQtyParseError =
                    importedItem.starting_qty == null ||
                    !Number.isInteger(importedItem.starting_qty);
            } else if (key === "item_code" || key === "description" || key === "type") {
                importedItem[key] = cell;
            }
        });

        const itemCode = importedItem.item_code.trim().toUpperCase();
        const description = importedItem.description.trim();
        const type = importedItem.type.trim();

        if (!hasMappedValue) continue;

        if (!itemCode || !description || !type) {
            errors.push({
                rowNumber,
                message: "Missing item code, asset name/particulars, or unit.",
            });
            continue;
        }

        if (unitCostParseError) {
            errors.push({
                rowNumber,
                message: "Invalid unit cost.",
            });
            continue;
        }

        if (startingQtyParseError) {
            errors.push({
                rowNumber,
                message: "Invalid District 2 on-hand quantity.",
            });
            continue;
        }

        if (itemsByCode.has(itemCode)) {
            errors.push({
                rowNumber,
                message: `Duplicate item code ${itemCode}.`,
            });
            continue;
        }

        itemsByCode.set(itemCode, {
            item_code: itemCode,
            description,
            type,
            unit_cost: importedItem.unit_cost,
            starting_qty: importedItem.starting_qty,
            rowNumber,
        });
    }

    return {
        items: Array.from(itemsByCode.values()),
        errors,
    };
};

const readItemImportPreview = async (file: File): Promise<ItemImportPreview> => {
    const workbook = XLSX.read(await file.arrayBuffer(), {
        type: "array",
        cellDates: false,
    });
    const firstSheetName = workbook.SheetNames[0];
    const firstSheet = firstSheetName ? workbook.Sheets[firstSheetName] : null;

    if (!firstSheet) {
        throw new Error("The selected file does not contain a worksheet.");
    }

    const rows = XLSX.utils.sheet_to_json<Array<unknown>>(firstSheet, {
        header: 1,
        raw: false,
        defval: "",
        blankrows: false,
    }) as Array<Array<unknown>>;

    return parseImportedInventoryItems(rows);
};

const appendExistingItemDuplicateErrors = async (preview: ItemImportPreview): Promise<ItemImportPreview> => {
    if (preview.items.length === 0) return preview;

    const itemCodes = preview.items.map((item) => item.item_code);
    const { data, error } = await supabaseClient
        .from("items")
        .select("item_code")
        .in("item_code", itemCodes);

    if (error) throw error;

    const existingCodes = new Set(
        (data ?? []).map((item) => String(item.item_code ?? "").toUpperCase())
    );

    if (existingCodes.size === 0) return preview;

    const duplicateErrors = preview.items
        .filter((item) => existingCodes.has(item.item_code))
        .map((item) => ({
            rowNumber: item.rowNumber,
            message: `Item code ${item.item_code} already exists.`,
        }));

    return {
        items: preview.items.filter((item) => !existingCodes.has(item.item_code)),
        errors: [...preview.errors, ...duplicateErrors].sort((a, b) => a.rowNumber - b.rowNumber),
    };
};

const buildDateFilters = (
    selectedYear: string,
    selectedMonth: string
): CrudFilters => {
    const filters: CrudFilters = [];

    if (selectedYear !== "all") {
        filters.push({ field: "year", operator: "eq", value: Number(selectedYear) });
    }

    if (selectedMonth !== "all") {
        const monthNumber = MONTH_TO_NUMBER[selectedMonth];
        if (monthNumber) {
            filters.push({ field: "month", operator: "eq", value: monthNumber });
        }
    }

    return filters;
};

const ItemList = () => {
    const initialFilters =
        typeof window === "undefined"
            ? null
            : readPersistedItemsFilters(ITEMS_LIST_FILTERS_STORAGE_KEY);
    const [searchQuery, setSearchQuery] = useState(
        initialFilters?.searchQuery ?? ""
    );
    const [debouncedSearchQuery, setDebouncedSearchQuery] = useState("");
    const [selectedMonth, setSelectedMonth] = useState(() =>
        initialFilters?.selectedMonth ?? getDefaultMonth()
    );
    const [selectedYear, setSelectedYear] = useState<string>(
        initialFilters?.selectedYear ?? getDefaultYear()
    );
    const [importDialogOpen, setImportDialogOpen] = useState(false);
    const [editDialogOpen, setEditDialogOpen] = useState(false);
    const [editingItem, setEditingItem] = useState<ItemInventoryRowWithId | null>(null);
    const [editingItemId, setEditingItemId] = useState<string | number | null>(null);
    const [editItemCode, setEditItemCode] = useState("");
    const [editDescription, setEditDescription] = useState("");
    const [editType, setEditType] = useState("");
    const [editBufferStock, setEditBufferStock] = useState<number | "">("");
    const [editUnitCost, setEditUnitCost] = useState<number | "">("");
    const [editStartingQty, setEditStartingQty] = useState<number | "">("");
    const [editEndingQty, setEditEndingQty] = useState<number | "">("");
    const [itemImportPreview, setItemImportPreview] = useState<ItemImportPreview | null>(null);
    const [itemImportPreviewError, setItemImportPreviewError] = useState<string | null>(null);
    const [isPreparingImport, setIsPreparingImport] = useState(false);
    const [isImportingItems, setIsImportingItems] = useState(false);
    const [isRolloverRunning, setIsRolloverRunning] = useState(false);
    const [exportDialogOpen, setExportDialogOpen] = useState(false);
    const [exportMonth, setExportMonth] = useState(() =>
        initialFilters?.selectedMonth ?? getDefaultMonth()
    );
    const [exportYear, setExportYear] = useState<string>(
        initialFilters?.selectedYear ?? getDefaultYear()
    );
    const [exportType, setExportType] = useState("all");
    const [exportFormat, setExportFormat] = useState<ExportFormat>("pdf");
    const [isExportingItems, setIsExportingItems] = useState(false);
    const { importFile, setImportFile, handleDialogOpenChange, hasImportFile } =
        useItemImport();
    const { open } = useNotification();
    const invalidate = useInvalidate();
    const [historyOpen, setHistoryOpen] = useState(false);
    const [historyItemId, setHistoryItemId] = useState<string>("");
    const [historyItemCode, setHistoryItemCode] = useState<string>("");
    const [historyLoading, setHistoryLoading] = useState(false);
    const [historyError, setHistoryError] = useState<string | null>(null);
    const [mctHistory, setMctHistory] = useState<ItemHistoryMctRow[]>([]);
    const [emergencyHistory, setEmergencyHistory] = useState<ItemHistoryEmergencyRow[]>([]);
    const [historyUsers, setHistoryUsers] = useState<Record<string, UserRow>>({});
    const [deleteDialogOpen, setDeleteDialogOpen] = useState(false);
    const [itemToDelete, setItemToDelete] = useState<ItemInventoryRowWithId | null>(null);
    const [deletingItemId, setDeletingItemId] = useState<string | number | null>(null);
    const [actionColumnSize, setActionColumnSize] = useState(90);
    const [actionContentEl, setActionContentEl] = useState<HTMLDivElement | null>(null);
    const [actionContentWidth, setActionContentWidth] = useState(0);
    const [isDesktop, setIsDesktop] = useState(false);
    const { data: identity } = useGetIdentity<{ id?: string | number; role?: string }>();
    const identityId = identity?.id ? String(identity.id) : "";
    const { result: userResult } = useOne<UserRow>({
        resource: "users",
        id: identityId,
        queryOptions: {
            enabled: Boolean(identityId),
        },
    });
    const normalizedRole = (userResult?.role ?? identity?.role ?? "user").toLowerCase();
    const isAdmin = normalizedRole === "admin";
    const { mutateAsync: updateRecord, mutation } = useUpdate({
        successNotification: false,
    });
    const isUpdatingItem = mutation.isPending;

    const resolveItemId = useCallback((item: ItemInventoryRowWithId) => {
        const fallbackId = item.id != null && isUuidLike(String(item.id)) ? item.id : null;
        return item.item_id ?? item.inventory_item_id ?? fallbackId;
    }, []);

    const openEditDialog = useCallback((item: ItemInventoryRowWithId) => {
        const resolvedId = resolveItemId(item);
        setEditingItem(item);
        setEditingItemId(resolvedId);
        setEditItemCode(item.item_code ?? "");
        setEditDescription(item.description ?? "");
        setEditType(item.type ?? "");
        setEditBufferStock(item.buffer_stock ?? 0);
        setEditUnitCost(item.unit_cost ?? 0);
        setEditStartingQty(item.starting_qty ?? 0);
        setEditEndingQty(item.ending_qty ?? 0);
        setEditDialogOpen(true);
    }, [resolveItemId]);

    const openHistoryDialog = useCallback(
        (item: ItemInventoryRowWithId) => {
            const resolvedId = resolveItemId(item);
            if (!resolvedId) {
                open?.({
                    type: "error",
                    message: "History unavailable",
                    description: "No inventory record was found for this item.",
                });
                return;
            }

            const label = item.item_code
                ? `${item.item_code} — ${item.description ?? ""}`.trim()
                : item.description ?? String(resolvedId);
            setHistoryItemId(String(resolvedId));
            setHistoryItemCode(label);
            setHistoryOpen(true);
        },
        [open, resolveItemId]
    );

    const requestDeleteItem = useCallback((item: ItemInventoryRowWithId) => {
        setItemToDelete(item);
        setDeleteDialogOpen(true);
    }, []);

    const handleSaveEdit = useCallback(async () => {
        if (!editingItemId || !editingItem) {
            open?.({
                type: "error",
                message: "Update failed",
                description: "Missing item details for this row.",
            });
            return;
        }

        try {
            await updateRecord({
                resource: "items",
                id: editingItemId,
                values: {
                    item_code: editItemCode.trim(),
                    description: editDescription.trim(),
                    type: editType.trim(),
                },
            });
        } catch (error) {
            open?.({
                type: "error",
                message: "Item update failed",
                description: getErrorMessage(error),
            });
            return;
        }

        try {
            if (editingItem.month && editingItem.year) {
                const nextStartingQty = editStartingQty === "" ? null : Number(editStartingQty);
                const nextEndingQty = editEndingQty === "" ? null : Number(editEndingQty);
                const nextBufferStock = editBufferStock === "" ? null : Number(editBufferStock);
                const nextUnitCost = editUnitCost === "" ? null : Number(editUnitCost);
                const { error } = await supabaseClient
                    .from("inventory_records")
                    .update({
                        starting_qty: nextStartingQty,
                        ending_qty: nextEndingQty,
                        buffer_stock: nextBufferStock,
                        unit_cost: nextUnitCost,
                    })
                    .eq("month", editingItem.month)
                    .eq("year", editingItem.year)
                    .eq("item_id", editingItemId);

                if (error) {
                    throw error;
                }
            }

            open?.({
                type: "success",
                message: "Item updated",
                description: "Item details have been saved.",
            });

            setEditDialogOpen(false);
            setEditingItem(null);
            setEditingItemId(null);
            itemTable.refineCore.setCurrentPage(1);
            itemTable.refineCore.tableQuery.refetch();
            invalidate({
                resource: "items_inventory_all",
                invalidates: ["list"],
            });
            invalidate({
                resource: "inventory_records",
                invalidates: ["list"],
            });
            invalidate({
                resource: "items",
                invalidates: ["list"],
            });
            invalidate({
                resource: "items_inventory_all",
                invalidates: ["list"],
            });
        } catch (error) {
            open?.({
                type: "error",
                message: "Quantity update failed",
                description: getErrorMessage(error),
            });
        }
    }, [
        editBufferStock,
        editDescription,
        editItemCode,
        editStartingQty,
        editEndingQty,
        editType,
        editUnitCost,
        editingItem,
        editingItemId,
        open,
        updateRecord,
    ]);

    const handleInlineInventorySave = useCallback(
        async (
            item: ItemInventoryRowWithId,
            field: EditableInventoryField,
            value: number | null
        ) => {
            const resolvedId = resolveItemId(item);
            if (!resolvedId || !item.month || !item.year) {
                open?.({
                    type: "error",
                    message: "Update failed",
                    description: "Missing inventory record details for this row.",
                });
                return false;
            }

            if (value != null && value < 0) {
                open?.({
                    type: "error",
                    message: "Update failed",
                    description: "Inventory values must be zero or greater.",
                });
                return false;
            }

            if (field !== "unit_cost" && value != null && !Number.isInteger(value)) {
                open?.({
                    type: "error",
                    message: "Update failed",
                    description: "Quantity and buffer stock values must be whole numbers.",
                });
                return false;
            }

            const { error } = await supabaseClient
                .from("inventory_records")
                .update({ [field]: value })
                .eq("month", item.month)
                .eq("year", item.year)
                .eq("item_id", resolvedId);

            if (error) {
                open?.({
                    type: "error",
                    message: "Update failed",
                    description: getErrorMessage(error),
                });
                return false;
            }

            open?.({
                type: "success",
                message: "Inventory updated",
                description: "The table value has been saved.",
            });
            itemTable.refineCore.tableQuery.refetch();
            invalidate({ resource: "items_inventory_all", invalidates: ["list"] });
            invalidate({ resource: "inventory_records", invalidates: ["list"] });
            return true;
        },
        [invalidate, open, resolveItemId]
    );

    const handleCopyItemCode = useCallback(
        async (itemCode: string) => {
            try {
                await navigator.clipboard.writeText(itemCode);
                open?.({
                    type: "success",
                    message: "Item code copied",
                    description: itemCode,
                });
            } catch {
                open?.({
                    type: "error",
                    message: "Copy failed",
                    description: "Could not copy item code to clipboard.",
                });
            }
        },
        [open]
    );

    const handleActionContentRef = useCallback((node: HTMLDivElement | null) => {
        if (node) {
            setActionContentEl(node);
        }
    }, []);

    useEffect(() => {
        if (typeof window === "undefined") return;
        const mediaQuery = window.matchMedia("(min-width: 1024px)");
        const updateMatch = () => setIsDesktop(mediaQuery.matches);
        updateMatch();
        if (mediaQuery.addEventListener) {
            mediaQuery.addEventListener("change", updateMatch);
            return () => mediaQuery.removeEventListener("change", updateMatch);
        }
        mediaQuery.addListener(updateMatch);
        return () => mediaQuery.removeListener(updateMatch);
    }, []);

    useEffect(() => {
        const el = actionContentEl;
        if (!el || typeof ResizeObserver === "undefined") return;
        const updateWidth = () => {
            setActionContentWidth(el.getBoundingClientRect().width);
        };
        updateWidth();
        const observer = new ResizeObserver(updateWidth);
        observer.observe(el);
        return () => observer.disconnect();
    }, [actionContentEl]);

    useEffect(() => {
        const nextSize = isDesktop ? 90 : actionContentWidth > 90 ? 120 : 90;
        setActionColumnSize(nextSize);
    }, [actionContentWidth, isDesktop]);

    const { result: yearsResult } = useList<ItemInventoryRowWithId>({
        resource: "items_inventory_all",
        pagination: { mode: "off" },
        filters: [],
        sorters: [{ field: "year", order: "desc" }],
    });

    const yearTabs = useMemo(() => {
        return (yearsResult.data ?? [])
            .map((row) => row.year)
            .filter((year): year is number => typeof year === "number")
            .filter((year, index, array) => array.indexOf(year) === index)
            .sort((a, b) => b - a);
    }, [yearsResult.data]);

    const typeOptions = useMemo(() => {
        return (yearsResult.data ?? [])
            .map((row) => row.type?.trim())
            .filter((type): type is string => Boolean(type))
            .filter((type, index, array) => array.indexOf(type) === index)
            .sort((a, b) => a.localeCompare(b));
    }, [yearsResult.data]);

    const getUserLabel = useCallback(
        (userId: string | null | undefined) => {
            if (!userId) return "—";
            const user = historyUsers[userId];
            if (!user) return userId;
            return (
                user.name?.trim() ||
                [user.first_name, user.last_name].filter(Boolean).join(" ") ||
                user.email ||
                userId
            );
        },
        [historyUsers]
    );

    useEffect(() => {
        if (!historyOpen || !historyItemId) return;

        let isMounted = true;
        const loadHistory = async () => {
            setHistoryLoading(true);
            setHistoryError(null);
            try {
                const [mctResult, emergencyResult] = await Promise.all([
                    supabaseClient
                        .from("mct_items")
                        .select(
                            "id,item_id,item_code,qty,unit_cost,total_cost,c2,deduct_from,remarks,created_at,mcts(id,mct_rel_number,request_number,release_date,requisitioner,purpose,created_at,created_by,status,rolled_back_at,rolled_back_by)"
                        )
                        .eq("item_id", historyItemId)
                        .order("created_at", { ascending: false }),
                    supabaseClient
                        .from("emergency_items")
                        .select(
                            "id,emergency_id,item_id,item_code,qty,unit_cost,total_cost,c2,deduct_from,remarks,created_at,emergencies(id,emergency_date,requisitioner,rel_number,purpose,created_at,created_by,status,rolled_back_at,rolled_back_by)"
                        )
                        .eq("item_id", historyItemId)
                        .order("created_at", { ascending: false }),
                ]);

                if (mctResult.error) throw mctResult.error;
                if (emergencyResult.error) throw emergencyResult.error;
                const rawRows: ItemHistoryMctRowDb[] = Array.isArray(mctResult.data)
                    ? (mctResult.data as ItemHistoryMctRowDb[])
                    : [];
                const mctRows: ItemHistoryMctRow[] = rawRows.map((row) => ({
                    ...row,
                    mcts: Array.isArray(row.mcts) ? row.mcts[0] ?? null : row.mcts ?? null,
                }));
                const rawEmergencyRows: ItemHistoryEmergencyRowDb[] = Array.isArray(emergencyResult.data)
                    ? (emergencyResult.data as ItemHistoryEmergencyRowDb[])
                    : [];
                const emergencyRows: ItemHistoryEmergencyRow[] = rawEmergencyRows.map((row) => ({
                    ...row,
                    emergencies: Array.isArray(row.emergencies)
                        ? row.emergencies[0] ?? null
                        : row.emergencies ?? null,
                }));

                if (!isMounted) return;
                setMctHistory(mctRows);
                setEmergencyHistory(emergencyRows);

                const userIds = new Set<string>();
                mctRows.forEach((row) => {
                    if (row.mcts?.created_by) userIds.add(row.mcts.created_by);
                    if (row.mcts?.rolled_back_by) userIds.add(row.mcts.rolled_back_by);
                });
                emergencyRows.forEach((row) => {
                    if (row.emergencies?.created_by) userIds.add(row.emergencies.created_by);
                    if (row.emergencies?.rolled_back_by) userIds.add(row.emergencies.rolled_back_by);
                });

                if (userIds.size > 0) {
                    const { data, error } = await supabaseClient
                        .from("users")
                        .select("id,name,first_name,last_name,email")
                        .in("id", Array.from(userIds));
                    if (error) throw error;
                    if (!isMounted) return;
                    const mapped = (data ?? []).reduce<Record<string, UserRow>>((acc, user) => {
                        acc[String(user.id)] = user as UserRow;
                        return acc;
                    }, {});
                    setHistoryUsers(mapped);
                } else {
                    setHistoryUsers({});
                }
            } catch (error) {
                if (!isMounted) return;
                setHistoryError(getErrorMessage(error));
            } finally {
                if (!isMounted) return;
                setHistoryLoading(false);
            }
        };

        loadHistory();
        return () => {
            isMounted = false;
        };
    }, [historyItemId, historyOpen]);

    const itemChangeEntries = useMemo(() => {
        const mctEntries = mctHistory.flatMap((row) => {
            const entries: ItemHistoryChangeEntry[] = [];
            const mctLabel = row.mcts?.mct_rel_number ?? "MCT";
            const mctDate = row.mcts?.created_at ?? row.created_at;
            if (row.qty != null) {
                const deductFromLabel = row.deduct_from === "buffer_stock" ? "Buffer Stock" : "Ending Qty";
                entries.push({
                    key: `${row.id}-issue`,
                    transactionType: "MCT",
                    label: `${mctLabel} · Deducted from ${deductFromLabel}`,
                    delta: -Math.abs(row.qty),
                    date: mctDate,
                    actor: row.mcts?.created_by,
                });
            }
            if (row.mcts?.status === "rolled_back" && row.mcts.rolled_back_at) {
                const deductFromLabel = row.deduct_from === "buffer_stock" ? "Buffer Stock" : "Ending Qty";
                entries.push({
                    key: `${row.id}-rollback`,
                    transactionType: "MCT",
                    label: `${mctLabel} (Rollback) · Added back to ${deductFromLabel}`,
                    delta: Math.abs(row.qty ?? 0),
                    date: row.mcts.rolled_back_at,
                    actor: row.mcts.rolled_back_by,
                    isRollback: true,
                });
            }
            return entries;
        });
        const emergencyEntries = emergencyHistory.flatMap((row) => {
            const entries: ItemHistoryChangeEntry[] = [];
            const emergencyLabel = row.emergencies?.rel_number ?? "Emergency";
            const emergencyDate = row.emergencies?.created_at ?? row.created_at;
            if (row.qty != null) {
                const deductFromLabel = row.deduct_from === "buffer_stock" ? "Buffer Stock" : "Ending Qty";
                entries.push({
                    key: `${row.id}-emergency-issue`,
                    transactionType: "Emergency",
                    label: `${emergencyLabel} · Emergency deducted from ${deductFromLabel}`,
                    delta: -Math.abs(row.qty),
                    date: emergencyDate,
                    actor: row.emergencies?.created_by,
                });
            }
            if (row.emergencies?.status === "rolled_back" && row.emergencies.rolled_back_at) {
                const deductFromLabel = row.deduct_from === "buffer_stock" ? "Buffer Stock" : "Ending Qty";
                entries.push({
                    key: `${row.id}-emergency-rollback`,
                    transactionType: "Emergency",
                    label: `${emergencyLabel} (Emergency Rollback) · Added back to ${deductFromLabel}`,
                    delta: Math.abs(row.qty ?? 0),
                    date: row.emergencies.rolled_back_at,
                    actor: row.emergencies.rolled_back_by,
                    isRollback: true,
                });
            }
            return entries;
        });

        return [...mctEntries, ...emergencyEntries]
            .sort((a, b) => {
                const dateA = a.date ? new Date(a.date).getTime() : 0;
                const dateB = b.date ? new Date(b.date).getTime() : 0;
                return dateB - dateA;
            });
    }, [emergencyHistory, mctHistory]);


    const itemTable = useTable<ItemInventoryRowWithId>({
        columns: useMemo<ColumnDef<ItemInventoryRowWithId>[]>(
            () => [
                {
                    id: "item_code",
                    accessorKey: "item_code",
                    size: 120,
                    header: ({ column }) => (
                        <div className="flex items-center gap-1">
                            <p className="column-title ml-2 whitespace-normal wrap-break-word leading-tight sm:whitespace-nowrap">
                                Item Code
                            </p>
                            <DataTableSorter column={column} title={undefined} />
                        </div>
                    ),
                    cell: ({ getValue }) => {
                        const itemCode = getValue<string>() ?? "-";
                        return (
                            <Tooltip>
                                <TooltipTrigger asChild>
                                    <Badge
                                        className="cursor-pointer select-none"
                                        onClick={() => void handleCopyItemCode(itemCode)}
                                    >
                                        {itemCode}
                                    </Badge>
                                </TooltipTrigger>
                                <TooltipContent side="top" align="start" className="w-max max-w-[min(90vw,48rem)] whitespace-normal break-words">
                                    {`Click to copy: ${itemCode}`}
                                </TooltipContent>
                            </Tooltip>
                        );
                    },
                },
                {
                    id: "description",
                    accessorKey: "description",
                    size: 450,
                    header: () => (
                        <p className="column-title whitespace-normal wrap-break-word leading-tight sm:whitespace-nowrap">
                            Description
                        </p>
                    ),
                    cell: ({ getValue }) => (
                        <span className="whitespace-normal break-words">{getValue<string>()}</span>
                    ),
                    filterFn: "includesString",
                },
                {
                    id: "type",
                    accessorKey: "type",
                    size: 75,
                    header: ({ column, table }) => (
                        <div className="column-title">
                            <span className="whitespace-normal wrap-break-word leading-tight sm:whitespace-nowrap">UOM</span>
                            <DataTableFilterCombobox
                                column={column}
                                table={table}
                                options={typeOptions.map((type) => ({ label: type, value: type }))}
                                placeholder="UOM"
                                operators={["eq"]}
                            />
                        </div>
                    ),
                    cell: ({ getValue }) => (
                        <Badge variant="secondary">{getValue<string>()}</Badge>
                    ),
                    filterFn: "includesString",
                },
                {
                    id: "unit_cost",
                    accessorKey: "unit_cost",
                    size: 90,
                    header: () => (
                        <p className="column-title whitespace-normal wrap-break-word leading-tight sm:whitespace-nowrap">
                            Unit Cost
                        </p>
                    ),
                    cell: ({ row, getValue }) => (
                        <EditableInventoryCell
                            ariaLabel={`Unit cost for ${row.original.item_code}`}
                            field="unit_cost"
                            isCurrency
                            item={row.original}
                            onSave={handleInlineInventorySave}
                            value={getValue<number | null>()}
                        />
                    ),
                },
                {
                    id: "buffer_stock",
                    accessorKey: "buffer_stock",
                    size: 90,
                    header: () => (
                        <p className="column-title whitespace-normal wrap-break-word leading-tight sm:whitespace-nowrap">
                            Buffer Stock
                        </p>
                    ),
                    cell: ({ row, getValue }) => (
                        <EditableInventoryCell
                            ariaLabel={`Buffer stock for ${row.original.item_code}`}
                            field="buffer_stock"
                            item={row.original}
                            onSave={handleInlineInventorySave}
                            value={getValue<number | null>()}
                        />
                    ),
                },
                {
                    id: "starting_qty",
                    accessorKey: "starting_qty",
                    size: 90,
                    header: () => (
                        <p className="column-title whitespace-normal wrap-break-word leading-tight sm:whitespace-nowrap">
                            Starting Qty.
                        </p>
                    ),
                    cell: ({ row, getValue }) => (
                        <EditableInventoryCell
                            ariaLabel={`Starting quantity for ${row.original.item_code}`}
                            field="starting_qty"
                            item={row.original}
                            onSave={handleInlineInventorySave}
                            value={getValue<number | null>()}
                        />
                    ),
                },
                {
                    id: "ending_qty",
                    accessorKey: "ending_qty",
                    size: 90,
                    header: () => (
                        <p className="column-title whitespace-normal wrap-break-word leading-tight sm:whitespace-nowrap">
                            Ending Qty.
                        </p>
                    ),
                    cell: ({ row, getValue }) => (
                        <EditableInventoryCell
                            ariaLabel={`Ending quantity for ${row.original.item_code}`}
                            field="ending_qty"
                            item={row.original}
                            onSave={handleInlineInventorySave}
                            value={getValue<number | null>()}
                        />
                    ),
                },
                {
                    id: "actions",
                    size: actionColumnSize,
                    header: () => <p className="column-title">Actions</p>,
                    enableSorting: false,
                    enableColumnFilter: false,
                    cell: ({ row }) => (
                        <div className="flex items-center">
                            <div ref={handleActionContentRef} className="inline-flex items-center gap-2">
                                <Button
                                    type="button"
                                    variant="outline"
                                    size="sm"
                                    title="View history"
                                    className="h-7 w-7 p-0"
                                    onClick={() => openHistoryDialog(row.original)}
                                >
                                    <History className="h-3.5 w-3.5" />
                                    <span className="sr-only">View history</span>
                                </Button>
                                <Button
                                    type="button"
                                    variant="outline"
                                    size="sm"
                                    title="Edit item"
                                    className="h-7 w-7 p-0"
                                    onClick={() => openEditDialog(row.original)}
                                >
                                    <Pencil className="h-3.5 w-3.5" />
                                    <span className="sr-only">Edit item</span>
                                </Button>
                                <Button
                                    type="button"
                                    variant="destructive"
                                    size="sm"
                                    title={isAdmin ? "Delete item" : "Only admins can delete items"}
                                    className="h-7 w-7 p-0"
                                    disabled={!isAdmin || Boolean(deletingItemId)}
                                    onClick={() => requestDeleteItem(row.original)}
                                >
                                    <Trash className="h-3.5 w-3.5" />
                                    <span className="sr-only">Delete item</span>
                                </Button>
                            </div>
                        </div>
                    ),
                },
            ],
            [
                actionColumnSize,
                deletingItemId,
                handleCopyItemCode,
                handleActionContentRef,
                handleInlineInventorySave,
                isAdmin,
                openEditDialog,
                openHistoryDialog,
                requestDeleteItem,
                typeOptions,
            ]
        ),
        refineCoreProps: {
            resource: "items_inventory_all",
            pagination: { pageSize: 10, mode: "server" },
            filters: {
                mode: "server",
                initial: buildDateFilters(selectedYear, selectedMonth),
            },
            sorters: {
                initial: [{ field: "item_code", order: "asc" }],
            },
        },
    });
    const columnFilters = itemTable.reactTable.getState().columnFilters;

    const openExportDialog = useCallback(() => {
        const typeFilterValue = columnFilters.find((filter) => filter.id === "type")?.value;
        setExportMonth(selectedMonth);
        setExportYear(selectedYear);
        setExportType(typeof typeFilterValue === "string" && typeFilterValue ? typeFilterValue : "all");
        setExportFormat("pdf");
        setExportDialogOpen(true);
    }, [columnFilters, selectedMonth, selectedYear]);

    const getExportRows = useCallback(async () => {
        let query = supabaseClient
            .from("items_inventory_all")
            .select("id,item_code,description,type,month,year,unit_cost,starting_qty,buffer_stock,ending_qty");

        if (exportYear !== "all") {
            query = query.eq("year", Number(exportYear));
        }

        if (exportMonth !== "all") {
            const monthNumber = MONTH_TO_NUMBER[exportMonth];
            if (monthNumber) {
                query = query.eq("month", monthNumber);
            }
        }

        if (exportType !== "all") {
            query = query.eq("type", exportType);
        }

        const { data, error } = await query.order("item_code", { ascending: true });
        if (error) throw error;
        return (data ?? []) as ItemInventoryRowWithId[];
    }, [exportMonth, exportType, exportYear]);

    const exportItemsToExcel = useCallback((rows: ItemInventoryRowWithId[], title: string) => {
        const worksheetRows = [
            [title],
            [],
            [
                "Item Code",
                "Description",
                "Type/UOM",
                "Unit Cost",
                "Buffer Stock",
                "Starting Qty.",
                "Ending Qty.",
            ],
            ...rows.map((item) => [
                item.item_code ?? "",
                item.description ?? "",
                item.type ?? "",
                formatExportNumber(item.unit_cost),
                formatExportNumber(item.buffer_stock),
                formatExportNumber(item.starting_qty),
                formatExportNumber(item.ending_qty),
            ]),
        ];
        const worksheet = XLSX.utils.aoa_to_sheet(worksheetRows);
        worksheet["!merges"] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 6 } }];
        worksheet["!cols"] = [
            { wch: 16 },
            { wch: 52 },
            { wch: 14 },
            { wch: 14 },
            { wch: 14 },
            { wch: 14 },
            { wch: 14 },
        ];

        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, "Inventory Items");
        XLSX.writeFile(workbook, `${sanitizeFileName(title) || "inventory-items"}.xlsx`);
    }, []);

    const exportItemsToPdf = useCallback((rows: ItemInventoryRowWithId[], title: string) => {
        const generatedAt = new Date().toLocaleString("en-US", {
            month: "short",
            day: "2-digit",
            year: "numeric",
            hour: "2-digit",
            minute: "2-digit",
        });
        const bodyRows = rows.map((item) => `
            <tr>
                <td>${escapeHtml(item.item_code)}</td>
                <td>${escapeHtml(item.description)}</td>
                <td>${escapeHtml(item.type)}</td>
                <td class="num">${escapeHtml(formatMoney(item.unit_cost))}</td>
                <td class="num">${escapeHtml(formatNumber(item.buffer_stock))}</td>
                <td class="num">${escapeHtml(formatNumber(item.starting_qty))}</td>
                <td class="num">${escapeHtml(formatNumber(item.ending_qty))}</td>
            </tr>
        `).join("");

        const reportHtml = `
            <!doctype html>
            <html>
                <head>
                    <title>${escapeHtml(title)}</title>
                    <style>
                        * { box-sizing: border-box; }
                        body {
                            margin: 32px;
                            color: #111827;
                            font-family: Arial, sans-serif;
                            font-size: 12px;
                        }
                        header {
                            margin-bottom: 18px;
                            border-bottom: 1px solid #d1d5db;
                            padding-bottom: 12px;
                        }
                        h1 {
                            margin: 0 0 6px;
                            font-size: 22px;
                            line-height: 1.2;
                        }
                        .meta {
                            color: #4b5563;
                            font-size: 11px;
                        }
                        table {
                            width: 100%;
                            border-collapse: collapse;
                        }
                        th, td {
                            border: 1px solid #d1d5db;
                            padding: 6px 7px;
                            text-align: left;
                            vertical-align: top;
                        }
                        th {
                            background: #f3f4f6;
                            font-weight: 700;
                        }
                        .num {
                            text-align: right;
                            white-space: nowrap;
                        }
                        .empty {
                            border: 1px solid #d1d5db;
                            padding: 16px;
                            text-align: center;
                            color: #4b5563;
                        }
                        @media print {
                            @page {
                                size: landscape;
                                margin: 0;
                            }
                            body {
                                margin: 0;
                                padding: 12mm;
                            }
                            thead { display: table-header-group; }
                            tr { page-break-inside: avoid; }
                        }
                    </style>
                </head>
                <body>
                    <header>
                        <h1>${escapeHtml(title)}</h1>
                        <div class="meta">Generated ${escapeHtml(generatedAt)} · ${rows.length} item${rows.length === 1 ? "" : "s"}</div>
                    </header>
                    ${rows.length > 0 ? `
                        <table>
                            <thead>
                                <tr>
                                    <th>Item Code</th>
                                    <th>Description</th>
                                    <th>Type/UOM</th>
                                    <th>Unit Cost</th>
                                    <th>Buffer Stock</th>
                                    <th>Starting Qty.</th>
                                    <th>Ending Qty.</th>
                                </tr>
                            </thead>
                            <tbody>${bodyRows}</tbody>
                        </table>
                    ` : `<div class="empty">No inventory items matched the selected range.</div>`}
                </body>
            </html>
        `;

        const existingFrame = document.getElementById("inventory-items-print-frame");
        existingFrame?.remove();

        const frame = document.createElement("iframe");
        frame.id = "inventory-items-print-frame";
        frame.title = title;
        frame.style.position = "fixed";
        frame.style.right = "0";
        frame.style.bottom = "0";
        frame.style.width = "0";
        frame.style.height = "0";
        frame.style.border = "0";
        frame.style.visibility = "hidden";
        document.body.appendChild(frame);

        const frameWindow = frame.contentWindow;
        const frameDocument = frameWindow?.document;
        if (!frameWindow || !frameDocument) {
            frame.remove();
            throw new Error("Unable to prepare the print report.");
        }

        frameDocument.open();
        frameDocument.write(reportHtml);
        frameDocument.close();

        window.setTimeout(() => {
            frameWindow.focus();
            frameWindow.print();
            window.setTimeout(() => frame.remove(), 1000);
        }, 250);
    }, []);

    const handleExportItems = useCallback(async () => {
        if (isExportingItems) return;

        setIsExportingItems(true);
        try {
            const rows = await getExportRows();
            const title = getInventoryReportTitle(exportYear, exportMonth, exportType);

            if (exportFormat === "excel") {
                exportItemsToExcel(rows, title);
            } else {
                exportItemsToPdf(rows, title);
            }

            setExportDialogOpen(false);
            open?.({
                type: "success",
                message: exportFormat === "excel" ? "Excel report created" : "Print report opened",
                description: `${rows.length} item${rows.length === 1 ? "" : "s"} included.`,
            });
        } catch (error) {
            open?.({
                type: "error",
                message: "Export failed",
                description: getErrorMessage(error),
            });
        } finally {
            setIsExportingItems(false);
        }
    }, [
        exportFormat,
        exportItemsToExcel,
        exportItemsToPdf,
        exportMonth,
        exportType,
        exportYear,
        getExportRows,
        isExportingItems,
        open,
    ]);

    const resetItemImportPreview = useCallback(() => {
        setItemImportPreview(null);
        setItemImportPreviewError(null);
    }, []);

    const handleImportFileChange = useCallback(
        (file: File | null) => {
            resetItemImportPreview();
            setImportFile(file);
        },
        [resetItemImportPreview, setImportFile]
    );

    const handlePrepareItemImport = useCallback(async () => {
        if (!importFile || isPreparingImport) return;

        setIsPreparingImport(true);
        setItemImportPreviewError(null);
        try {
            const parsedPreview = await readItemImportPreview(importFile);
            const preview = await appendExistingItemDuplicateErrors(parsedPreview);
            setItemImportPreview(preview);
        } catch (error) {
            setItemImportPreview(null);
            setItemImportPreviewError(getErrorMessage(error));
        } finally {
            setIsPreparingImport(false);
        }
    }, [importFile, isPreparingImport]);

    const handleImportItems = useCallback(async () => {
        if (!itemImportPreview || isImportingItems) return;

        setIsImportingItems(true);
        try {
            const importedItems = itemImportPreview.items;

            if (importedItems.length === 0) {
                throw new Error("There are no valid item rows to import.");
            }

            const itemCodes = importedItems.map((item) => item.item_code);
            const { data: existingItems, error: existingItemsError } = await supabaseClient
                .from("items")
                .select("id,item_code")
                .in("item_code", itemCodes);

            if (existingItemsError) throw existingItemsError;

            const existingByCode = new Map(
                (existingItems ?? []).map((item) => [String(item.item_code).toUpperCase(), item])
            );
            const duplicateCodes = importedItems
                .filter((item) => existingByCode.has(item.item_code))
                .map((item) => item.item_code);

            if (duplicateCodes.length > 0) {
                throw new Error(`Duplicate item code${duplicateCodes.length === 1 ? "" : "s"} found: ${duplicateCodes.join(", ")}.`);
            }

            const itemsToCreate = importedItems;
            const itemDetailsToCreate = itemsToCreate.map(({ item_code, description, type }) => ({
                item_code,
                description,
                type,
            }));

            let createdItems: Array<{ id: string | number; item_code: string }> = [];
            if (itemsToCreate.length > 0) {
                const { data, error } = await supabaseClient
                    .from("items")
                    .insert(itemDetailsToCreate)
                    .select("id,item_code");

                if (error) throw error;
                createdItems = (data ?? []) as Array<{ id: string | number; item_code: string }>;
            }

            const importedByCode = new Map(importedItems.map((item) => [item.item_code, item]));
            const importedByItemId = new Map<string, ImportedInventoryItem>();
            createdItems.forEach((item) => {
                const importedItem = importedByCode.get(String(item.item_code).toUpperCase());
                if (item.id != null && importedItem) {
                    importedByItemId.set(String(item.id), importedItem);
                }
            });

            const allImportedItemIds = [
                ...createdItems.map((item) => item.id),
            ].filter((id): id is string | number => id != null);

            if (allImportedItemIds.length > 0) {
                const now = new Date();
                const month = now.getMonth() + 1;
                const year = now.getFullYear();
                const { data: existingRecords, error: existingRecordsError } = await supabaseClient
                    .from("inventory_records")
                    .select("id,item_id")
                    .eq("month", month)
                    .eq("year", year)
                    .in("item_id", allImportedItemIds);

                if (existingRecordsError) throw existingRecordsError;

                const inventoryUpdateResults = await Promise.all(
                    (existingRecords ?? [])
                        .map((record) => {
                            const importedItem = importedByItemId.get(String(record.item_id));
                            if (!importedItem) return null;

                            const values: {
                                unit_cost?: number | null;
                                starting_qty?: number;
                                ending_qty?: number;
                            } = {};

                            if (importedItem.unit_cost != null) {
                                values.unit_cost = importedItem.unit_cost;
                            }

                            if (importedItem.starting_qty != null) {
                                values.starting_qty = importedItem.starting_qty;
                                values.ending_qty = importedItem.starting_qty;
                            }

                            if (Object.keys(values).length === 0) return null;

                            return supabaseClient
                                .from("inventory_records")
                                .update(values)
                                .eq("id", record.id);
                        })
                        .filter((request): request is NonNullable<typeof request> => request != null)
                );
                const inventoryUpdateError = inventoryUpdateResults.find((result) => result.error)?.error;
                if (inventoryUpdateError) throw inventoryUpdateError;

                const existingRecordItemIds = new Set(
                    (existingRecords ?? []).map((record) => String(record.item_id))
                );
                const missingInventoryRecords = allImportedItemIds
                    .filter((id) => !existingRecordItemIds.has(String(id)))
                    .map((itemId) => ({
                        item_id: itemId,
                        month,
                        year,
                        starting_qty: importedByItemId.get(String(itemId))?.starting_qty ?? 0,
                        ending_qty: importedByItemId.get(String(itemId))?.starting_qty ?? 0,
                        buffer_stock: 0,
                        unit_cost: importedByItemId.get(String(itemId))?.unit_cost ?? null,
                    }));

                if (missingInventoryRecords.length > 0) {
                    const { error } = await supabaseClient
                        .from("inventory_records")
                        .insert(missingInventoryRecords);

                    if (error) throw error;
                }
            }

            open?.({
                type: "success",
                message: "Items imported",
                description: `Imported ${importedItems.length} item${importedItems.length === 1 ? "" : "s"}.`,
            });

            handleDialogOpenChange(false);
            resetItemImportPreview();
            setImportDialogOpen(false);
            itemTable.refineCore.setCurrentPage(1);
            itemTable.refineCore.tableQuery.refetch();
            invalidate({ resource: "items_inventory_all", invalidates: ["list"] });
            invalidate({ resource: "inventory_records", invalidates: ["list"] });
            invalidate({ resource: "items", invalidates: ["list"] });
        } catch (error) {
            open?.({
                type: "error",
                message: "Import failed",
                description: getErrorMessage(error),
            });
        } finally {
            setIsImportingItems(false);
        }
    }, [
        handleDialogOpenChange,
        invalidate,
        isImportingItems,
        itemImportPreview,
        itemTable.refineCore,
        open,
        resetItemImportPreview,
    ]);

    const handleDeleteItem = useCallback(async () => {
        const target = itemToDelete;
        if (!target) return;
        const resolvedId = resolveItemId(target);
        if (!resolvedId) {
            open?.({
                type: "error",
                message: "Delete failed",
                description: "Missing item ID for this row.",
            });
            return;
        }

        setDeletingItemId(resolvedId);
        try {
            const { error: inventoryError } = await supabaseClient
                .from("inventory_records")
                .delete()
                .eq("item_id", resolvedId);
            if (inventoryError) {
                const message = inventoryError.message?.toLowerCase() ?? "";
                if (message.includes("inventory_item_id")) {
                    const { error: fallbackError } = await supabaseClient
                        .from("inventory_records")
                        .delete()
                        .eq("inventory_item_id", resolvedId);
                    if (fallbackError) throw fallbackError;
                } else {
                    throw inventoryError;
                }
            }

            const { error: itemError } = await supabaseClient
                .from("items")
                .delete()
                .eq("id", resolvedId);
            if (itemError) throw itemError;

            open?.({
                type: "success",
                message: "Item deleted",
                description: "The item has been removed from inventory.",
            });

            setDeleteDialogOpen(false);
            setItemToDelete(null);
            itemTable.refineCore.setCurrentPage(1);
            itemTable.refineCore.tableQuery.refetch();
            invalidate({ resource: "items_inventory_all", invalidates: ["list"] });
            invalidate({ resource: "inventory_records", invalidates: ["list"] });
            invalidate({ resource: "items", invalidates: ["list"] });
        } catch (error) {
            open?.({
                type: "error",
                message: "Delete failed",
                description: getErrorMessage(error),
            });
        } finally {
            setDeletingItemId(null);
        }
    }, [invalidate, itemTable.refineCore, itemToDelete, open, resolveItemId]);

    const handleManualRollover = useCallback(async () => {
        if (!isAdmin || isRolloverRunning) return;
        if (!identityId) {
            open?.({
                type: "error",
                message: "Rollover failed",
                description: "No user identity found for this session.",
            });
            return;
        }
        setIsRolloverRunning(true);
        let didNotifyFinal = false;
        try {
            const { data, error } = await supabaseClient.rpc("rollover_inventory_month", {
                p_recorded_by: identityId,
            });

            if (error) {
                open?.({
                    type: "error",
                    message: "Rollover failed",
                    description: getErrorMessage(error),
                });
                return;
            }

            const insertedCount =
                typeof data === "number"
                    ? data
                    : typeof data === "string"
                        ? Number(data)
                        : null;

            if (insertedCount === 0) {
                open?.({
                    type: "success",
                    message: "No rollover needed",
                    description: "All current-month records already exist.",
                });
                didNotifyFinal = true;
            } else if (typeof insertedCount === "number" && Number.isFinite(insertedCount)) {
                open?.({
                    type: "success",
                    message: "Rollover complete",
                    description:
                        `Created ${insertedCount} new record${insertedCount === 1 ? "" : "s"}.`,
                });
                didNotifyFinal = true;
            } else {
                open?.({
                    type: "success",
                    message: "Rollover complete",
                    description: "New month records were created when missing.",
                });
                didNotifyFinal = true;
            }

            itemTable.refineCore.tableQuery.refetch();
        } catch (error) {
            open?.({
                type: "error",
                message: "Rollover failed",
                description: getErrorMessage(error),
            });
        } finally {
            if (!didNotifyFinal) {
                open?.({
                    type: "success",
                    message: "Rollover finished",
                    description: "Check the inventory list for any new month records.",
                });
            }
            setIsRolloverRunning(false);
        }
    }, [identityId, isAdmin, isRolloverRunning, itemTable.refineCore, open]);

    useEffect(() => {
        const timeoutId = setTimeout(() => {
            setDebouncedSearchQuery(searchQuery.trim());
        }, 300);

        return () => clearTimeout(timeoutId);
    }, [searchQuery]);

    useEffect(() => {
        if (!identityId || typeof window === "undefined") return;
        const storedUserId = window.sessionStorage.getItem(ITEMS_LIST_FILTERS_USER_KEY);

        if (storedUserId && storedUserId !== identityId) {
            const nextMonth = getDefaultMonth();
            const nextYear = getDefaultYear();
            setSelectedMonth(nextMonth);
            setSelectedYear(nextYear);
            setSearchQuery("");
            window.sessionStorage.removeItem(ITEMS_LIST_FILTERS_STORAGE_KEY);
            window.sessionStorage.setItem(ITEMS_LIST_FILTERS_USER_KEY, identityId);
            return;
        }

        if (!storedUserId) {
            window.sessionStorage.setItem(ITEMS_LIST_FILTERS_USER_KEY, identityId);
        }
    }, [identityId]);

    useEffect(() => {
        if (typeof window === "undefined") return;
        const payload: ItemsListPersistedFilters = {
            selectedMonth,
            selectedYear,
            searchQuery,
        };
        window.sessionStorage.setItem(
            ITEMS_LIST_FILTERS_STORAGE_KEY,
            JSON.stringify(payload)
        );
    }, [searchQuery, selectedMonth, selectedYear]);

    useEffect(() => {
        const filters: CrudFilters = buildDateFilters(selectedYear, selectedMonth);
        const typeFilterValue = columnFilters.find(
            (filter) => filter.id === "type"
        )?.value;
        const selectedTypeFromColumn =
            typeof typeFilterValue === "string" ? typeFilterValue : undefined;

        if (selectedTypeFromColumn) {
            filters.push({ field: "type", operator: "eq", value: selectedTypeFromColumn });
        }

        if (debouncedSearchQuery) {
            filters.push({
                operator: "or",
                value: [
                    { field: "item_code", operator: "contains", value: debouncedSearchQuery },
                    { field: "description", operator: "contains", value: debouncedSearchQuery },
                    { field: "type", operator: "contains", value: debouncedSearchQuery },
                ],
            });
        }

        itemTable.refineCore.setFilters(filters, "replace");
    }, [
        selectedMonth,
        selectedYear,
        debouncedSearchQuery,
        columnFilters,
        itemTable.refineCore.setFilters,
    ]);

    return (
        <ListView>
            <ListViewHeader title="Inventory Items" />
            
            <div className="grid gap-6 min-w-0">
                <div className="intro-row">
                    <p className="text-muted-foreground">Manage and track all items in warehouse inventory</p>
                    <div className="actions-row">
                        <div className="search-field">
                            <Search className="search-icon" />
                            <Input
                                type="text"
                                placeholder="Search item..."
                                className="pl-10 w-full"
                                value={searchQuery}
                                onChange={(e) => setSearchQuery(e.target.value)}
                            />
                        </div>

                    <div className="flex flex-wrap gap-2 w-full sm:w-auto justify-start lg:justify-end">
                        <Dialog
                            open={historyOpen}
                            onOpenChange={(openState) => {
                                setHistoryOpen(openState);
                                if (!openState) {
                                    setHistoryItemId("");
                                    setHistoryItemCode("");
                                    setHistoryError(null);
                                    setMctHistory([]);
                                    setEmergencyHistory([]);
                                    setHistoryUsers({});
                                }
                            }}
                        >
                        <DialogContent
                                className="w-[calc(100vw-2rem)] max-h-[85vh] overflow-hidden sm:max-w-4xl p-4 sm:p-6"
                                onOpenAutoFocus={(event) => event.preventDefault()}
                            >
                                <DialogHeader>
                                    <DialogTitle>Item History</DialogTitle>
                                    <DialogDescription>
                                        Review material charge tickets, emergency releases, and rollbacks for a selected item.
                                    </DialogDescription>
                                </DialogHeader>

                                <div className="grid gap-4">
                                    {historyItemCode ? (
                                        <p className="text-sm text-muted-foreground break-words">
                                            Showing history for{" "}
                                            <span className="font-medium text-foreground">
                                                {historyItemCode}
                                            </span>
                                        </p>
                                    ) : null}

                                    <div className="grid gap-4">
                                        <div className="grid gap-2">
                                            <p className="text-sm font-semibold text-foreground">Inventory Changes</p>
                                            {historyLoading ? (
                                                <div className="flex items-center gap-2 text-sm text-muted-foreground">
                                                    <Loader2 className="h-4 w-4 animate-spin" />
                                                    Loading history...
                                                </div>
                                            ) : historyError ? (
                                                <p className="text-sm text-destructive">{historyError}</p>
                                            ) : itemChangeEntries.length === 0 ? (
                                                <p className="text-sm text-muted-foreground">No inventory changes recorded.</p>
                                            ) : (
                                                <div className="max-h-[40vh] overflow-auto rounded-lg border border-border/80 thin-scrollbar">
                                                    <ul className="divide-y divide-border/70 text-sm">
                                                        {itemChangeEntries.map((entry, index) => (
                                                            <li
                                                                key={entry.key}
                                                                className={`flex flex-col gap-1 p-3 ${index === 0 ? "bg-muted/20" : "bg-background"}`}
                                                            >
                                                                <div className="flex flex-wrap items-center gap-2">
                                                                    <span className="font-medium text-foreground">
                                                                        {entry.label}
                                                                    </span>
                                                                    <span
                                                                        className={`inline-flex items-center rounded-full px-2 py-0.5 text-xs font-semibold ${entry.delta > 0
                                                                            ? "bg-emerald-50 text-emerald-700"
                                                                            : "bg-rose-50 text-rose-700"
                                                                            }`}
                                                                    >
                                                                        {formatDelta(entry.delta)}
                                                                    </span>
                                                                </div>
                                                                <div className="flex flex-wrap items-center gap-2 text-xs text-muted-foreground">
                                                                    <span>{formatDate(entry.date)} · {getUserLabel(entry.actor)}</span>
                                                                    <span
                                                                        className={`inline-flex items-center rounded-full px-2 py-0.5 font-semibold ${entry.transactionType === "Emergency"
                                                                            ? "bg-amber-50 text-amber-700"
                                                                            : "bg-sky-50 text-sky-700"
                                                                            }`}
                                                                    >
                                                                        {entry.transactionType}
                                                                    </span>
                                                                </div>
                                                            </li>
                                                        ))}
                                                    </ul>
                                                </div>
                                            )}
                                        </div>

                                    </div>
                                </div>
                            </DialogContent>
                        </Dialog>
                            <Select value={selectedMonth} onValueChange={setSelectedMonth}>
                                <SelectTrigger>
                                    <SelectValue placeholder="Filter by Month" />
                                </SelectTrigger>
                                <SelectContent>
                                    <SelectItem value="all">All Months</SelectItem>
                                    {MONTHS_OPTIONS.map((month) => (
                                        <SelectItem key={month.value} value={month.value}>
                                            {month.label}
                                        </SelectItem>
                                    ))}
                                </SelectContent>
                            </Select>

                            <Dialog open={exportDialogOpen} onOpenChange={setExportDialogOpen}>
                                <DialogTrigger asChild>
                                    <Button type="button" variant="outline" onClick={openExportDialog}>
                                        <Printer className="w-4 h-4" />
                                        <span>Print / Export</span>
                                    </Button>
                                </DialogTrigger>
                                <DialogContent className="w-[calc(100vw-2rem)] max-w-2xl p-4 sm:p-6">
                                    <DialogHeader>
                                        <DialogTitle>Print Inventory Items</DialogTitle>
                                        <DialogDescription>
                                            Choose the inventory range and file format for the report.
                                        </DialogDescription>
                                    </DialogHeader>

                                    <div className="grid gap-4 py-2">
                                        <div className="grid gap-4 sm:grid-cols-2">
                                            <div className="grid gap-1.5">
                                                <p className="text-sm font-medium">Year</p>
                                                <Select value={exportYear} onValueChange={setExportYear}>
                                                    <SelectTrigger>
                                                        <SelectValue placeholder="Select year" />
                                                    </SelectTrigger>
                                                    <SelectContent>
                                                        <SelectItem value="all">All Years</SelectItem>
                                                        {yearTabs.map((year) => (
                                                            <SelectItem key={year} value={String(year)}>
                                                                {year}
                                                            </SelectItem>
                                                        ))}
                                                    </SelectContent>
                                                </Select>
                                            </div>

                                            <div className="grid gap-1.5">
                                                <p className="text-sm font-medium">Month</p>
                                                <Select value={exportMonth} onValueChange={setExportMonth}>
                                                    <SelectTrigger>
                                                        <SelectValue placeholder="Select month" />
                                                    </SelectTrigger>
                                                    <SelectContent>
                                                        <SelectItem value="all">All Months</SelectItem>
                                                        {MONTHS_OPTIONS.map((month) => (
                                                            <SelectItem key={month.value} value={month.value}>
                                                                {month.label}
                                                            </SelectItem>
                                                        ))}
                                                    </SelectContent>
                                                </Select>
                                            </div>

                                            <div className="grid gap-1.5">
                                                <p className="text-sm font-medium">Item Type / UOM</p>
                                                <Select value={exportType} onValueChange={setExportType}>
                                                    <SelectTrigger>
                                                        <SelectValue placeholder="Select item type" />
                                                    </SelectTrigger>
                                                    <SelectContent>
                                                        <SelectItem value="all">All Types</SelectItem>
                                                        {typeOptions.map((type) => (
                                                            <SelectItem key={type} value={type}>
                                                                {type}
                                                            </SelectItem>
                                                        ))}
                                                    </SelectContent>
                                                </Select>
                                            </div>

                                            <div className="grid gap-1.5">
                                                <p className="text-sm font-medium">Output</p>
                                                <Select value={exportFormat} onValueChange={(value) => setExportFormat(value as ExportFormat)}>
                                                    <SelectTrigger>
                                                        <SelectValue placeholder="Select output" />
                                                    </SelectTrigger>
                                                    <SelectContent>
                                                        <SelectItem value="pdf">PDF</SelectItem>
                                                        <SelectItem value="excel">Excel</SelectItem>
                                                    </SelectContent>
                                                </Select>
                                            </div>
                                        </div>

                                        <div className="rounded-md border bg-muted/20 p-3 text-sm">
                                            <p className="font-medium text-foreground">
                                                {getInventoryReportTitle(exportYear, exportMonth, exportType)}
                                            </p>
                                            <p className="text-xs text-muted-foreground">
                                                This header will be used in the generated report.
                                            </p>
                                        </div>
                                    </div>

                                    <DialogFooter>
                                        <Button
                                            type="button"
                                            variant="outline"
                                            onClick={() => setExportDialogOpen(false)}
                                            disabled={isExportingItems}
                                        >
                                            Cancel
                                        </Button>
                                        <Button
                                            type="button"
                                            onClick={() => void handleExportItems()}
                                            disabled={isExportingItems}
                                        >
                                            {isExportingItems ? (
                                                <span className="inline-flex items-center gap-2">
                                                    <Loader2 className="h-4 w-4 animate-spin" />
                                                    Preparing
                                                </span>
                                            ) : (
                                                "Generate"
                                            )}
                                        </Button>
                                    </DialogFooter>
                                </DialogContent>
                            </Dialog>

                            <Dialog
                                open={importDialogOpen}
                                onOpenChange={(isOpen) => {
                                    setImportDialogOpen(isOpen);
                                    handleDialogOpenChange(isOpen);
                                    if (!isOpen) {
                                        resetItemImportPreview();
                                    }
                                }}
                            >
                                <DialogTrigger asChild>
                                    <Button type="button" variant="outline">
                                        <FileSpreadsheet className="w-4 h-4" />
                                        <span>Import Excel</span>
                                    </Button>
                                </DialogTrigger>
                                <DialogContent className="w-[calc(100vw-2rem)] max-w-xl p-4 sm:p-6">
                                    <ItemImportPanel
                                        file={importFile}
                                        onFileChange={handleImportFileChange}
                                        onCancel={() => {
                                            handleDialogOpenChange(false);
                                            resetItemImportPreview();
                                            setImportDialogOpen(false);
                                        }}
                                        onContinue={itemImportPreview ? handleImportItems : handlePrepareItemImport}
                                        continueDisabled={
                                            !hasImportFile ||
                                            isPreparingImport ||
                                            isImportingItems ||
                                            (itemImportPreview ? itemImportPreview.items.length === 0 : false)
                                        }
                                        continueLabel={
                                            isPreparingImport
                                                ? "Checking..."
                                                : isImportingItems
                                                    ? "Importing..."
                                                    : itemImportPreview
                                                        ? "Import Items"
                                                        : "Continue"
                                        }
                                        preview={
                                            itemImportPreview || itemImportPreviewError ? (
                                                <div className="rounded-md border bg-muted/20 p-3 text-sm">
                                                    {itemImportPreview ? (
                                                        <div className="space-y-3">
                                                            <div className="grid grid-cols-2 gap-2">
                                                                <div className="rounded-md border bg-background p-3">
                                                                    <p className="text-xs text-muted-foreground">Successful imports</p>
                                                                    <p className="text-2xl font-semibold">{itemImportPreview.items.length}</p>
                                                                </div>
                                                                <div className="rounded-md border bg-background p-3">
                                                                    <p className="text-xs text-muted-foreground">Errors</p>
                                                                    <p className="text-2xl font-semibold">{itemImportPreview.errors.length}</p>
                                                                </div>
                                                            </div>
                                                            {itemImportPreview.errors.length > 0 ? (
                                                                <div className="space-y-1 text-xs text-muted-foreground">
                                                                    {itemImportPreview.errors.slice(0, 3).map((error) => (
                                                                        <p key={`${error.rowNumber}-${error.message}`}>
                                                                            Row {error.rowNumber}: {error.message}
                                                                        </p>
                                                                    ))}
                                                                    {itemImportPreview.errors.length > 3 ? (
                                                                        <p>
                                                                            {itemImportPreview.errors.length - 3} more error
                                                                            {itemImportPreview.errors.length - 3 === 1 ? "" : "s"}.
                                                                        </p>
                                                                    ) : null}
                                                                </div>
                                                            ) : null}
                                                        </div>
                                                    ) : (
                                                        <p className="text-sm text-destructive">{itemImportPreviewError}</p>
                                                    )}
                                                </div>
                                            ) : null
                                        }
                                    />
                                </DialogContent>
                            </Dialog>

                            {isAdmin ? (
                                <Button
                                    type="button"
                                    variant="outline"
                                    onClick={handleManualRollover}
                                    disabled={isRolloverRunning}
                                >
                                    {isRolloverRunning ? (
                                        <span className="inline-flex items-center gap-2">
                                            <Loader2 className="h-4 w-4 animate-spin" />
                                            Running
                                        </span>
                                    ) : (
                                        <span className="inline-flex items-center gap-2">
                                            <RefreshCw className="h-4 w-4" />
                                            Rollover Month
                                        </span>
                                    )}
                                </Button>
                            ) : null}

                            <CreateButton>
                                <div className="flex items-center gap-2 font-semibold">
                                    <Plus className="w-4 h-4" />
                                    <span>Add Item</span>
                                </div>
                            </CreateButton>

                        </div>
                    </div>
                </div>
                <div className="min-w-0">
                    <DataTable
                        table={itemTable}
                        bottomContent={
                            <div className="-mt-px overflow-x-auto px-2">
                                <div className="flex min-w-max items-start gap-1">
                                    <button
                                        type="button"
                                        onClick={() => setSelectedYear("all")}
                                        className={`${selectedYear === "all"
                                            ? "relative rounded-b-lg border border-t-0 border-border bg-background px-4 py-1.5 text-sm text-foreground shadow-sm [box-shadow:inset_0_1px_0_var(--background)]"
                                            : "relative rounded-b-lg border border-t-0 border-border bg-muted/40 px-4 py-1.5 text-sm text-muted-foreground"
                                            }`}
                                    >
                                        All
                                    </button>

                                    {yearTabs.map((year) => (
                                        <button
                                            key={year}
                                            type="button"
                                            onClick={() => setSelectedYear(String(year))}
                                            className={`${selectedYear === String(year)
                                                ? "relative rounded-b-lg border border-t-0 border-border bg-background px-4 py-1.5 text-sm text-foreground shadow-sm [box-shadow:inset_0_1px_0_var(--background)]"
                                                : "relative rounded-b-lg border border-t-0 border-border bg-muted/40 px-4 py-1.5 text-sm text-muted-foreground"
                                                }`}
                                        >
                                            {year}
                                        </button>
                                    ))}
                                </div>
                            </div>
                        }
                    />
                </div>
            </div>
            <Dialog
                open={editDialogOpen}
                onOpenChange={(openState) => {
                    setEditDialogOpen(openState);
                    if (!openState) {
                        setEditingItem(null);
                        setEditingItemId(null);
                    }
                }}
            >
                <DialogContent className="sm:max-w-2xl overflow-hidden p-0 border-border/80 shadow-sm">
                    <DialogHeader className="border-b px-6 py-5">
                        <DialogTitle className="text-2xl">Edit Item</DialogTitle>
                        <DialogDescription>Update item details and inventory quantity.</DialogDescription>
                    </DialogHeader>

                    <div className="grid gap-5 px-6 py-6">
                        <div className="grid gap-4 rounded-xl border border-border/80 bg-muted/10 p-4">
                            <p className="text-xs font-semibold uppercase tracking-wide text-muted-foreground">Item Details</p>
                            <div className="grid gap-4 sm:grid-cols-2">
                                <div className="grid gap-1.5">
                                    <p className="text-sm font-medium">Item Code</p>
                                    <Input
                                        value={editItemCode}
                                        onChange={(e) => setEditItemCode(e.target.value.toUpperCase())}
                                        placeholder="INV-××××××××"
                                        className="bg-background"
                                    />
                                </div>
                                <div className="grid gap-1.5">
                                    <p className="text-sm font-medium">Type</p>
                                    <Input
                                        value={editType}
                                        onChange={(e) => setEditType(e.target.value.toUpperCase())}
                                        placeholder="Type"
                                        className="bg-background"
                                    />
                                </div>
                            </div>

                            <div className="grid gap-1.5">
                                <p className="text-sm font-medium">Description</p>
                                <Textarea
                                    value={editDescription}
                                    onChange={(e) => setEditDescription(e.target.value.toUpperCase())}
                                    placeholder="Describe the item and specification"
                                    className="min-h-24 bg-background"
                                />
                            </div>
                        </div>

                        <div className="grid gap-4 rounded-xl border border-border/80 bg-muted/10 p-4">
                            <p className="text-xs font-semibold uppercase tracking-wide text-muted-foreground">Inventory</p>
                            <div className="grid gap-4 sm:grid-cols-2">
                                <div className="grid gap-1.5">
                                    <p className="text-sm font-medium">Buffer Stock</p>
                                    <Input
                                        type="number"
                                        min={0}
                                        step={1}
                                        value={editBufferStock}
                                        onChange={(e) => {
                                            const value = e.target.value;
                                            setEditBufferStock(value === "" ? "" : Number(value));
                                        }}
                                        className="bg-background"
                                    />
                                </div>
                                <div className="grid gap-1.5">
                                    <p className="text-sm font-medium">Unit Cost</p>
                                    <Input
                                        type="number"
                                        min={0}
                                        step={0.01}
                                        value={editUnitCost}
                                        onChange={(e) => {
                                            const value = e.target.value;
                                            setEditUnitCost(value === "" ? "" : Number(value));
                                        }}
                                        className="bg-background"
                                    />
                                </div>
                                <div className="grid gap-1.5">
                                    <p className="text-sm font-medium">Starting Quantity</p>
                                    <Input
                                        type="number"
                                        min={0}
                                        step={1}
                                        value={editStartingQty}
                                        onChange={(e) => {
                                            const value = e.target.value;
                                            setEditStartingQty(value === "" ? "" : Number(value));
                                        }}
                                        className="bg-background"
                                    />
                                </div>
                                <div className="grid gap-1.5">
                                    <p className="text-sm font-medium">Ending Quantity</p>
                                    <Input
                                        type="number"
                                        min={0}
                                        step={1}
                                        value={editEndingQty}
                                        onChange={(e) => {
                                            const value = e.target.value;
                                            setEditEndingQty(value === "" ? "" : Number(value));
                                        }}
                                        className="bg-background"
                                    />
                                </div>
                            </div>
                        </div>
                    </div>

                    <DialogFooter className="items-center border-t px-6 py-4 sm:justify-end">
                        <p className="mr-auto text-left text-xs text-muted-foreground">
                            {isUpdatingItem ? "Saving changes..." : "Changes apply immediately."}
                        </p>
                        <Button
                            type="button"
                            variant="outline"
                            onClick={() => setEditDialogOpen(false)}
                            disabled={isUpdatingItem}
                        >
                            Cancel
                        </Button>
                        <Button
                            type="button"
                            onClick={handleSaveEdit}
                            disabled={isUpdatingItem || !editingItem}
                        >
                            {isUpdatingItem ? (
                                <span className="inline-flex items-center gap-2">
                                    <Loader2 className="h-4 w-4 animate-spin" />
                                    Saving
                                </span>
                            ) : (
                                "Save Changes"
                            )}
                        </Button>
                    </DialogFooter>
                </DialogContent>
            </Dialog>
            <Dialog
                open={deleteDialogOpen}
                onOpenChange={(openState) => {
                    setDeleteDialogOpen(openState);
                    if (!openState) setItemToDelete(null);
                }}
            >
                <DialogContent className="sm:max-w-md">
                    <DialogHeader>
                        <DialogTitle>Delete Item</DialogTitle>
                        <DialogDescription>
                            This will permanently remove the item and its inventory records.
                        </DialogDescription>
                    </DialogHeader>

                    <div className="rounded-md border p-3 text-sm text-muted-foreground">
                        <p>
                            Item Code:{" "}
                            <span className="text-foreground">{itemToDelete?.item_code ?? "-"}</span>
                        </p>
                        <p>
                            Description:{" "}
                            <span className="text-foreground">{itemToDelete?.description ?? "-"}</span>
                        </p>
                    </div>

                    <DialogFooter>
                        <Button
                            type="button"
                            variant="outline"
                            onClick={() => {
                                setDeleteDialogOpen(false);
                                setItemToDelete(null);
                            }}
                            disabled={Boolean(deletingItemId)}
                        >
                            Cancel
                        </Button>
                        <Button
                            type="button"
                            variant="destructive"
                            onClick={() => void handleDeleteItem()}
                            disabled={Boolean(deletingItemId) || !itemToDelete}
                        >
                            {deletingItemId ? (
                                <span className="inline-flex items-center gap-2">
                                    <Loader2 className="h-4 w-4 animate-spin" />
                                    Deleting
                                </span>
                            ) : (
                                "Delete Item"
                            )}
                        </Button>
                    </DialogFooter>
                </DialogContent>
            </Dialog>
        </ListView>
    );
};

export default ItemList;

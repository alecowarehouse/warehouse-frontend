import { CreateView, CreateViewHeader } from "@/components/refine-ui/views/create-view";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardFooter, CardHeader, CardTitle } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from "@/components/ui/table";
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
import { useEffect, useMemo, useState, type KeyboardEventHandler } from "react";
import { Check, Plus, Search, Trash2 } from "lucide-react";
import { supabaseClient } from "@/providers/supabase-client";
import { useGetIdentity, useGo, useInvalidate, useNotification } from "@refinedev/core";

type EmergencyHeader = {
    date: string;
    requisitioner: string;
    relNumber: string;
    purpose: string;
};

type EmergencyItem = {
    id: string;
    item_id: string;
    item_code: string;
    particulars: string;
    unit: string;
    unit_cost: number | null;
    qty: number | null;
    total_cost: number | null;
    c2: number | null;
    deduct_from: "ending_qty" | "buffer_stock";
    remarks: string;
};

type CatalogItem = {
    id: string;
    item_code: string | null;
    description: string | null;
    type: string | null;
};

type InventoryRecord = {
    id: string;
    item_id: string;
    ending_qty: number | null;
    buffer_stock: number | null;
    unit_cost: number | null;
};

type AvailabilityStatus = "in_stock" | "insufficient" | "missing";

type AvailabilityInfo = {
    status: AvailabilityStatus;
    availableQty?: number | null;
    bufferStock?: number | null;
    endingQty?: number | null;
    deductFrom?: "ending_qty" | "buffer_stock";
};

const createRowId = () => `emergency-row-${Date.now()}-${Math.random().toString(16).slice(2)}`;

const parseNumber = (value: string): number | null => {
    const normalized = value.replace(/,/g, "").trim();
    if (!normalized) return null;
    const parsed = Number(normalized);
    return Number.isFinite(parsed) ? parsed : null;
};

const formatDecimal = (value: number | null) => {
    if (value == null || Number.isNaN(value)) return "-";
    return value.toLocaleString("en-US", {
        minimumFractionDigits: 2,
        maximumFractionDigits: 2,
    });
};

const sumNumbers = (values: Array<number | null>) =>
    values.reduce<number>((acc, value) => (value == null || Number.isNaN(value) ? acc : acc + value), 0);

const normalizeSearchText = (value: string) => value.toLowerCase().trim().replace(/\s+/g, " ");

const EmergencyMovementCreatePage = () => {
    const currentDate = new Date();
    const currentMonth = currentDate.getMonth() + 1;
    const currentYear = currentDate.getFullYear();

    const [header, setHeader] = useState<EmergencyHeader>({
        date: "",
        requisitioner: "",
        relNumber: "",
        purpose: "",
    });
    const [items, setItems] = useState<EmergencyItem[]>([]);
    const [itemSearch, setItemSearch] = useState("");
    const [activeSearchIndex, setActiveSearchIndex] = useState(0);
    const [validationErrors, setValidationErrors] = useState<string[]>([]);
    const [missingInventoryItems, setMissingInventoryItems] = useState<EmergencyItem[]>([]);
    const [confirmOpen, setConfirmOpen] = useState(false);
    const [errorDialogOpen, setErrorDialogOpen] = useState(false);
    const [isSubmitting, setIsSubmitting] = useState(false);
    const [catalogItems, setCatalogItems] = useState<CatalogItem[]>([]);
    const [inventoryRows, setInventoryRows] = useState<InventoryRecord[]>([]);
    const [isLookupLoading, setIsLookupLoading] = useState(false);
    const [lookupError, setLookupError] = useState<string | null>(null);

    const { data: identity } = useGetIdentity<{ id?: string | number }>();
    const { open } = useNotification();
    const invalidate = useInvalidate();
    const go = useGo();

    useEffect(() => {
        let isActive = true;

        const loadLookupData = async () => {
            setIsLookupLoading(true);
            setLookupError(null);
            try {
                const { data: itemRows, error: itemError } = await supabaseClient
                    .from("items")
                    .select("id,item_code,description,type")
                    .order("item_code", { ascending: true });

                if (itemError) {
                    throw itemError;
                }

                const { data: inventoryData, error: inventoryError } = await supabaseClient
                    .from("inventory_records")
                    .select("id,item_id,ending_qty,buffer_stock,unit_cost")
                    .eq("month", currentMonth)
                    .eq("year", currentYear);

                if (inventoryError) {
                    throw inventoryError;
                }

                if (!isActive) return;
                setCatalogItems((itemRows as CatalogItem[] | null) ?? []);
                setInventoryRows((inventoryData as InventoryRecord[] | null) ?? []);
            } catch (error) {
                const message = error instanceof Error ? error.message : "Unable to load item catalog.";
                if (!isActive) return;
                setLookupError(message);
                open?.({
                    type: "error",
                    message: "Unable to load item catalog",
                    description: message,
                });
            } finally {
                if (isActive) {
                    setIsLookupLoading(false);
                }
            }
        };

        void loadLookupData();

        return () => {
            isActive = false;
        };
    }, [currentMonth, currentYear, open]);

    const inventoryByItemId = useMemo(() => {
        const map = new Map<string, InventoryRecord>();
        for (const row of inventoryRows) {
            map.set(String(row.item_id), row);
        }
        return map;
    }, [inventoryRows]);

    const availabilityMap = useMemo(() => {
        const next: Record<string, AvailabilityInfo> = {};

        for (const row of items) {
            const inventory = inventoryByItemId.get(row.item_id);
            if (!inventory) {
                next[row.id] = { status: "missing" };
                continue;
            }

            const requestedQty = row.qty ?? 0;
            const deductFrom = row.deduct_from;
            const effectiveAvailable =
                deductFrom === "buffer_stock"
                    ? (inventory.buffer_stock ?? 0)
                    : (inventory.ending_qty ?? 0);

            if (effectiveAvailable - requestedQty < 0) {
                next[row.id] = {
                    status: "insufficient",
                    availableQty: effectiveAvailable,
                    bufferStock: inventory.buffer_stock ?? 0,
                    endingQty: inventory.ending_qty ?? 0,
                    deductFrom,
                };
                continue;
            }

            next[row.id] = {
                status: "in_stock",
                availableQty: effectiveAvailable,
                bufferStock: inventory.buffer_stock ?? 0,
                endingQty: inventory.ending_qty ?? 0,
                deductFrom,
            };
        }

        return next;
    }, [inventoryByItemId, items]);

    const totalQty = useMemo(() => sumNumbers(items.map((item) => item.qty)), [items]);
    const totalCost = useMemo(() => sumNumbers(items.map((item) => item.total_cost)), [items]);

    const filteredCatalog = useMemo(() => {
        const query = normalizeSearchText(itemSearch);
        if (!query) return [];

        const queryTokens = query.split(" ").filter(Boolean);
        const scored = catalogItems
            .map((item) => {
                const code = normalizeSearchText(item.item_code ?? "");
                const description = normalizeSearchText(item.description ?? "");
                const searchBlob = `${code} ${description}`.trim();
                if (!searchBlob) return null;

                let score = 0;
                const exactCodeMatch = code === query;
                const startsWithCode = code.startsWith(query);
                const includesCode = code.includes(query);
                const includesDescription = description.includes(query);

                if (exactCodeMatch) score += 1200;
                else if (startsWithCode) score += 900;
                else if (includesCode) score += 700;
                else if (includesDescription) score += 350;

                for (const token of queryTokens) {
                    if (code.startsWith(token)) score += 120;
                    else if (code.includes(token)) score += 80;
                    else if (description.includes(token)) score += 40;
                    else return null;
                }

                if (score <= 0) return null;

                return {
                    item,
                    score,
                    code,
                };
            })
            .filter((entry): entry is { item: CatalogItem; score: number; code: string } => entry !== null)
            .sort((a, b) => {
                if (b.score !== a.score) return b.score - a.score;
                return a.code.localeCompare(b.code);
            })
            .slice(0, 12)
            .map((entry) => entry.item);

        return scored;
    }, [catalogItems, itemSearch]);

    useEffect(() => {
        setActiveSearchIndex(0);
    }, [itemSearch]);

    useEffect(() => {
        if (filteredCatalog.length === 0) {
            setActiveSearchIndex(0);
            return;
        }
        setActiveSearchIndex((prev) => Math.max(0, Math.min(prev, filteredCatalog.length - 1)));
    }, [filteredCatalog]);

    const updateRow = (rowId: string, patch: Partial<EmergencyItem>) => {
        setItems((prev) =>
            prev.map((row) => {
                if (row.id !== rowId) return row;
                const next = { ...row, ...patch };
                if ((patch.qty !== undefined || patch.unit_cost !== undefined) && next.qty != null && next.unit_cost != null) {
                    next.total_cost = next.qty * next.unit_cost;
                }
                return next;
            })
        );
    };

    const addItemFromCatalog = (catalog: CatalogItem) => {
        const itemId = String(catalog.id);
        const existing = items.find((row) => row.item_id === itemId);
        if (existing) {
            open?.({
                type: "error",
                message: "Item already added",
                description: `${catalog.item_code ?? "Item"} is already in the table.`,
            });
            return;
        }

        const inventory = inventoryByItemId.get(itemId);
        const unitCost = inventory?.unit_cost ?? null;

        const nextRow: EmergencyItem = {
            id: createRowId(),
            item_id: itemId,
            item_code: catalog.item_code ?? "",
            particulars: catalog.description ?? "",
            unit: catalog.type ?? "",
            unit_cost: unitCost,
            qty: 1,
            total_cost: unitCost != null ? unitCost : null,
            c2: null,
            deduct_from: "ending_qty",
            remarks: "",
        };

        setItems((prev) => [...prev, nextRow]);
        setItemSearch("");
        setActiveSearchIndex(0);
    };

    const handleItemSearchKeyDown: KeyboardEventHandler<HTMLInputElement> = (event) => {
        if (!itemSearch.trim() || filteredCatalog.length === 0) return;

        if (event.key === "ArrowDown") {
            event.preventDefault();
            setActiveSearchIndex((prev) => Math.min(prev + 1, filteredCatalog.length - 1));
            return;
        }

        if (event.key === "ArrowUp") {
            event.preventDefault();
            setActiveSearchIndex((prev) => Math.max(prev - 1, 0));
            return;
        }

        if (event.key === "Enter") {
            event.preventDefault();
            const candidate = filteredCatalog[activeSearchIndex];
            if (candidate) {
                addItemFromCatalog(candidate);
            }
            return;
        }

        if (event.key === "Escape") {
            event.preventDefault();
            setItemSearch("");
            setActiveSearchIndex(0);
        }
    };

    const removeRow = (rowId: string) => {
        setItems((prev) => prev.filter((row) => row.id !== rowId));
    };

    const validateInputs = () => {
        const errors: string[] = [];

        if (!header.date.trim()) errors.push("Date is required.");
        if (!header.requisitioner.trim()) errors.push("Requisitioner is required.");
        if (!header.relNumber.trim()) errors.push("Rel # is required.");
        if (!header.purpose.trim()) errors.push("Purpose is required.");

        if (items.length === 0) {
            errors.push("Add at least one item using the search field.");
        }

        items.forEach((item, index) => {
            if (!item.item_code.trim()) errors.push(`Row ${index + 1}: Missing item code.`);
            if (item.qty == null || Number.isNaN(item.qty) || item.qty <= 0) {
                errors.push(`Row ${index + 1}: Quantity must be greater than 0.`);
            }
        });

        setValidationErrors(errors);
        setErrorDialogOpen(errors.length > 0);
        return errors.length === 0;
    };

    const buildHeaderPayload = () => ({
        emergency_date: header.date || null,
        requisitioner: header.requisitioner || null,
        rel_number: header.relNumber || null,
        purpose: header.purpose || null,
        notes: "Emergency",
    });

    const buildItemsPayload = () =>
        items.map((item) => ({
            item_code: item.item_code.trim() || null,
            particulars: item.particulars || null,
            unit: item.unit || null,
            unit_cost: item.unit_cost ?? null,
            qty: item.qty ?? null,
            total_cost: item.total_cost ?? null,
            c2: item.c2 ?? null,
            deduct_from: item.deduct_from,
            remarks: item.remarks.trim() || null,
        }));

    const parseCodesFromMessage = (message: string, prefix: string) => {
        if (!message.startsWith(prefix)) return [];
        const raw = message.slice(prefix.length).trim();
        if (!raw) return [];
        return raw.split(",").map((code) => code.trim()).filter(Boolean);
    };

    const saveEmergency = async (createMissingInventory: boolean) => {
        const { error } = await supabaseClient.rpc("create_emergency_transaction", {
            p_header: buildHeaderPayload(),
            p_items: buildItemsPayload(),
            p_create_missing_inventory: createMissingInventory,
            p_created_by: identity?.id ? String(identity.id) : null,
        });

        return error;
    };

    const handleSubmit = async () => {
        if (isSubmitting) return;
        setValidationErrors([]);
        setMissingInventoryItems([]);

        if (!validateInputs()) return;

        setIsSubmitting(true);
        try {
            const error = await saveEmergency(false);

            if (error) {
                const message = error.message ?? "Unable to save emergency record.";
                const duplicate = parseCodesFromMessage(message, "duplicate_emergency:");
                if (duplicate.length > 0) {
                    setValidationErrors([`Duplicate Rel # detected: ${duplicate.join(", ")}`]);
                    setErrorDialogOpen(true);
                    return;
                }

                const missingItems = parseCodesFromMessage(message, "missing_item_codes:");
                if (missingItems.length > 0) {
                    setValidationErrors(missingItems.map((code) => `Item code not found: ${code}`));
                    setErrorDialogOpen(true);
                    return;
                }

                const insufficientInventory = parseCodesFromMessage(message, "insufficient_inventory:");
                if (insufficientInventory.length > 0) {
                    setValidationErrors(
                        insufficientInventory.map((code) => `Insufficient inventory for item code: ${code}`)
                    );
                    setErrorDialogOpen(true);
                    return;
                }

                const missingInventory = parseCodesFromMessage(message, "missing_inventory:");
                if (missingInventory.length > 0) {
                    setMissingInventoryItems(
                        items.filter((item) => missingInventory.includes(item.item_code.trim().toUpperCase()))
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
                message: "Emergency saved",
                description: "Emergency transaction saved.",
            });

            await invalidate({ resource: "emergencies", invalidates: ["list"] });
            await invalidate({ resource: "emergency_items", invalidates: ["list"] });

            go({ to: "/emergency", type: "replace" });
        } catch (error) {
            const description = error instanceof Error ? error.message : "Unable to save emergency record.";
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
            const error = await saveEmergency(true);
            if (error) {
                const message = error.message ?? "Unable to save emergency record.";
                const duplicate = parseCodesFromMessage(message, "duplicate_emergency:");
                if (duplicate.length > 0) {
                    setValidationErrors([`Duplicate Rel # detected: ${duplicate.join(", ")}`]);
                } else {
                    setValidationErrors([message]);
                }
                setErrorDialogOpen(true);
                return;
            }

            open?.({
                type: "success",
                message: "Emergency saved",
                description: "Emergency transaction saved.",
            });

            await invalidate({ resource: "emergencies", invalidates: ["list"] });
            await invalidate({ resource: "emergency_items", invalidates: ["list"] });

            go({ to: "/emergency", type: "replace" });
        } catch (error) {
            const description = error instanceof Error ? error.message : "Unable to save emergency record.";
            setValidationErrors([description]);
            setErrorDialogOpen(true);
        } finally {
            setIsSubmitting(false);
            setConfirmOpen(false);
        }
    };

    const isCatalogLoading = isLookupLoading;

    return (
        <CreateView className="item-view">
            <CreateViewHeader resource="emergencies" title="Emergencies" />
            <div className="my-4 flex items-center">
                <Card className="w-full max-w-7xl mx-auto item-form-card gap-0 overflow-visible border-border/80 shadow-sm py-0">
                    <CardHeader className="border-b pt-6">
                        <CardTitle>Emergency Transaction</CardTitle>
                        <CardDescription>
                            Enter emergency details and add items by searching item code or description.
                        </CardDescription>
                    </CardHeader>
                    <CardContent className="pt-5 space-y-5">
                        <div className="rounded-lg border bg-muted/10 p-3 space-y-4 mb-4">
                            <p className="text-xs font-semibold uppercase tracking-wide text-muted-foreground">Ticket Details</p>
                            <div className="grid gap-4 lg:grid-cols-2">
                                <div className="grid gap-1.5">
                                    <p className="text-sm font-medium">Date</p>
                                    <Input
                                        type="date"
                                        value={header.date}
                                        onChange={(event) => setHeader((prev) => ({ ...prev, date: event.target.value }))}
                                        className="bg-background"
                                    />
                                </div>
                                <div className="grid gap-1.5">
                                    <p className="text-sm font-medium">Requisitioner</p>
                                    <Input
                                        value={header.requisitioner}
                                        onChange={(event) => setHeader((prev) => ({ ...prev, requisitioner: event.target.value }))}
                                        placeholder="Enter requisitioner"
                                        className="bg-background"
                                    />
                                </div>
                                <div className="grid gap-1.5">
                                    <p className="text-sm font-medium">Rel #</p>
                                    <Input
                                        value={header.relNumber}
                                        onChange={(event) => setHeader((prev) => ({ ...prev, relNumber: event.target.value }))}
                                        placeholder="Enter rel #"
                                        className="bg-background"
                                    />
                                </div>
                                <div className="grid gap-1.5">
                                    <p className="text-sm font-medium">Purpose</p>
                                    <Input
                                        value={header.purpose}
                                        onChange={(event) => setHeader((prev) => ({ ...prev, purpose: event.target.value }))}
                                        placeholder="Enter purpose"
                                        className="bg-background"
                                    />
                                </div>
                            </div>

                            <div className="grid gap-2">
                                <p className="text-xs font-semibold uppercase tracking-wide text-muted-foreground">Add Item</p>
                                <div className="relative">
                                    <Search className="pointer-events-none absolute left-3 top-3.5 h-4 w-4 text-muted-foreground" />
                                    <Input
                                        value={itemSearch}
                                        onChange={(event) => setItemSearch(event.target.value)}
                                        onKeyDown={handleItemSearchKeyDown}
                                        placeholder="Search item code or description"
                                        className="pl-9 bg-background"
                                    />
                                    {itemSearch.trim() ? (
                                        <div className="absolute left-0 right-0 top-full z-50 mt-1 overflow-hidden rounded-md border bg-background shadow-md">
                                            {isCatalogLoading ? (
                                                <p className="px-3 py-2 text-sm text-muted-foreground">Loading items...</p>
                                            ) : lookupError ? (
                                                <p className="px-3 py-2 text-sm text-destructive">{lookupError}</p>
                                            ) : filteredCatalog.length === 0 ? (
                                                <p className="px-3 py-2 text-sm text-muted-foreground">No items found.</p>
                                            ) : (
                                                <div className="max-h-72 overflow-y-auto divide-y">
                                                    {filteredCatalog.map((catalog, index) => {
                                                        const isAdded = items.some((row) => row.item_id === String(catalog.id));
                                                        const isActive = index === activeSearchIndex;
                                                        return (
                                                            <button
                                                                key={catalog.id}
                                                                type="button"
                                                                className={`flex w-full items-start justify-between gap-3 px-3 py-2 text-left hover:bg-muted/60 ${
                                                                    isActive ? "bg-muted/60" : ""
                                                                }`}
                                                                onClick={() => addItemFromCatalog(catalog)}
                                                                onMouseEnter={() => setActiveSearchIndex(index)}
                                                                disabled={isAdded}
                                                            >
                                                                <div className="grid gap-0.5">
                                                                    <span className="font-medium text-sm">{catalog.item_code ?? "-"}</span>
                                                                    <span className="text-xs text-muted-foreground">
                                                                        {catalog.description ?? "-"}
                                                                    </span>
                                                                </div>
                                                                {isAdded ? (
                                                                    <span className="inline-flex items-center gap-1 text-xs text-emerald-600">
                                                                        <Check className="h-3.5 w-3.5" /> Added
                                                                    </span>
                                                                ) : (
                                                                    <span className="inline-flex items-center gap-1 text-xs font-medium">
                                                                        <Plus className="h-3.5 w-3.5" /> Add
                                                                    </span>
                                                                )}
                                                            </button>
                                                        );
                                                    })}
                                                </div>
                                            )}
                                        </div>
                                    ) : null}
                                </div>
                            </div>

                            <p className="text-xs font-semibold uppercase tracking-wide text-muted-foreground">Ticket Items</p>
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
                                            <TableHead>Remarks</TableHead>
                                            <TableHead>Deduct From</TableHead>
                                            <TableHead className="w-16 text-right">Action</TableHead>
                                        </TableRow>
                                    </TableHeader>
                                    <TableBody>
                                        {items.length === 0 ? (
                                            <TableRow>
                                                <TableCell colSpan={10} className="py-8 text-center text-sm text-muted-foreground">
                                                    No items added yet.
                                                </TableCell>
                                            </TableRow>
                                        ) : (
                                            items.map((item, index) => {
                                                const availability = availabilityMap[item.id];
                                                const availableText =
                                                    availability?.endingQty != null ? `Ending: ${availability.endingQty}` : null;
                                                const bufferText =
                                                    availability?.bufferStock != null ? `Buffer: ${availability.bufferStock}` : null;

                                                const colorClass =
                                                    availability?.status === "in_stock"
                                                        ? "text-emerald-600"
                                                        : availability?.status === "insufficient"
                                                            ? "text-destructive"
                                                            : availability?.status === "missing"
                                                                ? "text-amber-600"
                                                                : "text-muted-foreground";

                                                return (
                                                    <TableRow key={item.id}>
                                                        <TableCell className="text-center text-muted-foreground">{index + 1}</TableCell>
                                                        <TableCell className="min-w-[170px]">
                                                            <div className="grid gap-1">
                                                                <span className="font-medium">{item.item_code || "-"}</span>
                                                                <span className={`text-xs ${colorClass}`}>
                                                                    {availability?.status === "in_stock"
                                                                        ? "In stock"
                                                                        : availability?.status === "insufficient"
                                                                            ? "Insufficient"
                                                                            : availability?.status === "missing"
                                                                                ? "No record"
                                                                                : ""}
                                                                    {availableText ? ` ${availableText}` : ""}
                                                                    {bufferText ? ` | ${bufferText}` : ""}
                                                                </span>
                                                            </div>
                                                        </TableCell>
                                                        <TableCell className="min-w-[220px] whitespace-normal">{item.particulars || "-"}</TableCell>
                                                        <TableCell>{item.unit || "-"}</TableCell>
                                                        <TableCell className="text-right">{formatDecimal(item.unit_cost)}</TableCell>
                                                        <TableCell className="min-w-[100px]">
                                                            <Input
                                                                inputMode="decimal"
                                                                value={item.qty ?? ""}
                                                                onChange={(event) => updateRow(item.id, { qty: parseNumber(event.target.value) })}
                                                                placeholder="0"
                                                                className="text-right"
                                                            />
                                                        </TableCell>
                                                        <TableCell className="text-right">{formatDecimal(item.total_cost)}</TableCell>
                                                        <TableCell className="min-w-[170px]">
                                                            <Input
                                                                value={item.remarks}
                                                                onChange={(event) => updateRow(item.id, { remarks: event.target.value })}
                                                                placeholder="Remarks"
                                                            />
                                                        </TableCell>
                                                        <TableCell className="min-w-[160px]">
                                                            <Select
                                                                value={item.deduct_from}
                                                                onValueChange={(value) =>
                                                                    updateRow(item.id, {
                                                                        deduct_from: value as "ending_qty" | "buffer_stock",
                                                                    })
                                                                }
                                                            >
                                                                <SelectTrigger>
                                                                    <SelectValue placeholder="Deduct from" />
                                                                </SelectTrigger>
                                                                <SelectContent>
                                                                    <SelectItem value="ending_qty">Ending Qty</SelectItem>
                                                                    <SelectItem value="buffer_stock">Buffer Stock</SelectItem>
                                                                </SelectContent>
                                                            </Select>
                                                        </TableCell>
                                                        <TableCell className="text-right">
                                                            <Button
                                                                type="button"
                                                                variant="ghost"
                                                                size="icon"
                                                                onClick={() => removeRow(item.id)}
                                                                aria-label="Remove row"
                                                            >
                                                                <Trash2 className="h-4 w-4 text-destructive" />
                                                            </Button>
                                                        </TableCell>
                                                    </TableRow>
                                                );
                                            })
                                        )}

                                        {items.length > 0 ? (
                                            <TableRow>
                                                <TableCell className="text-center text-sm font-semibold text-muted-foreground">-</TableCell>
                                                <TableCell />
                                                <TableCell />
                                                <TableCell />
                                                <TableCell className="text-right text-sm font-semibold">Total:</TableCell>
                                                <TableCell className="text-right text-sm font-semibold">
                                                    {Number.isFinite(totalQty) ? totalQty : "-"}
                                                </TableCell>
                                                <TableCell className="text-right text-sm font-semibold">{formatDecimal(totalCost)}</TableCell>
                                                <TableCell />
                                                <TableCell />
                                                <TableCell />
                                            </TableRow>
                                        ) : null}
                                    </TableBody>
                                </Table>
                            </div>
                        </div>
                    </CardContent>
                    <CardFooter className="border-t px-0 py-4 !pt-4">
                        <div className="flex w-full items-center justify-between px-6">
                            <p className="text-xs text-muted-foreground">
                                {isSubmitting ? "Saving emergency..." : "Review values before saving."}
                            </p>
                            <div className="flex justify-end gap-2">
                                <Button
                                    type="button"
                                    variant="outline"
                                    onClick={() => go({ to: "/emergency", type: "replace" })}
                                    disabled={isSubmitting}
                                >
                                    Cancel
                                </Button>
                                <Button type="button" onClick={handleSubmit} disabled={isSubmitting}>
                                    {isSubmitting ? "Saving..." : "Add Emergency"}
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
                        <DialogTitle>Unable to save emergency</DialogTitle>
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

export default EmergencyMovementCreatePage;

import MultiUploadWidget from "@/components/multi-upload-widget";
import { CreateView, CreateViewHeader } from "@/components/refine-ui/views/create-view";
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
import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardFooter, CardHeader, CardTitle } from "@/components/ui/card";
import { Dialog, DialogContent, DialogDescription, DialogFooter, DialogHeader, DialogTitle } from "@/components/ui/dialog";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from "@/components/ui/table";
import { Textarea } from "@/components/ui/textarea";
import { supabaseClient } from "@/providers/supabase-client";
import { useGetIdentity, useGo, useInvalidate, useNotification } from "@refinedev/core";
import { CheckCircle2, ChevronDown, ChevronRight, Loader2, XCircle } from "lucide-react";
import { Fragment, useEffect, useMemo, useState } from "react";
import * as XLSX from "xlsx";
import {
    EMPTY_HEADER,
    formatC2,
    formatDecimal,
    MaterialChargeTicketHeader,
    MaterialChargeTicketItem,
    normalizeRows,
    parseCsvRows,
    parseRowsToTicket,
    sumNumbers,
} from "./create";



type BatchTicketStatus = "ready" | "saving" | "saved" | "error";

type BatchTicket = {
    id: string;
    fileName: string;
    header: MaterialChargeTicketHeader;
    items: MaterialChargeTicketItem[];
    status: BatchTicketStatus;
    error: string | null;
};

const createTicketId = (file: File) => `batch-mct-${file.name}-${file.size}-${file.lastModified}`;

const parseFileToTicket = async (file: File): Promise<BatchTicket> => {
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

        const sheetRows = XLSX.utils.sheet_to_json(sheet, {
            header: 1,
            defval: "",
            raw: false,
        }) as Array<Array<string | number | null>>;
        rows = normalizeRows(sheetRows);
    } else {
        throw new Error(`${file.name}: unsupported file format.`);
    }

    const { header, items } = parseRowsToTicket(rows);

    return {
        id: createTicketId(file),
        fileName: file.name,
        header,
        items,
        status: "ready",
        error: null,
    };
};

const buildHeaderPayload = (header: MaterialChargeTicketHeader) => ({
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

const buildItemsPayload = (items: MaterialChargeTicketItem[]) =>
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

const parseErrorCodes = (message: string, prefix: string) => {
    if (!message.startsWith(prefix)) return [];
    return message
        .slice(prefix.length)
        .split(",")
        .map((code) => code.trim())
        .filter(Boolean);
};

const getSaveError = (message: string) => {
    const duplicate = parseErrorCodes(message, "duplicate_mct:");
    if (duplicate.length > 0) {
        return `Duplicate MCT/Rel # detected: ${duplicate.join(", ")}`;
    }

    const missingItems = parseErrorCodes(message, "missing_item_codes:");
    if (missingItems.length > 0) {
        return `Item code not found: ${missingItems.join(", ")}`;
    }

    const insufficientInventory = parseErrorCodes(message, "insufficient_inventory:");
    if (insufficientInventory.length > 0) {
        return `Insufficient inventory for item code: ${insufficientInventory.join(", ")}`;
    }

    return message || "Unable to save MCT.";
};

const getTicketLabel = (ticket: BatchTicket) => ticket.header.mctRelNumber || ticket.fileName;

type BatchUploadMctPageProps = {
    files?: File[];
    onFilesChange?: (next: File[]) => void;
};

const BatchUploadMctPage = ({ files: controlledFiles }: BatchUploadMctPageProps) => {
    const [files, setFiles] = useState<File[]>([]);


    useEffect(() => {
        if (controlledFiles) {
            setFiles(controlledFiles);
        }
    }, [controlledFiles]);

    const [tickets, setTickets] = useState<BatchTicket[]>([]);
    const [expandedTicketIds, setExpandedTicketIds] = useState<Set<string>>(new Set());
    const [parseStatus, setParseStatus] = useState<string | null>(null);
    const [parseError, setParseError] = useState<string | null>(null);
    const [validationErrors, setValidationErrors] = useState<string[]>([]);
    const [errorDialogOpen, setErrorDialogOpen] = useState(false);
    const [confirmOpen, setConfirmOpen] = useState(false);
    const [pendingCreateInventoryIds, setPendingCreateInventoryIds] = useState<string[]>([]);
    const [isSubmitting, setIsSubmitting] = useState(false);
    const { data: identity } = useGetIdentity<{ id?: string | number }>();
    const { open } = useNotification();
    const invalidate = useInvalidate();
    const go = useGo();

    useEffect(() => {
        if (files.length === 0) {
            setTickets([]);
            setExpandedTicketIds(new Set());
            setParseStatus(null);
            setParseError(null);
            return;
        }

        let isActive = true;

        const parseFiles = async () => {
            setParseError(null);
            setParseStatus(`Parsing ${files.length} file${files.length === 1 ? "" : "s"}...`);

            const parsedTickets: BatchTicket[] = [];
            const errors: string[] = [];

            for (const file of files) {
                try {
                    parsedTickets.push(await parseFileToTicket(file));
                } catch (error) {
                    errors.push(error instanceof Error ? error.message : `${file.name}: unable to parse file.`);
                }
            }

            if (!isActive) return;

            setTickets(parsedTickets);
            setExpandedTicketIds(new Set(parsedTickets.slice(0, 1).map((ticket) => ticket.id)));
            setParseStatus(
                parsedTickets.length > 0
                    ? `Parsed ${parsedTickets.length} MCT document${parsedTickets.length === 1 ? "" : "s"}.`
                    : null
            );
            setParseError(errors.length > 0 ? errors.join(" ") : null);
        };

        void parseFiles();

        return () => {
            isActive = false;
        };
    }, [files]);

    const totals = useMemo(() => {
        const itemCount = tickets.reduce((acc, ticket) => acc + ticket.items.length, 0);
        const qty = tickets.reduce((acc, ticket) => acc + sumNumbers(ticket.items.map((item) => item.qty)), 0);
        const cost = tickets.reduce((acc, ticket) => acc + sumNumbers(ticket.items.map((item) => item.total_cost)), 0);
        return { itemCount, qty, cost };
    }, [tickets]);

    const updateTicketStatus = (ticketId: string, status: BatchTicketStatus, error: string | null = null) => {
        setTickets((currentTickets) =>
            currentTickets.map((ticket) =>
                ticket.id === ticketId ? { ...ticket, status, error } : ticket
            )
        );
    };

    const updateItemDeductFrom = (
        ticketId: string,
        itemId: string,
        deductFrom: "ending_qty" | "buffer_stock"
    ) => {
        setTickets((currentTickets) =>
            currentTickets.map((ticket) =>
                ticket.id === ticketId
                    ? {
                        ...ticket,
                        items: ticket.items.map((item) =>
                            item.id === itemId ? { ...item, deduct_from: deductFrom } : item
                        ),
                    }
                    : ticket
            )
        );
    };

    const toggleExpanded = (ticketId: string) => {
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

    const validateTickets = () => {
        const errors: string[] = [];

        if (tickets.length === 0) {
            errors.push("No MCT documents detected. Upload files before saving.");
        }

        tickets.forEach((ticket, ticketIndex) => {
            const label = getTicketLabel(ticket) || `Document ${ticketIndex + 1}`;
            if (ticket.items.length === 0) {
                errors.push(`${label}: no item rows detected.`);
            }

            ticket.items.forEach((item, itemIndex) => {
                if (!item.item_code?.trim()) {
                    errors.push(`${label}, row ${itemIndex + 1}: missing item code.`);
                }
                if (item.qty == null || Number.isNaN(item.qty) || item.qty <= 0) {
                    errors.push(`${label}, row ${itemIndex + 1}: quantity must be greater than 0.`);
                }
                if (!item.deduct_from) {
                    errors.push(`${label}, row ${itemIndex + 1}: deduct from is required.`);
                }
            });
        });

        setValidationErrors(errors);
        setErrorDialogOpen(errors.length > 0);
        return errors.length === 0;
    };

    const saveTickets = async (ticketIds: string[], createMissingInventory: boolean) => {
        const savedIds: string[] = [];
        const missingInventoryIds: string[] = [];
        const errors: string[] = [];

        for (const ticket of tickets.filter((currentTicket) => ticketIds.includes(currentTicket.id))) {
            updateTicketStatus(ticket.id, "saving");

            const { error } = await supabaseClient.rpc("create_mct_transaction", {
                p_header: buildHeaderPayload(ticket.header),
                p_items: buildItemsPayload(ticket.items),
                p_create_missing_inventory: createMissingInventory,
                p_created_by: identity?.id ? String(identity.id) : null,
            });

            if (!error) {
                savedIds.push(ticket.id);
                updateTicketStatus(ticket.id, "saved");
                continue;
            }

            const message = error.message ?? "Unable to save MCT.";
            const missingInventory = parseErrorCodes(message, "missing_inventory:");

            if (!createMissingInventory && missingInventory.length > 0) {
                missingInventoryIds.push(ticket.id);
                updateTicketStatus(
                    ticket.id,
                    "error",
                    `Missing inventory record for item code: ${missingInventory.join(", ")}`
                );
                continue;
            }

            const displayError = getSaveError(message);
            errors.push(`${getTicketLabel(ticket)}: ${displayError}`);
            updateTicketStatus(ticket.id, "error", displayError);
        }

        return { savedIds, missingInventoryIds, errors };
    };

    const handleSaveBatch = async () => {
        if (isSubmitting || !validateTickets()) return;

        setIsSubmitting(true);
        setValidationErrors([]);
        setPendingCreateInventoryIds([]);

        try {
            const readyTicketIds = tickets
                .filter((ticket) => ticket.status === "ready" || ticket.status === "error")
                .map((ticket) => ticket.id);
            const result = await saveTickets(readyTicketIds, false);

            if (result.missingInventoryIds.length > 0) {
                setPendingCreateInventoryIds(result.missingInventoryIds);
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
                    message: "MCT batch saved",
                    description: `${result.savedIds.length} MCT document${result.savedIds.length === 1 ? "" : "s"} saved.`,
                });
                go({ to: "/mct", type: "replace" });
            }
        } catch (error) {
            setValidationErrors([error instanceof Error ? error.message : "Unable to save MCT batch."]);
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
            const result = await saveTickets(pendingCreateInventoryIds, true);
            setConfirmOpen(false);
            setPendingCreateInventoryIds([]);

            if (result.savedIds.length > 0) {
                invalidate({ resource: "mcts", invalidates: ["list"] });
                invalidate({ resource: "mct_items", invalidates: ["list"] });
            }

            if (result.errors.length > 0) {
                setValidationErrors(result.errors);
                setErrorDialogOpen(true);
                return;
            }

            open?.({
                type: "success",
                message: "MCT batch saved",
                description: "Batch upload completed.",
            });
            go({ to: "/mct", type: "replace" });
        } catch (error) {
            setValidationErrors([error instanceof Error ? error.message : "Unable to save MCT batch."]);
            setErrorDialogOpen(true);
        } finally {
            setIsSubmitting(false);
        }
    };

    return (
        <CreateView className="item-view">
            <CreateViewHeader title="Batch Upload MCTs" />
            <div className="my-4 flex items-center">
                <Card className="w-full max-w-7xl mx-auto item-form-card gap-0 overflow-hidden border-border/80 shadow-sm py-0">
                    <CardHeader className="border-b pt-6">
                        <CardTitle>Batch Material Charge Tickets</CardTitle>
                        <CardDescription>
                            Upload multiple Excel or CSV files and save each parsed document as an MCT.
                        </CardDescription>
                    </CardHeader>
                    <CardContent className="pt-5 space-y-5">
                        <MultiUploadWidget value={files} onFilesChange={setFiles} disabled={isSubmitting} />
                        {parseStatus ? <p className="text-sm text-muted-foreground">{parseStatus}</p> : null}
                        {parseError ? <p className="text-sm text-destructive">{parseError}</p> : null}

                        <div className="grid gap-3 rounded-lg border bg-muted/10 p-3 sm:grid-cols-3">
                            <div className="rounded-md border bg-background px-3 py-2">
                                <p className="text-xs font-semibold uppercase tracking-wide text-muted-foreground">Documents</p>
                                <p className="text-lg font-semibold">{tickets.length}</p>
                            </div>
                            <div className="rounded-md border bg-background px-3 py-2">
                                <p className="text-xs font-semibold uppercase tracking-wide text-muted-foreground">Items</p>
                                <p className="text-lg font-semibold">{totals.itemCount}</p>
                            </div>
                            <div className="rounded-md border bg-background px-3 py-2">
                                <p className="text-xs font-semibold uppercase tracking-wide text-muted-foreground">Total Cost</p>
                                <p className="text-lg font-semibold">{formatDecimal(totals.cost)}</p>
                            </div>
                        </div>

                        <div className="overflow-x-auto rounded-md border bg-background">
                            <Table>
                                <TableHeader>
                                    <TableRow>
                                        <TableHead className="w-12" />
                                        <TableHead>MCT/Rel #</TableHead>
                                        <TableHead>File</TableHead>
                                        <TableHead>Requisitioner</TableHead>
                                        <TableHead className="text-right">Items</TableHead>
                                        <TableHead className="text-right">Qty</TableHead>
                                        <TableHead className="text-right">Total Cost</TableHead>
                                        <TableHead>Status</TableHead>
                                    </TableRow>
                                </TableHeader>
                                <TableBody>
                                    {tickets.length === 0 ? (
                                        <TableRow>
                                            <TableCell colSpan={8} className="py-8 text-center text-sm text-muted-foreground">
                                                No MCT documents parsed yet.
                                            </TableCell>
                                        </TableRow>
                                    ) : (
                                        tickets.map((ticket) => {
                                            const header = ticket.header ?? EMPTY_HEADER;
                                            const isExpanded = expandedTicketIds.has(ticket.id);
                                            const totalQty = sumNumbers(ticket.items.map((item) => item.qty));
                                            const totalCost = sumNumbers(ticket.items.map((item) => item.total_cost));

                                            return (
                                                <Fragment key={ticket.id}>
                                                    <TableRow>
                                                        <TableCell>
                                                            <Button
                                                                type="button"
                                                                variant="ghost"
                                                                size="icon"
                                                                className="h-8 w-8"
                                                                onClick={() => toggleExpanded(ticket.id)}
                                                            >
                                                                {isExpanded ? (
                                                                    <ChevronDown className="h-4 w-4" />
                                                                ) : (
                                                                    <ChevronRight className="h-4 w-4" />
                                                                )}
                                                            </Button>
                                                        </TableCell>
                                                        <TableCell className="font-medium">{header.mctRelNumber || "-"}</TableCell>
                                                        <TableCell className="max-w-[240px] truncate">{ticket.fileName}</TableCell>
                                                        <TableCell>{header.requisitioner || "-"}</TableCell>
                                                        <TableCell className="text-right">{ticket.items.length}</TableCell>
                                                        <TableCell className="text-right">{totalQty}</TableCell>
                                                        <TableCell className="text-right">{formatDecimal(totalCost)}</TableCell>
                                                        <TableCell>
                                                            {ticket.status === "saving" ? (
                                                                <Badge variant="secondary" className="gap-1">
                                                                    <Loader2 className="h-3 w-3 animate-spin" />
                                                                    Saving
                                                                </Badge>
                                                            ) : ticket.status === "saved" ? (
                                                                <Badge className="gap-1 bg-emerald-600 hover:bg-emerald-600">
                                                                    <CheckCircle2 className="h-3 w-3" />
                                                                    Saved
                                                                </Badge>
                                                            ) : ticket.status === "error" ? (
                                                                <Badge variant="destructive" className="gap-1">
                                                                    <XCircle className="h-3 w-3" />
                                                                    Error
                                                                </Badge>
                                                            ) : (
                                                                <Badge variant="outline">Ready</Badge>
                                                            )}
                                                        </TableCell>
                                                    </TableRow>
                                                    {isExpanded ? (
                                                        <TableRow>
                                                            <TableCell colSpan={8} className="bg-muted/20 p-0">
                                                                <div className="rounded-lg border bg-muted/10 p-3 space-y-4 m-4">
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
                                                                                    ticket.items.map((item, index) => (
                                                                                        <TableRow key={item.id}>
                                                                                            <TableCell className="text-center text-muted-foreground">{index + 1}</TableCell>
                                                                                            <TableCell className="font-medium">{item.item_code || "-"}</TableCell>
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
                                                                                                        updateItemDeductFrom(
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
                                                                                    ))
                                                                                )}
                                                                                {ticket.items.length > 0 ? (
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
                                                                            <Textarea value={header.purpose} readOnly className="min-h-24 bg-background" />
                                                                        </div>
                                                                        <div className="grid gap-1.5">
                                                                            <p className="text-xs font-semibold uppercase tracking-wide text-muted-foreground">Notes / SR #</p>
                                                                            <Textarea value={header.notes} readOnly className="h-24 resize-y overflow-auto bg-background" />
                                                                        </div>
                                                                    </div>
                                                                </div>
                                                            </TableCell>
                                                        </TableRow>
                                                    ) : null}
                                                </Fragment>
                                            );
                                        })
                                    )}
                                </TableBody>
                            </Table>
                        </div>
                    </CardContent>
                    <CardFooter className="border-t px-0 py-4 !pt-4">
                        <div className="flex w-full items-center justify-between px-6">
                            <p className="text-xs text-muted-foreground">
                                {isSubmitting ? "Saving MCT batch..." : "Review parsed documents before saving."}
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
                                <Button type="button" onClick={handleSaveBatch} disabled={isSubmitting || tickets.length === 0}>
                                    {isSubmitting ? "Saving..." : "Save Batch"}
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
                            Some MCTs do not have inventory records for the current month. Add records now and continue saving?
                        </AlertDialogDescription>
                    </AlertDialogHeader>
                    <div className="grid gap-3 px-6 py-6 text-sm">
                        {tickets
                            .filter((ticket) => pendingCreateInventoryIds.includes(ticket.id))
                            .map((ticket) => (
                                <div key={ticket.id} className="flex items-center justify-between gap-4">
                                    <span className="font-medium">{getTicketLabel(ticket)}</span>
                                    <span className="text-muted-foreground">{ticket.error}</span>
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
                        <DialogTitle>Unable to save MCT batch</DialogTitle>
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

export default BatchUploadMctPage;

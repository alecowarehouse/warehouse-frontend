import { cn } from "@/lib/utils";
import { FileSpreadsheet, UploadCloud, X } from "lucide-react";
import { useRef, useState } from "react";

type MultiUploadWidgetProps = {
    value?: File[];
    onFilesChange?: (files: File[]) => void;
    maxSizeMb?: number;
    className?: string;
    disabled?: boolean;
};

const ACCEPTED_EXCEL_FILE_TYPES =
    ".xlsx,.xls,.csv,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel,text/csv";

const getFileKey = (file: File) => `${file.name}-${file.size}-${file.lastModified}`;

const MultiUploadWidget = ({
    value = [],
    onFilesChange,
    maxSizeMb = 5,
    className,
    disabled = false,
}: MultiUploadWidgetProps) => {
    const inputRef = useRef<HTMLInputElement>(null);
    const [isDragActive, setIsDragActive] = useState(false);
    const [error, setError] = useState<string | null>(null);

    const validateFile = (file: File) => {
        const extension = file.name.split(".").pop()?.toLowerCase();
        const allowedExtensions = new Set(["xlsx", "xls", "csv"]);
        const maxSizeInBytes = maxSizeMb * 1024 * 1024;

        if (!extension || !allowedExtensions.has(extension)) {
            return `${file.name}: only .xlsx, .xls, and .csv files are allowed.`;
        }

        if (file.size > maxSizeInBytes) {
            return `${file.name}: file is too large. Max size is ${maxSizeMb}MB.`;
        }

        return null;
    };

    const addFiles = (selectedFiles: FileList | File[]) => {
        const incomingFiles = Array.from(selectedFiles);
        const validationError = incomingFiles.map(validateFile).find(Boolean);

        if (validationError) {
            setError(validationError);
            if (inputRef.current) {
                inputRef.current.value = "";
            }
            return;
        }

        const existingKeys = new Set(value.map(getFileKey));
        const nextFiles = [
            ...value,
            ...incomingFiles.filter((file) => !existingKeys.has(getFileKey(file))),
        ];

        setError(null);
        onFilesChange?.(nextFiles);

        if (inputRef.current) {
            inputRef.current.value = "";
        }
    };

    const removeFile = (file: File) => {
        const key = getFileKey(file);
        onFilesChange?.(value.filter((currentFile) => getFileKey(currentFile) !== key));
    };

    return (
        <div className={cn("space-y-3", className)}>
            <input
                ref={inputRef}
                type="file"
                className="hidden"
                accept={ACCEPTED_EXCEL_FILE_TYPES}
                disabled={disabled}
                multiple
                onChange={(event) => {
                    const selectedFiles = event.target.files;
                    if (selectedFiles?.length) {
                        addFiles(selectedFiles);
                    }
                }}
            />

            <button
                type="button"
                disabled={disabled}
                onClick={() => inputRef.current?.click()}
                onDragOver={(event) => {
                    event.preventDefault();
                    if (!disabled) {
                        setIsDragActive(true);
                    }
                }}
                onDragLeave={() => setIsDragActive(false)}
                onDrop={(event) => {
                    event.preventDefault();
                    setIsDragActive(false);
                    if (disabled || !event.dataTransfer.files.length) {
                        return;
                    }
                    addFiles(event.dataTransfer.files);
                }}
                className={cn(
                    "w-full cursor-pointer rounded-lg border border-dashed px-4 py-7 text-center transition-colors",
                    "bg-muted/40 hover:bg-muted/60",
                    "focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-ring/50",
                    isDragActive && "border-primary bg-primary/5",
                    error && "border-destructive/80 bg-destructive/5",
                    disabled && "cursor-not-allowed opacity-60"
                )}
            >
                <div className="flex flex-col items-center gap-2">
                    <UploadCloud className="h-8 w-8 text-blue-500" />
                    <p className="text-sm font-semibold text-blue-600">Drag here or click to upload MCT files</p>
                    <p className="text-xs text-muted-foreground">
                        XLSX, XLS, CSV up to {maxSizeMb}MB each
                    </p>
                </div>
            </button>

            {value.length > 0 ? (
                <div className="grid gap-2">
                    {value.map((file) => (
                        <div
                            key={getFileKey(file)}
                            className="flex items-center justify-between gap-2 overflow-hidden rounded-md border bg-background px-3 py-2"
                        >
                            <div className="flex min-w-0 flex-1 items-center gap-2">
                                <FileSpreadsheet className="h-4 w-4 text-primary" />
                                <span className="truncate text-sm">{file.name}</span>
                            </div>
                            <button
                                type="button"
                                disabled={disabled}
                                onClick={() => removeFile(file)}
                                className="inline-flex shrink-0 items-center rounded p-1 text-muted-foreground transition-colors hover:bg-muted hover:text-foreground disabled:pointer-events-none disabled:opacity-50"
                                aria-label={`Remove ${file.name}`}
                            >
                                <X className="h-4 w-4" />
                            </button>
                        </div>
                    ))}
                </div>
            ) : null}

            {error ? <p className="text-sm text-destructive">{error}</p> : null}
        </div>
    );
};

export default MultiUploadWidget;

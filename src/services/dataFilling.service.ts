import { Request } from "express";
import XLSX from "xlsx";
import fs from "fs";
import path from "path";
import { v4 as uuidv4 } from "uuid";
import PMSHeaderTemplate from "../models/PMSHeader.model";
import { EmployeeModel } from "../models/Employee.model";
import { EmployeeDataModel } from "../models/EmployeeData.model";
import { ApiError } from "../utils/ApiError";

// Helper function to convert column letter to index
const columnLetterToIndex = (letter: string): number => {
    let index = 0;
    for (let i = 0; i < letter.length; i++) {
        index = index * 26 + letter.charCodeAt(i) - 64;
    }
    return index - 1;
};

// Helper function to parse cell address
const parseCellAddress = (address: string): { row: number; col: number } => {
    const match = address.match(/^([A-Z]+)(\d+)$/i);
    if (!match) throw new Error(`Invalid cell address: ${address}`);
    return {
        col: columnLetterToIndex(match[1].toUpperCase()),
        row: parseInt(match[2]) - 1
    };
};

// Helper function to evaluate formulas
const evaluateFormula = (
    formula: string,
    rowData: { [key: string]: number | string | null },
    columnMappings: { [key: number]: string },
    currentRow: number
): number | null => {
    try {
        // Replace cell references with actual values
        let evaluableFormula = formula;

        // Match cell references like A1, B2, AB12, $A$1, etc.
        const cellRefRegex = /\$?([A-Z]+)\$?(\d+)/gi;
        let match;

        while ((match = cellRefRegex.exec(formula)) !== null) {
            const colLetter = match[1].toUpperCase();
            const rowNum = parseInt(match[2]) - 1;
            const colIndex = columnLetterToIndex(colLetter);

            // Get the header name for this column
            const headerName = columnMappings[colIndex];

            if (headerName) {
                const value = rowData[headerName];
                const numValue = typeof value === 'number' ? value :
                    (typeof value === 'string' && !isNaN(parseFloat(value)) ? parseFloat(value) : 0);
                evaluableFormula = evaluableFormula.replace(match[0], String(numValue));
            } else {
                evaluableFormula = evaluableFormula.replace(match[0], '0');
            }
        }

        // Handle common Excel functions
        evaluableFormula = evaluableFormula
            .replace(/SUM\(([^)]+)\)/gi, (_, args) => {
                const values = args.split(',').map((v: string) => parseFloat(v.trim()) || 0);
                return String(values.reduce((a: number, b: number) => a + b, 0));
            })
            .replace(/AVERAGE\(([^)]+)\)/gi, (_, args) => {
                const values = args.split(',').map((v: string) => parseFloat(v.trim()) || 0);
                return String(values.reduce((a: number, b: number) => a + b, 0) / values.length);
            })
            .replace(/IF\(([^,]+),([^,]+),([^)]+)\)/gi, (_, condition, trueVal, falseVal) => {
                try {
                    const condResult = eval(condition);
                    return condResult ? String(eval(trueVal)) : String(eval(falseVal));
                } catch {
                    return '0';
                }
            });

        // Evaluate the formula
        const result = eval(evaluableFormula);
        return typeof result === 'number' ? result : null;
    } catch (error) {
        console.error('Formula evaluation error:', error);
        return null;
    }
};

export const DataFillingService = {
    // Get all available templates for selection
    async getTemplatesForSelection() {
        const templates = await PMSHeaderTemplate.find(
            { is_active: true },
            {
                _id: 1,
                template_name: 1,
                description: 1,
                version: 1,
                createdAt: 1,
                column_mappings: 1,
                employee_field_mapping: 1,
                header_row_index: 1,
                data_start_row: 1,
                "metadata.total_rows": 1,
                "metadata.total_columns": 1,
                "metadata.formula_cells": 1
            }
        ).sort({ createdAt: -1 });

        return templates;
    },

    // Get template preview data
    async getTemplatePreview(templateId: string) {
        const template = await PMSHeaderTemplate.findById(templateId);

        if (!template) {
            throw new ApiError(404, "Template not found");
        }

        return {
            _id: template._id,
            template_name: template.template_name,
            description: template.description,
            version: template.version,
            metadata: template.metadata,
            sheet_structure: template.sheet_structure,
            column_mappings: template.column_mappings,
            formula_definitions: template.formula_definitions,
            employee_field_mapping: template.employee_field_mapping,
            header_row_index: template.header_row_index,
            data_start_row: template.data_start_row
        };
    },

    // Download template as Excel file (original file with formulas)
    async downloadTemplate(templateId: string) {
        const template = await PMSHeaderTemplate.findById(templateId);

        if (!template) {
            throw new ApiError(404, "Template not found");
        }

        // If we have the original file buffer, return it
        if (template.file_buffer) {
            return {
                buffer: template.file_buffer,
                filename: template.file_name || `${template.template_name}.xlsx`,
                contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            };
        }

        // Fallback: Create workbook from sheet structure
        const workbook = XLSX.utils.book_new();
        const worksheetData: any[][] = [];

        let maxRow = 0;
        let maxCol = 0;

        template.sheet_structure.forEach((cell: any) => {
            maxRow = Math.max(maxRow, cell.row);
            maxCol = Math.max(maxCol, cell.col);
        });

        for (let r = 0; r <= maxRow; r++) {
            worksheetData[r] = [];
            for (let c = 0; c <= maxCol; c++) {
                worksheetData[r][c] = null;
            }
        }

        template.sheet_structure.forEach((cell: any) => {
            if (cell.formula) {
                worksheetData[cell.row][cell.col] = { f: cell.formula, v: cell.value };
            } else {
                worksheetData[cell.row][cell.col] = cell.value;
            }
        });

        const worksheet = XLSX.utils.aoa_to_sheet(worksheetData);

        // Preserve formulas
        template.sheet_structure.forEach((cell: any) => {
            if (cell.formula) {
                const cellRef = XLSX.utils.encode_cell({ r: cell.row, c: cell.col });
                if (worksheet[cellRef]) {
                    worksheet[cellRef].f = cell.formula;
                }
            }
        });

        XLSX.utils.book_append_sheet(workbook, worksheet, "PMS Data");
        const buffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' });

        return {
            buffer,
            filename: `${template.template_name}.xlsx`,
            contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        };
    },

    // Validate uploaded file against template
    async validateUploadedFile(filePath: string, templateId: string) {
        const template = await PMSHeaderTemplate.findById(templateId);
        if (!template) {
            throw new ApiError(404, "Template not found");
        }

        const workbook = XLSX.readFile(filePath, { cellFormula: true });
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        const range = XLSX.utils.decode_range(sheet["!ref"]!);

        const validationErrors: string[] = [];
        const warnings: string[] = [];

        // Check if employee ID column exists
        const employeeIdCol = template.employee_field_mapping?.employeeIdColumn;
        if (employeeIdCol === -1 || employeeIdCol === undefined) {
            validationErrors.push("Template does not have an Employee ID column mapping");
        }

        // Check column count
        if (range.e.c + 1 < (template.metadata?.total_columns || 0)) {
            warnings.push(`Uploaded file has fewer columns than template (${range.e.c + 1} vs ${template.metadata?.total_columns})`);
        }

        // Check headers match
        if (template.header_row_index !== undefined && template.header_row_index >= 0) {
            const expectedHeaders = template.column_mappings?.map(cm => cm.headerName) || [];
            const actualHeaders: string[] = [];

            for (let c = range.s.c; c <= range.e.c; c++) {
                const cellAddress = XLSX.utils.encode_cell({ r: template.header_row_index, c });
                const cell = sheet[cellAddress];
                if (cell && cell.v !== undefined) {
                    actualHeaders.push(String(cell.v).trim());
                }
            }

            // Check for missing required headers
            expectedHeaders.forEach((expected, idx) => {
                if (!actualHeaders.some(actual =>
                    actual.toLowerCase() === expected.toLowerCase()
                )) {
                    warnings.push(`Expected header "${expected}" not found in uploaded file`);
                }
            });
        }

        return {
            isValid: validationErrors.length === 0,
            errors: validationErrors,
            warnings,
            rowCount: range.e.r - (template.data_start_row || 1) + 1,
            columnCount: range.e.c + 1
        };
    },

    // Upload and process filled Excel file with formula calculation
    async uploadFilledData(req: Request) {
        if (!req.file) {
            throw new ApiError(400, "No file uploaded");
        }

        const { templateId } = req.body;
        const uploadedBy = (req as any).user?.id;

        if (!templateId) {
            throw new ApiError(400, "Template ID is required. Please select a template first.");
        }

        if (!uploadedBy) {
            throw new ApiError(401, "User authentication required");
        }

        // Verify template exists
        const template = await PMSHeaderTemplate.findById(templateId);
        if (!template) {
            fs.unlinkSync(req.file.path);
            throw new ApiError(404, "Template not found");
        }

        const filePath = req.file.path;
        const uploadBatchId = uuidv4();
        const results: any[] = [];
        const errors: any[] = [];

        try {
            // Validate file against template
            const validation = await this.validateUploadedFile(filePath, templateId);
            if (!validation.isValid) {
                throw new ApiError(400, `File validation failed: ${validation.errors.join(', ')}`);
            }

            // Read uploaded Excel file with formulas
            const workbook = XLSX.readFile(filePath, { cellFormula: true });
            const sheetName = workbook.SheetNames[0];
            const sheet = workbook.Sheets[sheetName];
            const range = XLSX.utils.decode_range(sheet["!ref"]!);

            // Get headers from the file
            const headers: { [key: number]: string } = {};
            const headerRow = template.header_row_index ?? 0;

            for (let c = range.s.c; c <= range.e.c; c++) {
                const cellAddress = XLSX.utils.encode_cell({ r: headerRow, c });
                const cell = sheet[cellAddress];
                if (cell && cell.v !== undefined) {
                    headers[c] = String(cell.v).trim();
                }
            }

            // Create column mappings lookup
            const columnMappings: { [key: number]: string } = headers;

            // Get employee ID column from template
            const employeeIdCol = template.employee_field_mapping?.employeeIdColumn ?? -1;
            let employeeIdColActual = employeeIdCol;

            // If employeeIdCol is -1, try to find it from headers
            if (employeeIdColActual === -1) {
                for (const [colStr, headerName] of Object.entries(headers)) {
                    const lowerHeader = headerName.toLowerCase();
                    if (lowerHeader.includes('employee') &&
                        (lowerHeader.includes('id') || lowerHeader.includes('no') || lowerHeader.includes('number'))) {
                        employeeIdColActual = parseInt(colStr);
                        break;
                    }
                }
            }

            if (employeeIdColActual === -1) {
                throw new ApiError(400, "Could not find Employee ID column in the uploaded file");
            }

            // Get formula definitions from template
            const formulaDefinitions = template.formula_definitions || [];
            const formulaCells = new Map<string, { formula: string; dependencies: string[] }>();

            formulaDefinitions.forEach(fd => {
                formulaCells.set(fd.cellAddress, {
                    formula: fd.formula,
                    dependencies: fd.dependentCells
                });
            });

            // Process each data row
            const dataStartRow = template.data_start_row ?? headerRow + 1;

            for (let r = dataStartRow; r <= range.e.r; r++) {
                // Check if row has data
                let hasData = false;
                for (let c = range.s.c; c <= range.e.c; c++) {
                    const cellAddress = XLSX.utils.encode_cell({ r, c });
                    const cell = sheet[cellAddress];
                    if (cell && cell.v !== undefined && cell.v !== null && cell.v !== '') {
                        hasData = true;
                        break;
                    }
                }

                if (!hasData) continue;

                const employeeIdCell = XLSX.utils.encode_cell({ r, c: employeeIdColActual });
                const employeeIdValue = sheet[employeeIdCell]?.v;

                if (!employeeIdValue) {
                    errors.push({
                        rowNumber: r + 1,
                        error: "Missing Employee ID"
                    });
                    continue;
                }

                // Find employee in database
                const employee = await EmployeeModel.findOne({
                    employeeId: String(employeeIdValue),
                    isActive: true
                });

                if (!employee) {
                    errors.push({
                        rowNumber: r + 1,
                        employeeId: employeeIdValue,
                        error: `Employee with ID ${employeeIdValue} not found`
                    });
                    continue;
                }

                // Create row data object
                const rowData: Record<string, any> = {};
                const calculatedValues: Record<string, any> = {};

                // First pass: collect all raw values
                for (let c = range.s.c; c <= range.e.c; c++) {
                    const cellAddress = XLSX.utils.encode_cell({ r, c });
                    const cell = sheet[cellAddress];
                    const headerName = headers[c];

                    if (headerName) {
                        if (cell) {
                            // If cell has a calculated value from formula
                            if (cell.f) {
                                // Store both the formula and calculated value
                                rowData[headerName] = {
                                    formula: cell.f,
                                    value: cell.v,
                                    calculatedValue: cell.v
                                };
                                calculatedValues[headerName] = cell.v;
                            } else if (cell.v !== undefined && cell.v !== null) {
                                rowData[headerName] = cell.v;
                            }
                        }
                    }
                }

                // Second pass: calculate formulas if needed
                // This is for cases where formulas reference other cells in the same row
                for (const [cellAddr, formulaInfo] of formulaCells.entries()) {
                    const { row: fRow, col: fCol } = parseCellAddress(cellAddr);

                    // Check if this formula applies to the current row pattern
                    if (fRow >= dataStartRow) {
                        const relativeRow = r - dataStartRow;
                        const actualCellAddr = XLSX.utils.encode_cell({
                            r: dataStartRow + relativeRow,
                            c: fCol
                        });

                        const headerName = headers[fCol];
                        if (headerName && !rowData[headerName]?.calculatedValue) {
                            // Calculate formula for this row
                            const adjustedFormula = formulaInfo.formula.replace(
                                /([A-Z]+)(\d+)/gi,
                                (match, col, row) => {
                                    const newRow = parseInt(row) + relativeRow;
                                    return `${col}${newRow}`;
                                }
                            );

                            const calculatedValue = evaluateFormula(
                                adjustedFormula,
                                rowData,
                                columnMappings,
                                r
                            );

                            if (calculatedValue !== null) {
                                rowData[headerName] = {
                                    formula: adjustedFormula,
                                    value: calculatedValue,
                                    calculatedValue
                                };
                                calculatedValues[headerName] = calculatedValue;
                            }
                        }
                    }
                }

                // Save to database
                const employeeData = await EmployeeDataModel.create({
                    templateId,
                    employeeId: employee._id,
                    uploadedBy,
                    uploadBatchId,
                    data: rowData,
                    calculatedData: calculatedValues,
                    rowNumber: r + 1,
                    status: 'validated',
                    validationErrors: []
                });

                results.push({
                    rowNumber: r + 1,
                    employeeId: employeeIdValue,
                    employeeName: employee.name,
                    employeeEmail: employee.email,
                    employeeDesignation: employee.designation,
                    employeeDepartment: employee.department,
                    dataId: employeeData._id,
                    calculatedValues
                });
            }

            // Cleanup uploaded file
            fs.unlinkSync(filePath);

            return {
                uploadBatchId,
                templateName: template.template_name,
                templateId: template._id,
                totalRows: range.e.r - dataStartRow + 1,
                successCount: results.length,
                errorCount: errors.length,
                validationWarnings: validation.warnings,
                results,
                errors
            };

        } catch (error) {
            if (fs.existsSync(filePath)) {
                fs.unlinkSync(filePath);
            }
            throw error;
        }
    },

    // Get upload history (shows all for SuperAdmin)
    async getUploadHistory(userId?: string) {
        const matchCondition: any = {};
        // Optional: Filter by userId if needed in the future
        // if (userId) {
        //     matchCondition.uploadedBy = userId;
        // }

        const history = await EmployeeDataModel.aggregate([
            // Remove user filter to show all uploads for SuperAdmin
            // { $match: matchCondition },
            {
                $group: {
                    _id: "$uploadBatchId",
                    templateId: { $first: "$templateId" },
                    uploadedBy: { $first: "$uploadedBy" },
                    uploadedAt: { $first: "$createdAt" },
                    totalRecords: { $sum: 1 },
                    successCount: {
                        $sum: { $cond: [{ $eq: ["$status", "validated"] }, 1, 0] }
                    },
                    errorCount: {
                        $sum: { $cond: [{ $eq: ["$status", "error"] }, 1, 0] }
                    }
                }
            },
            { $sort: { uploadedAt: -1 } },
            { $limit: 50 }
        ]);

        // Populate template names
        const populatedHistory = await PMSHeaderTemplate.populate(history, {
            path: 'templateId',
            select: 'template_name'
        });

        return populatedHistory;
    },

    // Get data by employee
    async getDataByEmployee(employeeId: string) {
        const data = await EmployeeDataModel.find({ employeeId })
            .populate('templateId', 'template_name')
            .populate('uploadedBy', 'firstName lastName')
            .sort({ createdAt: -1 });

        return data;
    },

    // Get data by batch
    async getDataByBatch(batchId: string) {
        const data = await EmployeeDataModel.find({ uploadBatchId: batchId })
            .populate('employeeId', 'employeeId name email designation department')
            .populate('templateId', 'template_name column_mappings')
            .sort({ rowNumber: 1 });

        return data;
    },

    // Get summary statistics
    async getDataSummary() {
        const summary = await EmployeeDataModel.aggregate([
            {
                $group: {
                    _id: null,
                    totalRecords: { $sum: 1 },
                    totalBatches: { $addToSet: "$uploadBatchId" },
                    validatedCount: {
                        $sum: { $cond: [{ $eq: ["$status", "validated"] }, 1, 0] }
                    },
                    pendingCount: {
                        $sum: { $cond: [{ $eq: ["$status", "pending"] }, 1, 0] }
                    },
                    errorCount: {
                        $sum: { $cond: [{ $eq: ["$status", "error"] }, 1, 0] }
                    }
                }
            },
            {
                $project: {
                    _id: 0,
                    totalRecords: 1,
                    totalBatches: { $size: "$totalBatches" },
                    validatedCount: 1,
                    pendingCount: 1,
                    errorCount: 1
                }
            }
        ]);

        return summary[0] || {
            totalRecords: 0,
            totalBatches: 0,
            validatedCount: 0,
            pendingCount: 0,
            errorCount: 0
        };
    },

    // Get uploaded data with calculated values and template structure
    async getUploadedDataWithCalculations(batchId: string) {
        const data = await EmployeeDataModel.find({ uploadBatchId: batchId })
            .populate('employeeId', 'employeeId name email designation department division')
            .populate('templateId')
            .sort({ rowNumber: 1 });

        if (data.length === 0) {
            return {
                records: [],
                template: null
            };
        }

        const template = data[0].templateId as any;

        const records = data.map(record => {
            const employee = record.employeeId as any;
            return {
                rowNumber: record.rowNumber,
                employee: employee ? {
                    id: employee.employeeId,
                    name: employee.name,
                    email: employee.email,
                    designation: employee.designation,
                    department: employee.department,
                    division: employee.division
                } : null,
                data: record.data,
                calculatedData: record.calculatedData,
                status: record.status,
                createdAt: record.createdAt
            };
        });

        return {
            records,
            template: template ? {
                _id: template._id,
                template_name: template.template_name,
                description: template.description,
                version: template.version,
                column_mappings: template.column_mappings,
                formula_definitions: template.formula_definitions,
                sheet_structure: template.sheet_structure,
                employee_field_mapping: template.employee_field_mapping,
                header_row_index: template.header_row_index,
                data_start_row: template.data_start_row,
                metadata: template.metadata
            } : null
        };
    },

    // Export uploaded data as Excel file
    async exportUploadedDataAsExcel(batchId: string) {
        const data = await EmployeeDataModel.find({ uploadBatchId: batchId })
            .populate('employeeId', 'employeeId name email designation department division')
            .populate('templateId', 'template_name column_mappings')
            .sort({ rowNumber: 1 });

        if (data.length === 0) {
            throw new ApiError(404, "No data found for this batch");
        }

        const template = data[0].templateId as any;
        const workbook = XLSX.utils.book_new();

        // Prepare headers
        const columnMappings = template.column_mappings || [];
        const headers = columnMappings.map((cm: any) => cm.headerName);

        // Prepare data rows
        const rows: any[][] = [];
        rows.push(headers); // Add header row

        data.forEach(record => {
            const employee = record.employeeId as any;
            const row: any[] = [];

            columnMappings.forEach((cm: any) => {
                const headerName = cm.headerName;
                const cellData = record.data[headerName];

                // Handle formula cells
                if (cellData && typeof cellData === 'object' && cellData.calculatedValue !== undefined) {
                    row.push(cellData.calculatedValue);
                } else {
                    row.push(cellData !== undefined ? cellData : '');
                }
            });

            rows.push(row);
        });

        // Create worksheet
        const worksheet = XLSX.utils.aoa_to_sheet(rows);

        // Auto-size columns
        const maxWidths = headers.map((h: string) => h.length);
        rows.slice(1).forEach(row => {
            row.forEach((cell, idx) => {
                const cellLength = String(cell || '').length;
                if (cellLength > maxWidths[idx]) {
                    maxWidths[idx] = cellLength;
                }
            });
        });

        worksheet['!cols'] = maxWidths.map((w: number) => ({ wch: Math.min(w + 2, 50) }));

        // Add worksheet to workbook
        XLSX.utils.book_append_sheet(workbook, worksheet, "Employee Data");

        // Generate buffer
        const buffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' });

        return {
            buffer,
            filename: `${template.template_name}_${batchId.substring(0, 8)}_${new Date().toISOString().split('T')[0]}.xlsx`,
            contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        };
    }
};

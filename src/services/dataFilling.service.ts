import { Request } from "express";
import XLSX from "xlsx";
import fs from "fs";
import path from "path";
import { v4 as uuidv4 } from "uuid";
import PMSHeaderTemplate from "../models/PMSHeader.model";
import { EmployeeModel } from "../models/Employee.model";
import { EmployeeDataModel } from "../models/EmployeeData.model";
import { ApiError } from "../utils/ApiError";

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
                "metadata.total_rows": 1,
                "metadata.total_columns": 1
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
            sheet_structure: template.sheet_structure
        };
    },

    // Download template as Excel file
    async downloadTemplate(templateId: string) {
        const template = await PMSHeaderTemplate.findById(templateId);

        if (!template) {
            throw new ApiError(404, "Template not found");
        }

        // Create a new workbook
        const workbook = XLSX.utils.book_new();

        // Create worksheet from template structure
        const worksheetData: any[][] = [];

        // Find the dimensions of the data
        let maxRow = 0;
        let maxCol = 0;

        template.sheet_structure.forEach((cell: any) => {
            maxRow = Math.max(maxRow, cell.row);
            maxCol = Math.max(maxCol, cell.col);
        });

        // Initialize empty array
        for (let r = 0; r <= maxRow; r++) {
            worksheetData[r] = [];
            for (let c = 0; c <= maxCol; c++) {
                worksheetData[r][c] = null;
            }
        }

        // Fill in the data from template
        template.sheet_structure.forEach((cell: any) => {
            if (cell.formula) {
                worksheetData[cell.row][cell.col] = { f: cell.formula, v: cell.value };
            } else {
                worksheetData[cell.row][cell.col] = cell.value;
            }
        });

        // Create worksheet
        const worksheet = XLSX.utils.aoa_to_sheet(worksheetData);

        // Add formulas back
        template.sheet_structure.forEach((cell: any) => {
            if (cell.formula) {
                const cellRef = XLSX.utils.encode_cell({ r: cell.row, c: cell.col });
                if (worksheet[cellRef]) {
                    worksheet[cellRef].f = cell.formula;
                }
            }
        });

        XLSX.utils.book_append_sheet(workbook, worksheet, "PMS Data");

        // Generate buffer
        const buffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' });

        return {
            buffer,
            filename: `${template.template_name}.xlsx`,
            contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        };
    },

    // Upload and process filled Excel file
    async uploadFilledData(req: Request) {
        if (!req.file) {
            throw new ApiError(400, "No file uploaded");
        }

        const { templateId } = req.body;
        const uploadedBy = req.body.userId;

        if (!templateId) {
            throw new ApiError(400, "Template ID is required");
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
            // Read uploaded Excel file
            const workbook = XLSX.readFile(filePath);
            const sheetName = workbook.SheetNames[0];
            const sheet = workbook.Sheets[sheetName];

            // Convert to JSON with headers
            const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

            if (jsonData.length < 2) {
                throw new ApiError(400, "Excel file is empty or has no data rows");
            }

            // First row is headers
            const headers = jsonData[0] as string[];

            // Find employee-related column indices
            const employeeIdIndex = headers.findIndex(h =>
                String(h).toLowerCase().includes('employee') &&
                (String(h).toLowerCase().includes('id') || String(h).toLowerCase().includes('no'))
            );

            if (employeeIdIndex === -1) {
                throw new ApiError(400, "Could not find Employee ID column in the uploaded file");
            }

            // Process each data row
            for (let i = 1; i < jsonData.length; i++) {
                const row = jsonData[i] as any[];

                if (!row || row.length === 0) continue;

                const employeeIdValue = row[employeeIdIndex];

                if (!employeeIdValue) {
                    errors.push({
                        rowNumber: i + 1,
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
                        rowNumber: i + 1,
                        employeeId: employeeIdValue,
                        error: `Employee with ID ${employeeIdValue} not found`
                    });
                    continue;
                }

                // Create row data object
                const rowData: Record<string, any> = {};
                headers.forEach((header, idx) => {
                    if (row[idx] !== undefined && row[idx] !== null) {
                        rowData[String(header)] = row[idx];
                    }
                });

                // Save to database
                const employeeData = await EmployeeDataModel.create({
                    templateId,
                    employeeId: employee._id,
                    uploadedBy,
                    uploadBatchId,
                    data: rowData,
                    rowNumber: i + 1,
                    status: 'validated',
                    validationErrors: []
                });

                results.push({
                    rowNumber: i + 1,
                    employeeId: employeeIdValue,
                    employeeName: employee.name,
                    dataId: employeeData._id
                });
            }

            // Cleanup uploaded file
            fs.unlinkSync(filePath);

            return {
                uploadBatchId,
                templateName: template.template_name,
                totalRows: jsonData.length - 1,
                successCount: results.length,
                errorCount: errors.length,
                results,
                errors
            };

        } catch (error) {
            // Cleanup on error
            if (fs.existsSync(filePath)) {
                fs.unlinkSync(filePath);
            }
            throw error;
        }
    },

    // Get upload history for a user
    async getUploadHistory(userId: string) {
        const history = await EmployeeDataModel.aggregate([
            { $match: { uploadedBy: userId } },
            {
                $group: {
                    _id: "$uploadBatchId",
                    templateId: { $first: "$templateId" },
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
            .populate('employeeId', 'employeeId name email')
            .populate('templateId', 'template_name')
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
    }
};

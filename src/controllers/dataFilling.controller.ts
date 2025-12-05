import { DataFillingService } from "../services/dataFilling.service";
import { controllerHandler } from "../utils/controllerHandler";

export const getTemplatesForSelection = controllerHandler(
    async (req) => {
        const result = await DataFillingService.getTemplatesForSelection();
        return result;
    },
    {
        statusCode: 200,
        message: "Templates retrieved successfully",
    }
);

export const getTemplatePreview = controllerHandler(
    async (req) => {
        const { templateId } = req.params;
        const result = await DataFillingService.getTemplatePreview(templateId);
        return result;
    },
    {
        statusCode: 200,
        message: "Template preview retrieved successfully",
    }
);

// Note: Download controller is slightly different as it returns a file
export const downloadTemplate = async (req: any, res: any, next: any) => {
    try {
        const { templateId } = req.params;
        const result = await DataFillingService.downloadTemplate(templateId);

        res.setHeader('Content-Type', result.contentType);
        res.setHeader('Content-Disposition', `attachment; filename=${result.filename}`);
        res.send(result.buffer);
    } catch (error) {
        next(error);
    }
};

export const validateUploadedFile = controllerHandler(
    async (req) => {
        if (!req.file) {
            throw new Error("No file uploaded");
        }
        const { templateId } = req.body;
        if (!templateId) {
            throw new Error("Template ID is required");
        }
        const result = await DataFillingService.validateUploadedFile(req.file.path, templateId);
        return result;
    },
    {
        statusCode: 200,
        message: "File validated successfully",
    }
);

export const uploadFilledData = controllerHandler(
    async (req) => {
        const result = await DataFillingService.uploadFilledData(req);
        return result;
    },
    {
        statusCode: 201,
        message: "Data uploaded and processed successfully",
    }
);

export const getUploadHistory = controllerHandler(
    async (req) => {
        // Get userId from authenticated user (set by verifyToken middleware)
        const userId = (req as any).user?.id;
        const result = await DataFillingService.getUploadHistory(userId);
        return result;
    },
    {
        statusCode: 200,
        message: "Upload history retrieved successfully",
    }
);

export const getDataByEmployee = controllerHandler(
    async (req) => {
        const { employeeId } = req.params;
        const result = await DataFillingService.getDataByEmployee(employeeId);
        return result;
    },
    {
        statusCode: 200,
        message: "Employee data retrieved successfully",
    }
);

export const getDataByBatch = controllerHandler(
    async (req) => {
        const { batchId } = req.params;
        const result = await DataFillingService.getDataByBatch(batchId);
        return result;
    },
    {
        statusCode: 200,
        message: "Batch data retrieved successfully",
    }
);

export const getDataSummary = controllerHandler(
    async (req) => {
        const result = await DataFillingService.getDataSummary();
        return result;
    },
    {
        statusCode: 200,
        message: "Data summary retrieved successfully",
    }
);

export const getUploadedDataWithCalculations = controllerHandler(
    async (req) => {
        const { batchId } = req.params;
        const result = await DataFillingService.getUploadedDataWithCalculations(batchId);
        return result;
    },
    {
        statusCode: 200,
        message: "Calculated data retrieved successfully",
    }
);

export const exportUploadedDataAsExcel = async (req: any, res: any, next: any) => {
    try {
        const { batchId } = req.params;
        const result = await DataFillingService.exportUploadedDataAsExcel(batchId);

        res.setHeader('Content-Type', result.contentType);
        res.setHeader('Content-Disposition', `attachment; filename=${result.filename}`);
        res.send(result.buffer);
    } catch (error) {
        next(error);
    }
};


import { Router } from "express";
import multer from "multer";
import {
    getTemplatesForSelection,
    getTemplatePreview,
    downloadTemplate,
    validateUploadedFile,
    uploadFilledData,
    getUploadHistory,
    getDataByEmployee,
    getDataByBatch,
    getDataSummary,
    getUploadedDataWithCalculations,
    exportUploadedDataAsExcel
} from "../controllers/dataFilling.controller";
import { verifyToken } from "../middlewares/auth.middleware";

const router = Router();
const upload = multer({ dest: "uploads/" });

// Apply auth middleware to all routes
router.use(verifyToken);

// Template routes
router.get("/templates", getTemplatesForSelection);
router.get("/template/:templateId/preview", getTemplatePreview);
router.get("/download/:templateId", downloadTemplate);

// File validation route (before actual upload)
router.post("/validate", upload.single("file"), validateUploadedFile);

// Upload filled data
router.post("/upload", upload.single("file"), uploadFilledData);

// History and data retrieval routes
router.get("/history", getUploadHistory);
router.get("/summary", getDataSummary);
router.get("/employee/:employeeId", getDataByEmployee);
router.get("/batch/:batchId", getDataByBatch);
router.get("/batch/:batchId/calculated", getUploadedDataWithCalculations);
router.get("/batch/:batchId/export", exportUploadedDataAsExcel);

export default router;

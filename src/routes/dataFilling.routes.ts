import { Router } from "express";
import multer from "multer";
import {
    getTemplatesForSelection,
    getTemplatePreview,
    downloadTemplate,
    uploadFilledData,
    getUploadHistory,
    getDataByEmployee,
    getDataSummary
} from "../controllers/dataFilling.controller";
import { verifyToken } from "../middlewares/auth.middleware";

const router = Router();
const upload = multer({ dest: "uploads/" });

// Apply auth middleware to all routes
router.use(verifyToken);

router.get("/templates", getTemplatesForSelection);
router.get("/template/:templateId/preview", getTemplatePreview);
router.get("/download/:templateId", downloadTemplate);
router.post("/upload", upload.single("file"), uploadFilledData);
router.get("/history", getUploadHistory);
router.get("/summary", getDataSummary);
router.get("/employee/:employeeId", getDataByEmployee);

export default router;

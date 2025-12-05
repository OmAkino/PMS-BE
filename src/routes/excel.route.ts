import { Router } from "express";
import multer from "multer";
import {
  getHeaderTemplate,
  uploadHeaderExcel,
  getAllTemplates,
  getTemplateById,
  getTemplateByName,
  downloadTemplateFile,
  deleteTemplate
} from "../controllers/excel.controller";
import { verifyToken } from "../middlewares/auth.middleware";

const router = Router();
const upload = multer({ dest: "uploads/" });

// Apply auth middleware to protected routes
router.post("/uploadHeaderExcel", verifyToken, upload.single("file"), uploadHeaderExcel);
router.get("/template", getHeaderTemplate); // GET /excel/template?templateName=name
router.get("/templates", getAllTemplates); // GET /excel/templates - for dropdown
router.get("/template/id/:templateId", verifyToken, getTemplateById); // GET /excel/template/id/:id
router.get("/template/:templateName", getTemplateByName); // GET /excel/template/PMS-APAC-Header
router.get("/download/:templateId", verifyToken, downloadTemplateFile); // Download template file
router.delete("/template/:templateId", verifyToken, deleteTemplate); // Delete template

export default router;
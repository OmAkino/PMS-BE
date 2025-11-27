import { Router } from "express";
import multer from "multer";
import { 
  getHeaderTemplate, 
  uploadHeaderExcel, 
  getAllTemplates,
  getTemplateByName 
} from "../controllers/excel.controller";

const router = Router();
const upload = multer({ dest: "uploads/" });

router.post("/uploadHeaderExcel", upload.single("file"), uploadHeaderExcel);
router.get("/template", getHeaderTemplate); // GET /excel/template?templateName=name
router.get("/templates", getAllTemplates); // GET /excel/templates - for dropdown
router.get("/template/:templateName", getTemplateByName); // GET /excel/template/PMS-APAC-Header

export default router;
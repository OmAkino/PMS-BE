import { ExcelService } from "../services/excel.service";
import { controllerHandler } from "../utils/controllerHandler";

export const uploadHeaderExcel = controllerHandler(
  async (req) => {
    const result = await ExcelService.uploadHeaderExcel(req);
    return result;
  },
  {
    statusCode: 201,
    message: "Header Excel uploaded successfully",
  }
);

export const getHeaderTemplate = controllerHandler(
  async (req) => {
    const { templateName } = req.query;
    console.log("Query templateName:", templateName);

    // If no templateName provided, use default
    const name = templateName as string || "PMS-APAC-Header";
    const result = await ExcelService.getHeaderTemplate(name);
    return result;
  },
  {
    statusCode: 200,
    message: "Template retrieved successfully",
  }
);

export const getAllTemplates = controllerHandler(
  async (req) => {
    const result = await ExcelService.getAllTemplates();
    return result;
  },
  {
    statusCode: 200,
    message: "Templates list retrieved successfully",
  }
);

export const getTemplateById = controllerHandler(
  async (req) => {
    const { templateId } = req.params;
    const result = await ExcelService.getTemplateById(templateId);
    return result;
  },
  {
    statusCode: 200,
    message: "Template retrieved successfully",
  }
);

export const getTemplateByName = controllerHandler(
  async (req) => {
    const { templateName } = req.params;
    console.log("Param templateName:", templateName);

    const result = await ExcelService.getTemplateByName(templateName);
    return result;
  },
  {
    statusCode: 200,
    message: "Template retrieved successfully",
  }
);

// Download template file
export const downloadTemplateFile = async (req: any, res: any, next: any) => {
  try {
    const { templateId } = req.params;
    const result = await ExcelService.downloadTemplateFile(templateId);

    res.setHeader('Content-Type', result.contentType);
    res.setHeader('Content-Disposition', `attachment; filename=${result.filename}`);
    res.send(result.buffer);
  } catch (error) {
    next(error);
  }
};

export const deleteTemplate = controllerHandler(
  async (req) => {
    const { templateId } = req.params;
    const result = await ExcelService.deleteTemplate(templateId);
    return result;
  },
  {
    statusCode: 200,
    message: "Template deleted successfully",
  }
);
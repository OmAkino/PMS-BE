import { Request } from "express";
import XLSX from "xlsx";
import fs from "fs";
import PMSHeaderTemplate, { ICellDefinition } from "../models/PMSHeader.model";

export const ExcelService = {
  uploadHeaderExcel: async (req: Request) => {
    if (!req.file) {
      throw new Error("No file uploaded");
    }

    const filePath = req.file.path;

    try {
      const workbook = XLSX.readFile(filePath, { cellFormula: true, cellText: false });
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      const range = XLSX.utils.decode_range(sheet["!ref"]!);

      const sheetStructure: ICellDefinition[] = [];
      const formulaRows = new Set<number>();
      let dataInputRanges: string[] = [];
      let protectedRanges: string[] = [];

      // Parse all cells in the sheet
      for (let r = range.s.r; r <= range.e.r; r++) {
        for (let c = range.s.c; c <= range.e.c; c++) {
          const cellAddress = XLSX.utils.encode_cell({ r, c });
          const cell = sheet[cellAddress];
          
          if (!cell) continue;

          // Determine cell type
          let cellType: 'header' | 'formula' | 'data' | 'metadata' = 'data';
          let isLocked = false;
          let dataType: 'string' | 'number' | 'date' | 'formula' | 'percentage' = 'string';

          // Detect formula cells
          if (cell.f) {
            cellType = 'formula';
            dataType = 'formula';
            formulaRows.add(r);
            isLocked = true; // Formula cells are typically locked
          }

          // Detect header rows (usually first few rows)
          if (r <= 2) {
            cellType = 'header';
            isLocked = true;
          }

          // Detect metadata cells (like "Name:", "Designation:", etc.)
          if (typeof cell.v === 'string' && (cell.v.includes(':') || cell.v.includes('Part:'))) {
            cellType = 'metadata';
            isLocked = true;
          }

          // Detect data input cells (empty or numeric cells in data areas)
          if (r >= 7 && r <= 19 && [1, 4, 7, 10, 13, 16, 19].includes(c)) { // Based on your Excel structure
            cellType = 'data';
            isLocked = false;
            if (typeof cell.v === 'number') {
              dataType = 'number';
            }
          }

          // Detect percentage cells
          if (typeof cell.v === 'number' && cell.v <= 1) {
            dataType = 'percentage';
          }

          const cellDefinition: ICellDefinition = {
            row: r,
            col: c,
            address: cellAddress,
            value: cell.v,
            formula: cell.f || null,
            type: cellType,
            is_locked: isLocked,
            data_type: dataType
          } as ICellDefinition;

          sheetStructure.push(cellDefinition);

          // Collect data input ranges for UI
          if (cellType === 'data' && !isLocked) {
            dataInputRanges.push(cellAddress);
          }

          if (isLocked) {
            protectedRanges.push(cellAddress);
          }
        }
      }

      // Create or update template - use dynamic template name from request or default
      // Generate unique template name
      const baseName = req.body.templateName || "PMS-Header";
      const timestamp = new Date().getTime();
      const randomSuffix = Math.random().toString(36).substring(2, 8);
      const templateName = `${baseName}-${timestamp}-${randomSuffix}`;

      const templateData = {
        template_name: templateName,
        description: req.body.description || "Performance Management System Header Template",
        version: "1.0.0",
        sheet_structure: sheetStructure,
        metadata: {
          total_rows: range.e.r + 1,
          total_columns: range.e.c + 1,
          formula_rows: Array.from(formulaRows),
          data_input_ranges: dataInputRanges,
          protected_ranges: protectedRanges
        },
        is_active: true
      };

      // Create new template (remove upsert to always create new)
      const template = await PMSHeaderTemplate.create(templateData);
      fs.unlinkSync(filePath);
      return {
        template_id: template._id,
        template_name: template.template_name,
        total_cells: sheetStructure.length,
        data_input_cells: dataInputRanges.length,
        formula_cells: formulaRows.size
      };


    } catch (error) {
      fs.unlinkSync(filePath);
      throw error;
    }
  },

  // Get template for UI rendering
  getHeaderTemplate: async (templateName: string = "PMS-APAC-Header") => {
    const template = await PMSHeaderTemplate.findOne({ 
      template_name: templateName, 
      is_active: true 
    });

    if (!template) {
      throw new Error(`Template "${templateName}" not found`);
    }

    return template;
  },

  // Get all templates (for dropdown list)
  getAllTemplates: async () => {
    const templates = await PMSHeaderTemplate.find(
      { is_active: true },
      { 
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

  // Get template by name (specific endpoint)
  getTemplateByName: async (templateName: string) => {
    const template = await PMSHeaderTemplate.findOne({ 
      template_name: templateName, 
      is_active: true 
    });

    if (!template) {
      throw new Error(`Template "${templateName}" not found`);
    }

    return template;
  }
};
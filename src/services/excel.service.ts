import { Request } from "express";
import XLSX from "xlsx";
import fs from "fs";
import PMSHeaderTemplate, { ICellDefinition, IColumnMapping, IFormulaDefinition } from "../models/PMSHeader.model";

// Helper function to extract cell references from a formula
const extractCellReferences = (formula: string): string[] => {
  // Match cell references like A1, B2, AB12, etc.
  const cellRefRegex = /\$?[A-Z]+\$?\d+/gi;
  const matches = formula.match(cellRefRegex);
  return matches ? [...new Set(matches.map(m => m.replace(/\$/g, '')))] : [];
};

// Helper function to convert column index to letter
const columnIndexToLetter = (index: number): string => {
  let letter = '';
  let temp = index;
  while (temp >= 0) {
    letter = String.fromCharCode((temp % 26) + 65) + letter;
    temp = Math.floor(temp / 26) - 1;
  }
  return letter;
};

// Helper function to detect employee-related column mappings
const detectEmployeeFieldMapping = (headerName: string): string | null => {
  const lowerHeader = headerName.toLowerCase().trim();

  if (lowerHeader.includes('employee') && (lowerHeader.includes('id') || lowerHeader.includes('no') || lowerHeader.includes('number'))) {
    return 'employeeId';
  }
  if (lowerHeader === 'name' || lowerHeader === 'employee name' || lowerHeader === 'full name') {
    return 'name';
  }
  if (lowerHeader.includes('email') || lowerHeader.includes('e-mail')) {
    return 'email';
  }
  if (lowerHeader.includes('designation') || lowerHeader.includes('title') || lowerHeader.includes('position')) {
    return 'designation';
  }
  if (lowerHeader.includes('department') || lowerHeader.includes('dept')) {
    return 'department';
  }
  if (lowerHeader.includes('division')) {
    return 'division';
  }
  if (lowerHeader.includes('geography') || lowerHeader.includes('location') || lowerHeader.includes('region')) {
    return 'geography';
  }

  return null;
};

export const ExcelService = {
  uploadHeaderExcel: async (req: Request) => {
    if (!req.file) {
      throw new Error("No file uploaded");
    }

    const filePath = req.file.path;

    try {
      // Read file buffer for storage
      const fileBuffer = fs.readFileSync(filePath);

      const workbook = XLSX.readFile(filePath, { cellFormula: true, cellText: false });
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      const range = XLSX.utils.decode_range(sheet["!ref"]!);

      const sheetStructure: ICellDefinition[] = [];
      const columnMappings: IColumnMapping[] = [];
      const formulaDefinitions: IFormulaDefinition[] = [];
      const formulaRows = new Set<number>();
      const formulaCells: string[] = [];
      let dataInputRanges: string[] = [];
      let protectedRanges: string[] = [];

      // Detect header row and data start row
      let headerRowIndex = -1;
      let dataStartRow = -1;
      const headerRowCandidates: { row: number; cellCount: number }[] = [];

      // First pass: detect header row (row with most non-empty text cells)
      for (let r = range.s.r; r <= Math.min(range.e.r, 10); r++) {
        let textCellCount = 0;
        for (let c = range.s.c; c <= range.e.c; c++) {
          const cellAddress = XLSX.utils.encode_cell({ r, c });
          const cell = sheet[cellAddress];
          if (cell && typeof cell.v === 'string' && cell.v.trim().length > 0 && !cell.f) {
            textCellCount++;
          }
        }
        if (textCellCount > 0) {
          headerRowCandidates.push({ row: r, cellCount: textCellCount });
        }
      }

      // Find the row with most headers
      if (headerRowCandidates.length > 0) {
        headerRowCandidates.sort((a, b) => b.cellCount - a.cellCount);
        headerRowIndex = headerRowCandidates[0].row;
        dataStartRow = headerRowIndex + 1;
      }

      // Extract headers and create column mappings
      const headers: { [key: number]: string } = {};
      const employeeFieldMapping: {
        employeeIdColumn: number;
        nameColumn?: number;
        emailColumn?: number;
        designationColumn?: number;
        departmentColumn?: number;
        divisionColumn?: number;
      } = {
        employeeIdColumn: -1
      };

      if (headerRowIndex >= 0) {
        for (let c = range.s.c; c <= range.e.c; c++) {
          const cellAddress = XLSX.utils.encode_cell({ r: headerRowIndex, c });
          const cell = sheet[cellAddress];
          if (cell && cell.v !== undefined && cell.v !== null) {
            const headerName = String(cell.v).trim();
            headers[c] = headerName;

            // Detect employee field mapping
            const mappedField = detectEmployeeFieldMapping(headerName);

            // Update employee field mapping
            if (mappedField === 'employeeId') {
              employeeFieldMapping.employeeIdColumn = c;
            } else if (mappedField === 'name') {
              employeeFieldMapping.nameColumn = c;
            } else if (mappedField === 'email') {
              employeeFieldMapping.emailColumn = c;
            } else if (mappedField === 'designation') {
              employeeFieldMapping.designationColumn = c;
            } else if (mappedField === 'department') {
              employeeFieldMapping.departmentColumn = c;
            } else if (mappedField === 'division') {
              employeeFieldMapping.divisionColumn = c;
            }

            // Create column mapping
            const columnMapping: IColumnMapping = {
              columnIndex: c,
              columnName: columnIndexToLetter(c),
              headerName: headerName,
              mappedField: mappedField,
              dataType: 'string',
              isRequired: mappedField === 'employeeId'
            } as IColumnMapping;

            columnMappings.push(columnMapping);
          }
        }
      }

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
          let formulaDependencies: string[] = [];

          // Detect formula cells
          if (cell.f) {
            cellType = 'formula';
            dataType = 'formula';
            formulaRows.add(r);
            formulaCells.push(cellAddress);
            isLocked = true;

            // Extract cell references from formula
            formulaDependencies = extractCellReferences(cell.f);

            // Create formula definition
            const formulaDefinition: IFormulaDefinition = {
              cellAddress,
              row: r,
              col: c,
              formula: cell.f,
              dependentCells: formulaDependencies,
              resultType: typeof cell.v === 'number' ? 'number' : 'string'
            } as IFormulaDefinition;

            formulaDefinitions.push(formulaDefinition);
          }

          // Detect header rows
          if (r === headerRowIndex) {
            cellType = 'header';
            isLocked = true;
          }

          // Detect metadata cells (like "Name:", "Designation:", etc.)
          if (typeof cell.v === 'string' && (cell.v.includes(':') || cell.v.includes('Part:'))) {
            cellType = 'metadata';
            isLocked = true;
          }

          // Detect data type
          if (cell.t === 'n') {
            dataType = cell.f ? 'formula' : 'number';
            // Check if it's a percentage (value between 0 and 1, or has % format)
            if (typeof cell.v === 'number' && cell.v >= 0 && cell.v <= 1) {
              if (cell.z && cell.z.includes('%')) {
                dataType = 'percentage';
              }
            }
          } else if (cell.t === 'd') {
            dataType = 'date';
          }

          // Determine if this is a data input cell
          if (r >= dataStartRow && !cell.f && !isLocked) {
            cellType = 'data';
            isLocked = false;
            dataInputRanges.push(cellAddress);
          }

          const cellDefinition: ICellDefinition = {
            row: r,
            col: c,
            address: cellAddress,
            value: cell.v,
            formula: cell.f || null,
            formulaDependencies: formulaDependencies,
            type: cellType,
            is_locked: isLocked,
            data_type: dataType,
            headerRowIndex: headerRowIndex,
            columnHeader: headers[c] || null
          } as ICellDefinition;

          sheetStructure.push(cellDefinition);

          if (isLocked || cell.f) {
            protectedRanges.push(cellAddress);
          }
        }
      }

      // Generate unique template name
      const baseName = req.body.templateName || "PMS-Template";
      const timestamp = new Date().getTime();
      const randomSuffix = Math.random().toString(36).substring(2, 8);
      const templateName = `${baseName}-${timestamp}-${randomSuffix}`;

      const templateData = {
        template_name: templateName,
        description: req.body.description || "Performance Management System Template",
        version: "1.0.0",
        sheet_structure: sheetStructure,
        column_mappings: columnMappings,
        formula_definitions: formulaDefinitions,
        employee_field_mapping: employeeFieldMapping,
        header_row_index: headerRowIndex,
        data_start_row: dataStartRow,
        metadata: {
          total_rows: range.e.r + 1,
          total_columns: range.e.c + 1,
          formula_rows: Array.from(formulaRows),
          data_input_ranges: dataInputRanges,
          protected_ranges: protectedRanges,
          formula_cells: formulaCells
        },
        file_buffer: fileBuffer,
        file_name: req.file?.originalname || `${templateName}.xlsx`,
        is_active: true
      };

      // Create new template
      const template = await PMSHeaderTemplate.create(templateData);
      fs.unlinkSync(filePath);

      return {
        template_id: template._id,
        template_name: template.template_name,
        total_cells: sheetStructure.length,
        data_input_cells: dataInputRanges.length,
        formula_cells: formulaCells.length,
        header_row_index: headerRowIndex,
        data_start_row: dataStartRow,
        column_mappings: columnMappings.map(cm => ({
          column: cm.columnName,
          header: cm.headerName,
          mappedField: cm.mappedField
        })),
        employee_field_mapping: employeeFieldMapping,
        formula_definitions: formulaDefinitions.map(fd => ({
          cell: fd.cellAddress,
          formula: fd.formula,
          dependencies: fd.dependentCells
        }))
      };

    } catch (error) {
      if (fs.existsSync(filePath)) {
        fs.unlinkSync(filePath);
      }
      throw error;
    }
  },

  // Get template for UI rendering
  getHeaderTemplate: async (templateName: string = "PMS-APAC-Header") => {
    // Try to find the specific template first
    let template = await PMSHeaderTemplate.findOne({
      template_name: templateName,
      is_active: true
    });

    // If not found, try to get the most recent active template
    if (!template) {
      console.log(`Template "${templateName}" not found, fetching most recent template...`);
      template = await PMSHeaderTemplate.findOne({
        is_active: true
      }).sort({ createdAt: -1 });

      if (!template) {
        throw new Error(`No active templates found. Please upload a template first.`);
      }

      console.log(`Using most recent template: ${template.template_name}`);
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

  // Get template by ID
  getTemplateById: async (templateId: string) => {
    const template = await PMSHeaderTemplate.findById(templateId);

    if (!template) {
      throw new Error(`Template not found`);
    }

    return template;
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
  },

  // Download original template file
  downloadTemplateFile: async (templateId: string) => {
    const template = await PMSHeaderTemplate.findById(templateId);

    if (!template) {
      throw new Error("Template not found");
    }

    if (template.file_buffer) {
      return {
        buffer: template.file_buffer,
        filename: template.file_name || `${template.template_name}.xlsx`,
        contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      };
    }

    // Fallback: regenerate from sheet structure
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
    const buffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' });

    return {
      buffer,
      filename: `${template.template_name}.xlsx`,
      contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    };
  },

  // Delete template
  deleteTemplate: async (templateId: string) => {
    const result = await PMSHeaderTemplate.findByIdAndUpdate(
      templateId,
      { is_active: false },
      { new: true }
    );

    if (!result) {
      throw new Error("Template not found");
    }

    return { message: "Template deleted successfully" };
  }
};
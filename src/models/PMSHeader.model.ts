// models/PMSHeader.model.ts
import mongoose, { Schema, Document } from "mongoose";

// Column mapping for employee-related fields
export interface IColumnMapping {
  columnIndex: number;
  columnName: string;
  headerName: string;
  mappedField: string | null; // Maps to employee fields like 'employeeId', 'name', 'email', etc.
  dataType: 'string' | 'number' | 'date' | 'formula' | 'percentage';
  isRequired: boolean;
  validationRules?: {
    min?: number;
    max?: number;
    pattern?: string;
  };
}

// Formula definition with cell dependencies
export interface IFormulaDefinition {
  cellAddress: string;
  row: number;
  col: number;
  formula: string;
  dependentCells: string[]; // Cells this formula depends on
  resultType: 'number' | 'percentage' | 'string';
}

export interface ICellDefinition extends Document {
  row: number;
  col: number;
  address: string;
  value: string | number | null;
  formula: string | null;
  formulaDependencies?: string[]; // Cells this formula depends on
  type: 'header' | 'formula' | 'data' | 'metadata';
  is_locked: boolean;
  data_type: 'string' | 'number' | 'date' | 'formula' | 'percentage';
  headerRowIndex?: number; // Which row this column header is on
  columnHeader?: string; // The header text for this column
  validation_rules?: {
    required?: boolean;
    min?: number;
    max?: number;
    pattern?: string;
  };
}

export interface IPMSHeaderTemplate extends Document {
  template_name: string;
  description?: string;
  version: string;
  sheet_structure: ICellDefinition[];
  column_mappings: IColumnMapping[];
  formula_definitions: IFormulaDefinition[];
  employee_field_mapping: {
    employeeIdColumn: number;
    nameColumn?: number;
    emailColumn?: number;
    designationColumn?: number;
    departmentColumn?: number;
    divisionColumn?: number;
  };
  header_row_index: number; // Which row contains the headers
  data_start_row: number; // First row of actual data
  metadata: {
    total_rows: number;
    total_columns: number;
    formula_rows: number[];
    data_input_ranges: string[];
    protected_ranges: string[];
    formula_cells: string[];
  };
  is_active: boolean;
  file_buffer?: Buffer; // Store original file for download
  file_name?: string;
  createdAt: Date;
  updatedAt: Date;
}

const ColumnMappingSchema = new Schema({
  columnIndex: { type: Number, required: true },
  columnName: { type: String, required: true }, // e.g., "A", "B", "C"
  headerName: { type: String, required: true }, // e.g., "Employee ID", "Name"
  mappedField: { type: String }, // e.g., "employeeId", "name", "email"
  dataType: {
    type: String,
    enum: ['string', 'number', 'date', 'formula', 'percentage'],
    default: 'string'
  },
  isRequired: { type: Boolean, default: false },
  validationRules: {
    min: { type: Number },
    max: { type: Number },
    pattern: { type: String }
  }
});

const FormulaDefinitionSchema = new Schema({
  cellAddress: { type: String, required: true },
  row: { type: Number, required: true },
  col: { type: Number, required: true },
  formula: { type: String, required: true },
  dependentCells: [{ type: String }],
  resultType: {
    type: String,
    enum: ['number', 'percentage', 'string'],
    default: 'number'
  }
});

const CellDefinitionSchema = new Schema({
  row: { type: Number, required: true },
  col: { type: Number, required: true },
  address: { type: String, required: true }, // e.g., "A1", "B2"
  value: { type: Schema.Types.Mixed }, // Original value from Excel
  formula: { type: String }, // Excel formula
  formulaDependencies: [{ type: String }], // Cells this formula depends on
  type: {
    type: String,
    enum: ['header', 'formula', 'data', 'metadata'],
    required: true
  },
  is_locked: { type: Boolean, default: false },
  headerRowIndex: { type: Number },
  columnHeader: { type: String },
  data_type: {
    type: String,
    enum: ['string', 'number', 'date', 'formula', 'percentage'],
    default: 'string'
  },
  validation_rules: {
    required: { type: Boolean, default: false },
    min: { type: Number },
    max: { type: Number },
    pattern: { type: String }
  }
});

const PMSHeaderTemplateSchema = new Schema(
  {
    template_name: { type: String, required: true, unique: true },
    description: { type: String },
    version: { type: String, default: "1.0.0" },
    sheet_structure: [CellDefinitionSchema],
    column_mappings: [ColumnMappingSchema],
    formula_definitions: [FormulaDefinitionSchema],
    employee_field_mapping: {
      employeeIdColumn: { type: Number, default: -1 },
      nameColumn: { type: Number },
      emailColumn: { type: Number },
      designationColumn: { type: Number },
      departmentColumn: { type: Number },
      divisionColumn: { type: Number }
    },
    header_row_index: { type: Number, default: 0 },
    data_start_row: { type: Number, default: 1 },
    metadata: {
      total_rows: { type: Number, required: true },
      total_columns: { type: Number, required: true },
      formula_rows: [Number], // Rows that contain formulas
      data_input_ranges: [String], // e.g., ["B8:B19", "C8:C19"]
      protected_ranges: [String], // Cells that shouldn't be edited
      formula_cells: [String] // All cells containing formulas
    },
    file_buffer: { type: Buffer }, // Store original file for download
    file_name: { type: String },
    is_active: { type: Boolean, default: true }
  },
  { timestamps: true }
);

// Index for faster queries
PMSHeaderTemplateSchema.index({ template_name: 1 });
PMSHeaderTemplateSchema.index({ is_active: 1 });

export default mongoose.model<IPMSHeaderTemplate>("PMSHeaderTemplate", PMSHeaderTemplateSchema);
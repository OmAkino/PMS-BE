// models/PMSHeader.model.ts
import mongoose, { Schema, Document } from "mongoose";

export interface ICellDefinition extends Document {
  row: number;
  col: number;
  address: string;
  value: string | number | null;
  formula: string | null;
  type: 'header' | 'formula' | 'data' | 'metadata';
  is_locked: boolean;
  data_type: 'string' | 'number' | 'date' | 'formula' | 'percentage';
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
  metadata: {
    total_rows: number;
    total_columns: number;
    formula_rows: number[];
    data_input_ranges: string[];
    protected_ranges: string[];
  };
  is_active: boolean;
}

const CellDefinitionSchema = new Schema({
  row: { type: Number, required: true },
  col: { type: Number, required: true },
  address: { type: String, required: true }, // e.g., "A1", "B2"
  value: { type: Schema.Types.Mixed }, // Original value from Excel
  formula: { type: String }, // Excel formula
  type: { 
    type: String, 
    enum: ['header', 'formula', 'data', 'metadata'],
    required: true 
  },
  is_locked: { type: Boolean, default: false },
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
    metadata: {
      total_rows: { type: Number, required: true },
      total_columns: { type: Number, required: true },
      formula_rows: [Number], // Rows that contain formulas
      data_input_ranges: [String], // e.g., ["B8:B19", "C8:C19"]
      protected_ranges: [String] // Cells that shouldn't be edited
    },
    is_active: { type: Boolean, default: true }
  },
  { timestamps: true }
);

export default mongoose.model<IPMSHeaderTemplate>("PMSHeaderTemplate", PMSHeaderTemplateSchema);
import mongoose, { Schema, Document } from "mongoose";

export interface IEmployeeData extends Document {
    templateId: mongoose.Types.ObjectId;
    templateName?: string;
    employeeId: mongoose.Types.ObjectId;
    employeeCode?: string;
    uploadedBy?: mongoose.Types.ObjectId;
    uploadBatchId: string;
    data: Record<string, any>;
    calculatedData?: Record<string, any>; // Stores formula calculation results
    rawData?: Record<string, any>; // Stores raw cell values before calculation
    rowNumber: number;
    status: 'pending' | 'validated' | 'error' | 'processed';
    validationErrors: string[];
    uploadedAt?: Date;
    createdAt: Date;
    updatedAt: Date;
}

const employeeDataSchema = new Schema<IEmployeeData>(
    {
        templateId: {
            type: Schema.Types.ObjectId,
            ref: "PMSHeaderTemplate",
            required: true
        },
        templateName: {
            type: String
        },
        employeeId: {
            type: Schema.Types.ObjectId,
            ref: "Employee",
            required: true
        },
        employeeCode: {
            type: String,
            index: true
        },
        uploadedBy: {
            type: Schema.Types.ObjectId,
            ref: "User"
        },
        uploadBatchId: {
            type: String,
            required: true,
            index: true
        },
        data: {
            type: Schema.Types.Mixed,
            required: true
        },
        calculatedData: {
            type: Schema.Types.Mixed,
            default: {}
        },
        rawData: {
            type: Schema.Types.Mixed,
            default: {}
        },
        rowNumber: {
            type: Number,
            required: true
        },
        status: {
            type: String,
            enum: ['pending', 'validated', 'error', 'processed'],
            default: 'processed'
        },
        validationErrors: [{
            type: String
        }],
        uploadedAt: {
            type: Date,
            default: Date.now
        }
    },
    { timestamps: true }
);

// Compound indexes for efficient queries
employeeDataSchema.index({ templateId: 1, employeeId: 1 });
employeeDataSchema.index({ uploadedBy: 1, createdAt: -1 });
employeeDataSchema.index({ status: 1 });
employeeDataSchema.index({ uploadBatchId: 1, rowNumber: 1 });
employeeDataSchema.index({ employeeCode: 1, uploadBatchId: 1 });

export const EmployeeDataModel = mongoose.model<IEmployeeData>("EmployeeData", employeeDataSchema);

export default EmployeeDataModel;

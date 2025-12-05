import mongoose, { Schema, Document } from "mongoose";

export interface IEmployeeData extends Document {
    templateId: mongoose.Types.ObjectId;
    employeeId: mongoose.Types.ObjectId;
    uploadedBy: mongoose.Types.ObjectId;
    uploadBatchId: string;
    data: Record<string, any>;
    rowNumber: number;
    status: 'pending' | 'validated' | 'error';
    validationErrors: string[];
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
        employeeId: {
            type: Schema.Types.ObjectId,
            ref: "Employee",
            required: true
        },
        uploadedBy: {
            type: Schema.Types.ObjectId,
            ref: "User",
            required: true
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
        rowNumber: {
            type: Number,
            required: true
        },
        status: {
            type: String,
            enum: ['pending', 'validated', 'error'],
            default: 'pending'
        },
        validationErrors: [{
            type: String
        }]
    },
    { timestamps: true }
);

// Compound indexes for efficient queries
employeeDataSchema.index({ templateId: 1, employeeId: 1 });
employeeDataSchema.index({ uploadedBy: 1, createdAt: -1 });
employeeDataSchema.index({ status: 1 });

export const EmployeeDataModel = mongoose.model<IEmployeeData>("EmployeeData", employeeDataSchema);

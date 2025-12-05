import mongoose, { Schema, Document } from "mongoose";

export interface IEmployee extends Document {
    employeeId: string;
    name: string;
    email: string;
    designation: string;
    division: string;
    geography: string;
    department: string;
    reportsTo: mongoose.Types.ObjectId | null;
    isActive: boolean;
    createdBy: mongoose.Types.ObjectId;
    createdAt: Date;
    updatedAt: Date;
}

const employeeSchema = new Schema<IEmployee>(
    {
        employeeId: {
            type: String,
            required: true,
            unique: true,
            trim: true
        },
        name: {
            type: String,
            required: true,
            trim: true
        },
        email: {
            type: String,
            required: true,
            unique: true,
            trim: true,
            lowercase: true
        },
        designation: {
            type: String,
            required: true,
            trim: true
        },
        division: {
            type: String,
            required: true,
            trim: true
        },
        geography: {
            type: String,
            required: true,
            trim: true
        },
        department: {
            type: String,
            required: true,
            trim: true
        },
        reportsTo: {
            type: Schema.Types.ObjectId,
            ref: "Employee",
            default: null
        },
        isActive: {
            type: Boolean,
            default: true
        },
        createdBy: {
            type: Schema.Types.ObjectId,
            ref: "User",
            required: true
        }
    },
    { timestamps: true }
);

// Index for faster queries
employeeSchema.index({ reportsTo: 1 });
employeeSchema.index({ isActive: 1 });
employeeSchema.index({ department: 1 });

export const EmployeeModel = mongoose.model<IEmployee>("Employee", employeeSchema);

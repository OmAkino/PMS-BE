import { Request } from "express";
import { EmployeeModel, IEmployee } from "../models/Employee.model";
import { UserModel } from "../models/user.model";
import { ApiError } from "../utils/ApiError";

export const EmployeeService = {
    // Get default manager (SuperAdmin user)
    async getDefaultManager() {
        const superAdmin = await UserModel.findOne({ role: "superAdmin" });
        if (!superAdmin) {
            throw new ApiError(404, "No SuperAdmin found in the system");
        }
        return superAdmin;
    },

    // Create new employee
    async createEmployee(req: Request) {
        const {
            employeeId,
            name,
            email,
            designation,
            division,
            geography,
            department,
            reportsTo
        } = req.body;

        // Validate required fields
        if (!employeeId || !name || !email || !designation || !division || !geography || !department) {
            throw new ApiError(400, "All required fields must be provided");
        }

        // Check for duplicate employeeId
        const existingById = await EmployeeModel.findOne({ employeeId });
        if (existingById) {
            throw new ApiError(400, `Employee with ID ${employeeId} already exists`);
        }

        // Check for duplicate email
        const existingByEmail = await EmployeeModel.findOne({ email: email.toLowerCase() });
        if (existingByEmail) {
            throw new ApiError(400, `Employee with email ${email} already exists`);
        }

        // Get createdBy from authenticated user (assuming userId comes from auth middleware)
        const createdBy = req.body.userId;
        if (!createdBy) {
            throw new ApiError(401, "User authentication required");
        }

        // If no reportsTo provided, set to null (will show as "SuperAdmin" in UI)
        const employeeData = {
            employeeId,
            name,
            email: email.toLowerCase(),
            designation,
            division,
            geography,
            department,
            reportsTo: reportsTo || null,
            createdBy
        };

        const employee = await EmployeeModel.create(employeeData);

        // Populate reportsTo for response
        await employee.populate('reportsTo', 'name employeeId');

        return employee;
    },

    // Get all active employees
    async getAllEmployees() {
        const employees = await EmployeeModel.find({ isActive: true })
            .populate('reportsTo', 'name employeeId email')
            .populate('createdBy', 'firstName lastName email')
            .sort({ createdAt: -1 });

        return employees;
    },

    // Get employee by ID
    async getEmployeeById(id: string) {
        const employee = await EmployeeModel.findById(id)
            .populate('reportsTo', 'name employeeId email designation')
            .populate('createdBy', 'firstName lastName email');

        if (!employee) {
            throw new ApiError(404, "Employee not found");
        }

        return employee;
    },

    // Update employee
    async updateEmployee(req: Request) {
        const { id } = req.params;
        const updateData = req.body;

        // Remove fields that shouldn't be updated directly
        delete updateData.createdBy;
        delete updateData.createdAt;
        delete updateData.userId;

        // Check if employee exists
        const existing = await EmployeeModel.findById(id);
        if (!existing) {
            throw new ApiError(404, "Employee not found");
        }

        // Check for duplicate employeeId if being updated
        if (updateData.employeeId && updateData.employeeId !== existing.employeeId) {
            const duplicate = await EmployeeModel.findOne({
                employeeId: updateData.employeeId,
                _id: { $ne: id }
            });
            if (duplicate) {
                throw new ApiError(400, `Employee ID ${updateData.employeeId} already exists`);
            }
        }

        // Check for duplicate email if being updated
        if (updateData.email && updateData.email.toLowerCase() !== existing.email) {
            const duplicate = await EmployeeModel.findOne({
                email: updateData.email.toLowerCase(),
                _id: { $ne: id }
            });
            if (duplicate) {
                throw new ApiError(400, `Email ${updateData.email} already exists`);
            }
            updateData.email = updateData.email.toLowerCase();
        }

        const updatedEmployee = await EmployeeModel.findByIdAndUpdate(
            id,
            updateData,
            { new: true }
        )
            .populate('reportsTo', 'name employeeId email')
            .populate('createdBy', 'firstName lastName email');

        return updatedEmployee;
    },

    // Soft delete employee
    async deleteEmployee(id: string) {
        const employee = await EmployeeModel.findById(id);
        if (!employee) {
            throw new ApiError(404, "Employee not found");
        }

        // Check if employee has direct reports
        const directReports = await EmployeeModel.countDocuments({
            reportsTo: id,
            isActive: true
        });

        if (directReports > 0) {
            throw new ApiError(400, `Cannot delete employee. ${directReports} employee(s) report to this person. Please reassign them first.`);
        }

        employee.isActive = false;
        await employee.save();

        return { message: "Employee deleted successfully" };
    },

    // Get employees by manager
    async getEmployeesByManager(managerId: string) {
        const employees = await EmployeeModel.find({
            reportsTo: managerId,
            isActive: true
        })
            .populate('reportsTo', 'name employeeId')
            .sort({ name: 1 });

        return employees;
    },

    // Get all employees for dropdown (minimal data)
    async getEmployeesForDropdown() {
        const employees = await EmployeeModel.find(
            { isActive: true },
            { _id: 1, employeeId: 1, name: 1, designation: 1 }
        ).sort({ name: 1 });

        return employees;
    },

    // Search employees
    async searchEmployees(query: string) {
        const searchRegex = new RegExp(query, 'i');

        const employees = await EmployeeModel.find({
            isActive: true,
            $or: [
                { name: searchRegex },
                { employeeId: searchRegex },
                { email: searchRegex },
                { department: searchRegex },
                { designation: searchRegex }
            ]
        })
            .populate('reportsTo', 'name employeeId')
            .limit(50);

        return employees;
    }
};

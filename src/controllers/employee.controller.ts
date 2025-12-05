import { EmployeeService } from "../services/employee.service";
import { controllerHandler } from "../utils/controllerHandler";

export const createEmployee = controllerHandler(
    async (req) => {
        const result = await EmployeeService.createEmployee(req);
        return result;
    },
    {
        statusCode: 201,
        message: "Employee created successfully",
    }
);

export const getAllEmployees = controllerHandler(
    async (req) => {
        const result = await EmployeeService.getAllEmployees();
        return result;
    },
    {
        statusCode: 200,
        message: "Employees retrieved successfully",
    }
);

export const getEmployeeById = controllerHandler(
    async (req) => {
        const { id } = req.params;
        const result = await EmployeeService.getEmployeeById(id);
        return result;
    },
    {
        statusCode: 200,
        message: "Employee retrieved successfully",
    }
);

export const updateEmployee = controllerHandler(
    async (req) => {
        const result = await EmployeeService.updateEmployee(req);
        return result;
    },
    {
        statusCode: 200,
        message: "Employee updated successfully",
    }
);

export const deleteEmployee = controllerHandler(
    async (req) => {
        const { id } = req.params;
        const result = await EmployeeService.deleteEmployee(id);
        return result;
    },
    {
        statusCode: 200,
        message: "Employee deleted successfully",
    }
);

export const getEmployeesByManager = controllerHandler(
    async (req) => {
        const { managerId } = req.params;
        const result = await EmployeeService.getEmployeesByManager(managerId);
        return result;
    },
    {
        statusCode: 200,
        message: "Direct reports retrieved successfully",
    }
);

export const getEmployeesForDropdown = controllerHandler(
    async (req) => {
        const result = await EmployeeService.getEmployeesForDropdown();
        return result;
    },
    {
        statusCode: 200,
        message: "Employees list retrieved successfully",
    }
);

export const getDefaultManager = controllerHandler(
    async (req) => {
        const result = await EmployeeService.getDefaultManager();
        return result;
    },
    {
        statusCode: 200,
        message: "Default manager retrieved successfully",
    }
);

import { Router } from "express";
import {
    createEmployee,
    getAllEmployees,
    getEmployeeById,
    updateEmployee,
    deleteEmployee,
    getEmployeesByManager,
    getEmployeesForDropdown,
    getDefaultManager
} from "../controllers/employee.controller";
import { verifyToken } from "../middlewares/auth.middleware";

const router = Router();

// Apply auth middleware to all routes
router.use(verifyToken);

router.post("/", createEmployee);
router.get("/", getAllEmployees);
router.get("/dropdown", getEmployeesForDropdown);
router.get("/default-manager", getDefaultManager);
router.get("/:id", getEmployeeById);
router.put("/:id", updateEmployee);
router.delete("/:id", deleteEmployee);
router.get("/by-manager/:managerId", getEmployeesByManager);

export default router;

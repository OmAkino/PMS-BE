import { Router } from "express";
import userRoutes from "./user.routes";
import adminRoutes from "./admin.routes";
import authRoutes from "./auth.routes";
import excelRoute from "./excel.route";
import employeeRoutes from "./employee.routes";
import dataFillingRoutes from "./dataFilling.routes";

const router = Router();

router.use("/auth", authRoutes);
router.use("/users", userRoutes);
router.use("/admins", adminRoutes);
router.use("/excel", excelRoute);
router.use("/employees", employeeRoutes);
router.use("/data-filling", dataFillingRoutes);

export default router;

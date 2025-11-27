import { Router } from "express";
import userRoutes from "./user.routes";
import adminRoutes from "./admin.routes";
import authRoutes from "./auth.routes";
import excelRoute from "./excel.route";

const router = Router();

router.use("/auth", authRoutes);
router.use("/users", userRoutes);
router.use("/admins", adminRoutes);
router.use("/excel",excelRoute);

export default router;

import { Router } from "express";
import { getUser, createUser } from "../controllers/user.controller";

const router = Router();

router.get("/getUser", getUser);
router.post("/create", createUser);

export default router;

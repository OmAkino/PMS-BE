import { Router } from "express";
import { login, register, refresh, logout } from "../controllers/auth.controller";
import { validate } from "../validation/validate";
import { loginSchema, logoutSchema, refreshSchema, registerSchema } from "../validation/auth.validate";

const router = Router();

router.post("/register",validate(registerSchema) ,register);
router.post("/login",validate(loginSchema) ,login);
router.post("/refresh",validate(refreshSchema) ,refresh);
router.post("/logout",validate(logoutSchema) ,logout);

export default router;

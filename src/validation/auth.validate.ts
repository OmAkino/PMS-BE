// src/validations/auth.validation.ts
import { z } from "zod";

export const registerSchema = z.object({
  firstName: z.string().trim().min(2, "Name must be at least 2 characters long"),
  lastName: z.string().trim().min(2, "Name must be at least 2 characters long"),
  email: z.string().trim().email("Invalid email format"),
  password: z.string().trim().min(6, "Password must be at least 6 characters long"),
});

export const loginSchema = z.object({
  email: z.string().trim().email("Invalid email format"),
  password: z.string().trim().min(6, "Password must be at least 6 characters long"),
});

export const refreshSchema = z.object({
  refreshToken: z.string().trim().nonempty("Refresh token is required"),
});

export const logoutSchema = z.object({
  refreshToken: z.string().trim().nonempty("Refresh token is required"),
});

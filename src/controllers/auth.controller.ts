import { controllerHandler } from "../utils/controllerHandler";
import { AuthService } from "../services/auth.service";

export const register = controllerHandler(
  async (req) => AuthService.register(req),
  { statusCode: 201, message: "User registered successfully" }
);

export const login = controllerHandler(
  async (req) => AuthService.login(req),
  { statusCode: 200, message: "Login successful" }
);

export const refresh = controllerHandler(
  async (req) => AuthService.refreshToken(req),
  { statusCode: 200, message: "Token refreshed successfully" }
);

export const logout = controllerHandler(
  async (req) => AuthService.logout(req),
  { statusCode: 200, message: "Logged out successfully" }
);

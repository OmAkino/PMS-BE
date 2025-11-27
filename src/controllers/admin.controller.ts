import { RequestHandler } from "express";
import { controllerHandler } from "../utils/controllerHandler";
import { AuthService } from "../services/auth.service";


export const getAdmin = controllerHandler(
  async (req) => {
    // Call the AuthService register function
    const user = await AuthService.register(req);

    // If you need to send additional things like emails or audit logs,
    // you can call them here, for example:
    // await EmailService.sendVerificationEmail(user.email);
    // await AuditService.logEvent("USER_REGISTERED", user.id);

    return { user };
  },
  {
    statusCode: 201,
    message: "User registered successfully ",
  }
);



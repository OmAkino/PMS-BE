import { Request, Response, NextFunction } from "express";
import { sendResponse } from "./sendResponse";
import { ApiError } from "./ApiError";
import { catchAsync } from "./catchAsync";

type ControllerFn<T = any> = (req: Request) => Promise<T>;

interface HandlerOptions {
  statusCode?: number;
  message?: string;
}

/**
 * controllerHandler - wraps an async controller to handle errors and responses consistently
 * @param fn - async controller/service function that returns data
 * @param options - optional { statusCode, message }
 */
export const controllerHandler = <T>(
  fn: ControllerFn<T>,
  options?: HandlerOptions
) =>
  catchAsync(async (req: Request, res: Response, next: NextFunction) => {
    try {
      const data = await fn(req);

      // Send standardized success response
      sendResponse(res, data, options?.message || "Request successful", options?.statusCode || 200);
    } catch (error: any) {
      // Handle known (operational) errors
      if (error instanceof ApiError) {
        return res.status(error?.statusCode).json({
          success: false,
          message: error?.message || "Request unsuccessful",
        });
      }

      // Pass unhandled errors to global error middleware
      next(error);
    }
  });

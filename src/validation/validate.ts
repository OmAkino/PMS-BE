import { Request, Response, NextFunction } from "express";
import { ZodSchema, ZodError } from "zod";

/**
 * validate - middleware for validating request body, query, or params
 * @param schema - a Zod schema to validate against
 * @param source - optional: "body" | "query" | "params" (default: "body")
 */
export const validate =
  (schema: ZodSchema<any>, source: "body" | "query" | "params" = "body") =>
  (req: Request, res: Response, next: NextFunction) => {
    const result = schema.safeParse(req[source]);

    if (!result.success) {
      const zodError = result.error as ZodError;
      const errors = zodError.issues.map((issue) => ({
        path: issue.path.join("."),
        message: issue.message,
      }));

      return res.status(400).json({
        status: "error",
        message: "Validation failed ",
        errors,
      });
    }

    // Replace req[source] with parsed data
    req[source] = result.data;
    next();
  };

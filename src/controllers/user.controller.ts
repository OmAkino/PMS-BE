import { RequestHandler } from "express";
import { controllerHandler } from "../utils/controllerHandler";

export const getUser = controllerHandler(
  async (req) => {
    
    return {  };
  },
  {
    statusCode: 201,
    message: "User registered successfully ",
  }
);

export const createUser = controllerHandler(
  async (req) => {
    
    return {  };
  },
  {
    statusCode: 201,
    message: "User registered successfully ",
  }
);

import jwt from "jsonwebtoken";
import config from "../config/db/index";

const ACCESS_SECRET = config.ACCESS_SECRET!;
const REFRESH_SECRET = config.REFRESH_TOKEN_SECRET!;
const ACCESS_EXPIRES = "7d";
const REFRESH_EXPIRES = "7d";

export const generateAccessToken = (payload: object) =>
  jwt.sign(payload, ACCESS_SECRET, { expiresIn: ACCESS_EXPIRES });

export const generateRefreshToken = (payload: object) =>
  jwt.sign(payload, REFRESH_SECRET, { expiresIn: REFRESH_EXPIRES });

export const verifyAccessToken = (token: string) =>
  jwt.verify(token, ACCESS_SECRET);

export const verifyRefreshToken = (token: string) =>
  jwt.verify(token, REFRESH_SECRET);

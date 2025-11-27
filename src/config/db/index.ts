// src/config/index.ts
import dotenv from 'dotenv';

dotenv.config();

export default {
  port: process.env.PORT || 5000,
  database_url: process.env.DATABASE_URL,
  node_env: process.env.NODE_ENV || 'development',
  ACCESS_SECRET: process.env.ACCESS_SECRET || "default_access_secret",
  REFRESH_TOKEN_SECRET: process.env.REFRESH_TOKEN_SECRET || "default_refresh_secret",
  GEMINI_API_KEY: process.env.GEMINI_API_KEY || "api_key_missing",
};

import mongoose from "mongoose";
import app from "./app";
import config from "./config/db";
import logger from "./utils/logger";
import initialUserCreation from "./utils/initialUserCreation";

/**
 * When running on Vercel (serverless), we simply export the app.
 * When running locally, we connect to MongoDB and start the server normally.
 */

// If NOT running on Vercel (local / other platform)
if (!process.env.VERCEL) {
  const startServer = async () => {
    try {
      await mongoose.connect(config.database_url as string);
      logger.info("✅ Connected to MongoDB");

      app.listen(config.port, () => {
        logger.info(`🚀 Server running at http://localhost:${config.port}`);
      });

      await initialUserCreation();
    } catch (err) {
      logger.error("❌ Failed to connect to MongoDB: " + err);
    }
  };

  startServer();
}

// ✅ Always export app (required for Vercel)
export default app;

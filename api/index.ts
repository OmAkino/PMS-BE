import app from "../src/app";

// Vercel's Node runtime will use the default export as the request handler.
// Exporting the Express `app` here makes the whole Express app act as the
// serverless handler. Keep this file minimal so Vercel can find the entry.
export default app;

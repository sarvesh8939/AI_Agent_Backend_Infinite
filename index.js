import "dotenv/config";
import express from "express";
import agentRoutes from "./routes/agentRoutes.js";

const app = express();

app.use(express.json());
app.use("/api/agent", agentRoutes);

app.listen(3000, () => {
    console.log("Server started on port 3000");
});
import express from "express";
import path from "path";
import * as console from "console";

const app = express();
const port = process.env.PORT || 3000;

// Serve static files from the 'src' directory
app.use(express.static(path.join(__dirname, "../src")));

// Serve the task pane HTML file
app.get("/taskpane", (req, res) => {
  res.sendFile(path.join(__dirname, "../src/taskpane/taskpane.html"));
});

// Start the server
app.listen(port, () => {
  console.log(`Server is running on port ${port}`);
});

// Error handling middleware
app.use((err, req, res) => {
  console.error(err.stack);
  res.status(500).send("Something broke!");
});

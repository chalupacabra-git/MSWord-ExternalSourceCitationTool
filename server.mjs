import express from "express";
import path from "path";

const app = express();
const PORT = process.env.PORT || 3000;

// Serve static files from the 'src' directory
app.use(express.static(path.join(import.meta.url, "../src")));

// Serve the task pane HTML file
app.get("/taskpane", (req, res) => {
  res.sendFile(path.join(import.meta.url, "../src/taskpane/taskpane.html"));
});

// Start the server
app.listen(PORT, () => {
  console.log(`Server is running on port ${PORT}`);
});

// Error handling middleware
app.use((err, req, res, next) => {
  console.error(err.stack);
  res.status(500).send("Something broke!");
});

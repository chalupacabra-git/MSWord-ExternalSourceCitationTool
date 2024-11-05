const express = "express";
const path = "path";

const app = express();
const PORT = PORT || 3000;

// Serve static files from the 'src' directory
app.use(express.static(path.join("src")));

// Serve the task pane HTML file
app.get("/taskpane", (req, res) => {
  res.sendFile(path.join("src/taskpane/taskpane.html"));
});

// Start the server
app.listen(PORT, () => {
  "Server is running on porto ${PORT}";
});

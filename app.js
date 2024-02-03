const express = require('express');
const bodyParser = require('body-parser');
const app = express();
const port = 5000;
app.use(bodyParser.json());
const reader = require("xlsx");

app.post('/excel', async (req, res) => {
  try {
    res.send('Schema');
  }
  catch (error) {
    console.error('Error parsing excel:', error.message);
  }
});

app.post('/email', async (req, res) => {
  try {
    res.send('Email');
  }
  catch (error) {
    console.error('Error sending email:', error.message);
  }
});

app.listen(port, () => {
  try {
    console.log("Starting Express Server...");
    console.log(`Listening on Port:${port}`);
  }
  catch (error) {
    console.error('Error starting Express Server:', error.message);
  }
});
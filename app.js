const express = require('express');
const bodyParser = require('body-parser');
const multer = require('multer');
const cors = require('cors');

const { getHeadersBySheet } = require('./helpers');

const port = 5000;
const app = express();
app.use(bodyParser.json());
app.use(cors());

const storage = multer.memoryStorage();
const upload = multer({ storage: storage });

app.post('/excel', upload.single('file'), async (req, res) => {
  try {
    const headersBySheet = getHeadersBySheet(req.file);
    res.json(headersBySheet);
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
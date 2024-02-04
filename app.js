const express = require('express');
const bodyParser = require('body-parser');
const multer = require('multer');
const cors = require('cors');
const reader = require('xlsx');

const { getHeadersBySheet, constructMessageFromHtml, sendEmail } = require('./helpers');

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

app.post('/email', upload.single('file'), async (req, res) => {
  try {
    const excel = await reader.read(req.file.buffer);
    const recipientEmail = req.body.recipientEmail;
    const message = req.body.message;
    const selectedSheet = req.body.selectedSheet;
    const json = reader.utils.sheet_to_json(excel.Sheets[selectedSheet]);

    for (let i = 0; i < json.length; i++) {
      const updatedMessage = constructMessageFromHtml(message, json[i]);
      sendEmail(json[i][recipientEmail], 'B# Piano Competition', updatedMessage);
    }
    res.json(updatedMessage);
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
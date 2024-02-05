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
    if (!req.file) {
      return res.status(400).json({ error: 'No file uploaded.' });
    }

    const headersBySheet = getHeadersBySheet(req.file);
    if (!headersBySheet) {
      return res.status(400).json({ error: 'Invalid file format or no headers found.' });
    }

    res.json(headersBySheet);
  } catch (error) {
    console.error('Error processing file:', error);
    res.status(500).json({ error: 'Internal Server Error. Please try again later.' });
  }
});

app.post('/email', upload.single('file'), async (req, res) => {
  try {
    const { selectedSheet, recipientEmailExcelColumn, subject, messageHtml, type, testRecipientEmail } = req.body;
    
    if (!req.file) {
      return res.status(400).json({ error: 'No file uploaded.' });
    }
    if (!selectedSheet || !recipientEmailExcelColumn || !subject || !messageHtml || !type) {
      return res.status(400).json({ error: 'Bad Request' });
    }

    const workbook = reader.read(req.file.buffer);
    const sheet = workbook.Sheets[selectedSheet];
    const sheetJson = reader.utils.sheet_to_json(sheet);

    if (type === "preview") {
      const updatedMessageHtml = constructMessageFromHtml(messageHtml, sheetJson[0]);
      res.json({
        selectedSheet,
        recipientEmailExcelColumn,
        subject,
        updatedMessageHtml
      });
    } else if (type === "test") {
      const updatedMessageHtml = constructMessageFromHtml(messageHtml, sheetJson[0]);
      sendEmail(testRecipientEmail, subject, updatedMessageHtml);
      res.json("Test Email Successfully Sent!");
    } else {
      for (let i = 0; i < sheetJson.length; i++) {
        const updatedMessage = constructMessageFromHtml(messageHtml, sheetJson[i]);
        sendEmail(sheetJson[i][recipientEmailExcelColumn], subject, updatedMessage);
      }
      res.json("Emails Successfully Sent!");
    }
  }
  catch (error) {
    console.error('Error sending emails:', error);
    return res.status(500).json({ error: 'Internal Server Error. Please try again later.' });
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
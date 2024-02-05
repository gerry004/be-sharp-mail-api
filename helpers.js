const nodemailer = require("nodemailer");
const reader = require("xlsx");

const transporter = nodemailer.createTransport({
  host: 'smtp.titan.email',
  port: 465,
  secure: true,
  auth: {
    user: 'team@besharppiano.ie',
    pass: 'Oobleck!23'
  }
});

const delay = async (ms = 1000) => new Promise(resolve => setTimeout(resolve, ms))

const sendEmail = async (recipient, subject, message) => {
  const mailOptions = {
    from: '"B# Piano Competition" team@besharppiano.ie',
    to: recipient,
    subject: subject,
    html: message
  };

  transporter.sendMail(mailOptions, (error, info) => {
    if (error) {
      console.log(error);
    } else {
      console.log('Email Successfully Sent to ' + recipient);
    }
  });

  await delay(1000)
}

const constructMessageFromHtml = (message, excelSheetData) => {
  updatedMessage = replacePlaceholders(message, excelSheetData);
  return updatedMessage;
}

function replacePlaceholders(html, data) {
  return html.replace(/{{([^}]+)}}/g, function (match, key) {
    return data.hasOwnProperty(key) ? data[key] :'';
  });
}

const getHeadersBySheet = (file) => {
  const workbook = reader.read(file.buffer);
  const sheetNames = workbook.SheetNames;
  const headersBySheet = {};
  sheetNames.forEach(sheetName => {
    headersBySheet[sheetName] = getSheetHeaders(workbook.Sheets[sheetName]);
  });
  return headersBySheet;
};

const getSheetHeaders = (sheet) => {
  const headers = [];
  const range = sheet && sheet['!ref'] ? reader.utils.decode_range(sheet['!ref']) : null;

  if (range) {
    for (let C = range.s.c; C <= range.e.c; ++C) {
      const cellAddress = { c: C, r: range.s.r };
      const cellRef = reader.utils.encode_cell(cellAddress);
      const header = sheet[cellRef]?.v;
      if (header !== undefined) headers.push(header);
    }
  } else {
    console.error('Sheet range is undefined.');
  }

  return headers;
};

module.exports = { sendEmail, getHeadersBySheet, constructMessageFromHtml };

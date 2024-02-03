const nodemailer = require("nodemailer");

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

const sendEmail = async (sender, recipient, subject, message) => {
  const mailOptions = {
    from: sender, 
    to: recipient, 
    subject: subject,
    text: message
  };

  transporter.sendMail(mailOptions, (error, info) => {
    if (error) {
      console.log(error);
    } else {
      console.log('Email Successfully Sent to ' + addressee);
    }
  });

  await delay(1000)
}

module.exports = { sendEmail };

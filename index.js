const express = require("express");
const nodemailer = require("nodemailer");
const exceljs = require("exceljs");
require("dotenv").config();

const app = express();
const port = 3000;

// Middleware to parse JSON bodies
app.use(express.json());

// Function to read emails from Excel file and filter based on a specific domain
async function readEmailsFromExcel(filename, domain) {
  const workbook = new exceljs.Workbook();
  await workbook.xlsx.readFile(filename);
  const worksheet = workbook.getWorksheet(1);
  const emails = [];

  worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    if (rowNumber > 1) {
      const email = row.getCell(1).value;
      if (typeof email === 'string') {
        const sanitizedEmail = email.replace(/\s/g, ''); // Remove all whitespace characters
        if (sanitizedEmail.endsWith(domain)) {
          emails.push(sanitizedEmail);
        }
      }
    }
  });

  return emails;
}


// Send the email when the server starts
(async () => {
  try {
    const domainToFilter = "@gmail.com"; // Specify the domain you want to filter for

    // Read emails from Excel file and filter based on the specified domain
    console.log(`Reading emails from Excel file and filtering based on domain "${domainToFilter}"...`);
    const emails = await readEmailsFromExcel("emails.xlsx", domainToFilter);
    console.log("Emails read from Excel:", emails);

    // Create a transporter object
    console.log("Creating transporter object...");
    const transporter = nodemailer.createTransport({
      service: "gmail",
      auth: {
        type: "OAuth2",
        user: process.env.MAIL_USERNAME,
        pass: process.env.MAIL_PASSWORD,
        clientId: process.env.OAUTH_CLIENTID,
        clientSecret: process.env.OAUTH_CLIENT_SECRET,
        refreshToken: process.env.OAUTH_REFRESH_TOKEN,
      },
    });

    // Loop through each email and send
    for (const email of emails) {
      // Create mail options
      const mailOptions = {
        from: process.env.MAIL_USERNAME,
        to: email,
        subject: "Enhance Your School's Online Presence With a Professional Website.",
        text: `Hello Principal,
        
I hope this email finds you well. I'm a web developer with five years of experience, and I'm reaching out because I believe your school would benefit greatly from having a professional website.
        
I'd love to discuss how we can work together to create a tailored platform that reflects your school's values and achievements. If you're interested, please reply to this email, and we can schedule a brief call.
        
Looking forward to the possibility of collaborating with you.
        
Best regards,
        
Mahfujul
Founder at Folix`,
      };

      // Send email
      await transporter.sendMail(mailOptions);
      console.log(`Email sent successfully to ${email}`);
    }

    console.log("All emails sent successfully");
  } catch (error) {
    console.error("Error sending emails:", error);
  }
})();

// Start the server
app.listen(port, () => {
  console.log(`Server is listening at http://localhost:${port}`);
});

import * as nodemailer from 'nodemailer';
import * as fs from 'fs';
import * as xlsx from 'xlsx';
import * as dotenv from 'dotenv';

dotenv.config(); // Load environment variables from .env file

// Interface for the data structure in the Excel file
interface EmailData {
  Name: string;
  email: string;
}

// Function to read data from Excel file
function readExcel(filePath: string): EmailData[] {
  const workbook = xlsx.readFile(filePath);
  const sheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];
  return xlsx.utils.sheet_to_json(worksheet) as EmailData[];
}

// Function to send emails using Nodemailer
async function sendEmails(emailsData: EmailData[]): Promise<void> {
  const transporter = nodemailer.createTransport({
    service: 'your-email-service', // e.g., 'gmail'
    auth: {
      user: process.env.EMAIL,
      pass: process.env.PASS,
    },
  });

  // Loop through each entry in the Excel file and send emails
  for (const data of emailsData) {
    const { Name, email } = data;

    // Email content
    const mailOptions: nodemailer.SendMailOptions = {
      from: process.env.EMAIL,
      to: email,
      subject: 'Subject of the Email',
      text: `Hello ${Name},\n\nThis is the body of your email.`,
    };

    // Send email
    try {
      await transporter.sendMail(mailOptions);
      console.log(`Email sent to ${Name} (${email})`);
    } catch (error) {
      console.error(`Error sending email to ${Name} (${email}): ${error.message}`);
    }
  }
}

// Main function
function main(): void {
  const excelFilePath = './file.xlsx';

  // Check if the Excel file exists
  if (!fs.existsSync(excelFilePath)) {
    console.error(`Error: Excel file not found at ${excelFilePath}`);
    return;
  }

  // Read data from Excel file
  const emailsData = readExcel(excelFilePath);

  // Check if there is any data to process
  if (emailsData.length === 0) {
    console.error('Error: No data found in the Excel file');
    return;
  }

  // Send emails
  sendEmails(emailsData);
}

// Run the main function
main();
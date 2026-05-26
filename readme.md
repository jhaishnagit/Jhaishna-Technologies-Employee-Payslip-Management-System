# Jhaishna Technologies Employee Payslip Management System

A secure, efficient web application to upload Excel payroll data, generate professional payslips, preview them, download individual or bulk PDFs, and send payslips via email.
## Features

- Drag & drop or file upload of .xlsx payroll files
- Automatic parsing of employee data and month/year from Excel
- Beautiful, printable payslip preview (A4 format)
- Generate & download single payslip as PDF
- Bulk download all payslips as ZIP archive
- Send single or bulk payslips via email (Gmail supported)
- Computer-generated payslip watermark (no signature required)
- Responsive design + print/PDF optimized layout
- Environment variable support for secure email credentials

## Tech Stack

- **Backend**: Node.js + Express
- **Frontend**: HTML, CSS, Vanilla JavaScript + EJS templating
- **Excel Parsing**: SheetJS (`xlsx`)
- **PDF Generation**: Puppeteer (headless Chrome)
- **Email Sending**: Nodemailer (Gmail with App Password)
- **File Handling**: Multer + Archiver (ZIP)
- **Dependencies Management**: dotenv (environment variables)

## Project Structure


# Install dependences
 npm install dotenv
 npm install express multer xlsx ejs puppeteer nodemailer archiver dotenv
 npx puppeteer browsers install chrome


# Install the nodemon 
# Nodemon is a development tool used in Node.js that automatically restarts the server whenever changes are made to the code. It helps save time and makes development faster.

 npm install -g nodemon
 nodemon -v

# In package.json write this after {},

"scripts": {
  "start": "node app.js",
  "dev": "nodemon app.js"
}

# And Finally Run This Command We Get The Output
nodemon server.js

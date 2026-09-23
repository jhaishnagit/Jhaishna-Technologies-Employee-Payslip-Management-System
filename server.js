const axios = require("axios");
const FormData = require("form-data");

console.log("RUNNING FROM:", __dirname);

require("dotenv").config();
const express = require("express");
const multer = require("multer");
const xlsx = require("xlsx");
const ejs = require("ejs");
const puppeteer = require("puppeteer");
const nodemailer = require("nodemailer");
const fs = require("fs");
const path = require("path");
const archiver = require("archiver");

const app = express();
const upload = multer({ dest: "uploads/" });

app.use(express.json());
app.use(express.static("public"));
app.set("view engine", "ejs");
app.set("views", path.join(__dirname, "views"));

const logoBase64 = fs.readFileSync(
  path.join(__dirname, "public/images/image.png"),
  "base64",
);

let employeesData = [];
let excelMonth = "";

const transporter = nodemailer.createTransport({
  host: "smtp.zoho.com",
  port: 465,
  secure: true, // VERY IMPORTANT for 465
  auth: {
    user: process.env.EMAIL_USER,
    pass: process.env.EMAIL_PASSWORD,
  },
});

function numberToWords(num) {
  num = Math.round(num);
  if (num === 0) return "Zero";
  const a = [
    "",
    "One",
    "Two",
    "Three",
    "Four",
    "Five",
    "Six",
    "Seven",
    "Eight",
    "Nine",
  ];
  const b = [
    "",
    "",
    "Twenty",
    "Thirty",
    "Forty",
    "Fifty",
    "Sixty",
    "Seventy",
    "Eighty",
    "Ninety",
  ];
  const c = [
    "Ten",
    "Eleven",
    "Twelve",
    "Thirteen",
    "Fourteen",
    "Fifteen",
    "Sixteen",
    "Seventeen",
    "Eighteen",
    "Nineteen",
  ];
  let w = "";
  if (num >= 10000000) {
    w += numberToWords(Math.floor(num / 10000000)) + " Crore ";
    num %= 10000000;
  }
  if (num >= 100000) {
    w += numberToWords(Math.floor(num / 100000)) + " Lakh ";
    num %= 100000;
  }
  if (num >= 1000) {
    w += numberToWords(Math.floor(num / 1000)) + " Thousand ";
    num %= 1000;
  }
  if (num >= 100) {
    w += a[Math.floor(num / 100)] + " Hundred ";
    num %= 100;
  }
  if (num >= 10 && num <= 19) w += c[num - 10];
  else {
    if (num >= 20) w += b[Math.floor(num / 10)] + " ";
    if (num % 10) w += a[num % 10];
  }
  return w.trim();
}

const delay = (ms) => new Promise((r) => setTimeout(r, ms));

app.post("/upload", upload.single("excel"), (req, res) => {
  try {
    const wb = xlsx.readFile(req.file.path);
    const sheet = wb.Sheets[wb.SheetNames[0]];
    const rows = xlsx.utils.sheet_to_json(sheet, { header: 1, defval: "" });

    excelMonth = "";

    const monthRegex =
      /(jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|jun(?:e)?|jul(?:y)?|aug(?:ust)?|sep(?:tember)?|oct(?:ober)?|nov(?:ember)?|dec(?:ember)?)[^\d]*(\d{4})/i;

    outer: for (const row of rows) {
      for (const cell of row) {
        if (typeof cell === "string") {
          const match = cell.match(monthRegex);
          if (match) {
            const monthName = match[1];
            const year = match[2];

            const formattedMonth =
              monthName.charAt(0).toUpperCase() +
              monthName.slice(1).toLowerCase();

            excelMonth = `${formattedMonth} ${year}`;
            break outer;
          }
        }
      }
    }

    if (!excelMonth) {
      excelMonth = "January 2026";
    }

    const headerIndex = rows.findIndex((r) =>
      r.some((c) =>
        String(c).trim().toLowerCase().includes("name of the employee"),
      ),
    );

    if (headerIndex === -1) {
      throw new Error("Header row not found");
    }

    const header = rows[headerIndex].map((h) =>
      String(h || "")
        .trim()
        .toLowerCase(),
    );
    console.log("HEADERS:", header);
    const lastCol = header.length - 1;

    const col = {
      name: header.findIndex((h) => h.includes("name of the employee")),

      email: header.findIndex(
        (h) => h.includes("e-mail") || h.includes("email"),
      ),
      pan: header.findIndex((h) => h.includes("pan")),
      empId: header.findIndex(
        (h) => h.includes("employee id") || h.includes("id"),
      ),
      uan: header.findIndex((h) => h.includes("uan")),
      grossExcel: header.findIndex(
        (h) => h.includes("gross salary") || h.includes("gross"),
      ),
      daysWorked: header.findIndex((h) => h.includes("no.of days worked")),
      totalDays: header.findIndex(
        (h) => h.includes("emp. working days") || h.includes("total days"),
      ),
      gender: header.findIndex((h) => h.includes("gender")),
      basicDA: header.findIndex(
        (h) => h.includes("basic + da") || h.includes("basic"),
      ),
      hra: header.findIndex((h) => h.includes("hra")),
      performanceBonus: header.findIndex(
        (h) => h.includes("bonous") || h.includes("bonus"),
      ),
      otherAllowance: header.findIndex((h) => h.includes("allowance")),
      epfEmp: header.findIndex((h) => h.includes("epf") && !h.includes("13%")),

      esicEmp: header.findIndex((h) => h.trim() === "esic"),

      advance: header.findIndex(
        (h) => h.includes("sal.adv") || h.includes("advance"),
      ),

      profTax: header.findIndex(
        (h) => h.includes("p.tax") || h.includes("prof tax"),
      ),
      mediClaim: header.findIndex(
        (h) =>
          h.includes("medi") || h.includes("health") || h.includes("medical"),
      ),
      tds: header.findIndex((h) => h.includes("tds")),
      // 🔥 FIXED CTC POSITIONS FROM RIGHT SIDE
      monthlyCTC: header.findIndex((h) => h.includes("ctc")),
      actualCTC: header.findIndex((h) => h.includes("actual")),
      diffLeaves: header.findIndex((h) => h.includes("diff")),
    };
    // ✅ ADD THIS HERE
    if (col.name === -1) {
      throw new Error("Employee Name column not found in Excel");
    }

    employeesData = rows
      .slice(headerIndex + 2)
      .filter((r) => r[col.name])

      .map((r, i) => {
        const n = (v) => {
          if (v === "-" || v === "" || v === null || v === undefined) return 0;
          return parseFloat(String(v).replace(/,/g, "").trim()) || 0;
        };

        const daysWorked = n(r[col.daysWorked]);
        const totalDays = n(r[col.totalDays]) || 31;

        const basic = Math.round(n(r[col.basicDA]));
        const hra = Math.round(n(r[col.hra]));
        const bonus = Math.round(n(r[col.performanceBonus]));
        const other = Math.round(n(r[col.otherAllowance]));

        let gross = basic + hra + bonus + other;

        // ================= EMPLOYEE CONTRIBUTIONS =================
        const empEpf = Math.round(n(r[col.epfEmp]));
        const empEsic = Math.round(n(r[col.esicEmp]));

        // Find Net Pay column
        const netPayCol = header.findIndex((h) => h.includes("net pay"));

        // Employer columns are next after Net Pay
        const erEpf = netPayCol !== -1 ? Math.round(n(r[netPayCol + 1])) : 0;
        const erEsic = netPayCol !== -1 ? Math.round(n(r[netPayCol + 2])) : 0;

        const subTotal = erEpf + erEsic;

        // ================= OTHER DEDUCTIONS =================
        const ptax = Math.round(n(r[col.profTax]));
        const medi = Math.round(n(r[col.mediClaim]));
        const tds = Math.round(n(r[col.tds]));
        const advanceSalary = Math.round(n(r[col.advance]));
        console.log(
          "Advance column index:",
          col.advance,
          "Value:",
          r[col.advance],
          "Parsed:",
          advanceSalary,
        );

        // Use the same bonus value that is already in earnings
        const bonusDeducted = bonus; // ← this is the key line

        const ded =
          empEpf + empEsic + ptax + bonusDeducted + medi + tds + advanceSalary;
        const net = gross - ded;

        // Original Offer CTC from Excel
        const monthlyCtc = n(r[col.monthlyCTC]);
        const actualCtc = n(r[col.actualCTC]);
        const diffLeaves = n(r[col.diffLeaves]);

        return {
          empId: String(r[col.empId] || `JTPL${100 + i}`).trim(),
          empName: String(r[col.name] || "").trim(),
          email: String(r[col.email] || "").trim(),
          panNo: String(r[col.pan] || "").trim(),
          esicNo: String(
            r[header.findIndex((h) => h === "esic number")] || "New joining",
          ).trim(),

          epfUan: String(r[col.uan] || "").trim(),
          workDays: daysWorked.toFixed(0),
          totalWorkDays: totalDays.toFixed(0),
          gender: String(r[col.gender] || "Male").trim(),

          basic: basic.toFixed(2),
          hra: hra.toFixed(2),
          performanceBonus: bonus.toFixed(2),
          performanceBonusDeducted: bonusDeducted.toFixed(2),
          otherAllowance: other.toFixed(2),
          totalEarnings: gross.toFixed(2),

          empEpf: empEpf.toFixed(2),
          empEsic: empEsic.toFixed(2),
          profTax: ptax.toFixed(2),
          healthInsurance: medi.toFixed(2),
          tds: tds.toFixed(2),
          advanceSalary: advanceSalary.toFixed(2),
          totalDeductions: ded.toFixed(2),

          netPay: net.toFixed(2),
          netPayWords: `(Rupees ${numberToWords(net)} Only)`,

          employerEpf: erEpf.toFixed(2),
          employerEsic: erEsic.toFixed(2),
          subTotal: subTotal.toFixed(2),
          // monthlyCTC: monthlyCtc.toFixed(2),
          // actualCTC: actualCtc.toFixed(2),
          // diffLeaves: diffLeaves.toFixed(2),
          monthlyCTC: Math.round(monthlyCtc),
          actualCTC: Math.round(actualCtc),
          diffLeaves: Math.round(diffLeaves),

          remarks: "",
          laptopDeposit: 0,
          employerLaptop: 0,
        };
      });

    fs.unlinkSync(req.file.path);
    res.json({ success: true, count: employeesData.length, month: excelMonth });
  } catch (err) {
    console.error(err);
    if (req.file?.path) fs.unlinkSync(req.file.path);
    res.status(500).json({ success: false, error: err.message });
  }
});

app.get("/employees", (req, res) => {
  res.json({
    month: excelMonth,
    employees: employeesData,
  });
});

app.post("/download-single", async (req, res) => {
  console.log("Employees length:", employeesData.length);

  try {
    const emp = employeesData[req.body.empIndex];
    if (!emp) return res.status(400).send("Invalid employee");
    const html = await ejs.renderFile(
      path.join(__dirname, "views/payslip.ejs"),
      { ...emp, month: req.body.month || excelMonth, logoBase64 },
    );
    const browser = await puppeteer.launch({
      executablePath: "/usr/bin/chromium-browser",
      headless: "new",
      args: ["--no-sandbox", "--disable-setuid-sandbox"],
    });
    const page = await browser.newPage();
    await page.setContent(html, { waitUntil: "networkidle0" });
    await page.emulateMediaType("print");

    const pdf = await page.pdf({
      format: "A4",
      printBackground: true,
      margin: {
        top: "10mm",
        bottom: "10mm",
        left: "10mm",
        right: "10mm",
      },
    });
    await browser.close();
    res.setHeader("Content-Type", "application/pdf");
    res.send(pdf);
  } catch (err) {
    console.error("DOWNLOAD SINGLE ERROR:", err);
    res.status(500).send("Failed");
  }
});

app.post("/download-all", async (req, res) => {
  const { month } = req.body;
  res.setHeader("Content-Type", "application/zip");
  res.setHeader(
    "Content-Disposition",
    'attachment; filename="all_payslips.zip"',
  );
  const archive = archiver("zip", { zlib: { level: 9 } });
  archive.pipe(res);
  let browser;
  try {
    browser = await puppeteer.launch({
      headless: "new",
      args: ["--no-sandbox", "--disable-setuid-sandbox"],
    });
    for (const emp of employeesData) {
      const html = await ejs.renderFile(
        path.join(__dirname, "views/payslip.ejs"),
        { ...emp, month: month || excelMonth, logoBase64 },
      );
      const page = await browser.newPage();
      await page.setContent(html, {
        waitUntil: "networkidle0",
        timeout: 0,
      });
      await page.emulateMediaType("print");

      const pdf = await page.pdf({
        format: "A4",
        printBackground: true,
        margin: {
          top: "10mm",
          bottom: "10mm",
          left: "10mm",
          right: "10mm",
        },
      });

      await page.close();
      archive.append(Buffer.from(pdf), {
        name: `${emp.empName}_payslip.pdf`,
      });
    }
    await browser.close();
    await archive.finalize();
  } catch (err) {
    console.error("DOWNLOAD ALL ERROR:", err);
    if (browser) await browser.close();
    res.status(500).end();
  }
});

app.post("/send-single", async (req, res) => {
  try {
    const emp = employeesData[req.body.empIndex];
    if (!emp || !emp.email || !emp.email.trim()) {
      return res.status(400).json({ success: false });
    }
    const html = await ejs.renderFile(
      path.join(__dirname, "views/payslip.ejs"),
      { ...emp, month: req.body.month || excelMonth, logoBase64 },
    );
    const browser = await puppeteer.launch({
      // executablePath: '/usr/bin/chromium-browser',
      headless: "new",
      args: ["--no-sandbox", "--disable-setuid-sandbox"],
    });
    const page = await browser.newPage();
    await page.setContent(html, { waitUntil: "load" });
    await page.emulateMediaType("print");

    const pdf = await page.pdf({
      format: "A4",
      printBackground: true,
      margin: {
        top: "10mm",
        bottom: "10mm",
        left: "10mm",
        right: "10mm",
      },
    });

    await browser.close();
    await transporter.sendMail({
      from: `"Jhaishna Technologies Pvt Ltd" <${process.env.EMAIL_USER}>`,
      to: emp.email.trim(),
      subject: `Payslip for ${req.body.month || excelMonth}`,
      text: `Dear ${emp.empName},

Please find attached your payslip for ${req.body.month || excelMonth}.

Regards,
HR Department
Jhaishna Technologies Pvt Ltd`,
      html: `
    <p>Dear ${emp.empName},</p>
    <p>Please find attached your payslip for <b>${req.body.month || excelMonth}</b>.</p>
    <br>
    <p>Regards,<br>HR Department<br>Jhaishna Technologies Pvt Ltd</p>
  `,
      attachments: [
        {
          filename: `${emp.empName}.pdf`,
          content: pdf,
        },
      ],
    });

    const form = new FormData();

    const monthYear = (req.body.month || excelMonth).split(" ");

    form.append("email", emp.email.trim());
    form.append("month", monthYear[0]);
    form.append("year", monthYear[1]);

    form.append("gross_salary", emp.totalEarnings);
    form.append("deductions", emp.totalDeductions);
    form.append("net_salary", emp.netPay);

    form.append("pdf", pdf, {
      filename: `${emp.empName}.pdf`,
      contentType: "application/pdf",
    });

    await axios.post(process.env.HRMS_URL, form, {
      headers: form.getHeaders(),
    });

    console.log(`✅ Single payslip uploaded for ${emp.empName}`);
    res.json({ success: true });
  } catch (err) {
    console.error("SEND SINGLE ERROR:", err);
    res.status(500).json({ success: false });
  }
});

app.post("/send-all", async (req, res) => {
  const { month } = req.body;

  let sent = 0;
  let skipped = [];
  let failed = [];
  let browser;

  try {
    browser = await puppeteer.launch({
      headless: "new",
      args: ["--no-sandbox", "--disable-setuid-sandbox"],
    });

    // ✅ ZIP STREAM FIX (IMPORTANT)
    const { PassThrough } = require("stream");
    const zipStream = new PassThrough();
    const archive = archiver("zip", { zlib: { level: 9 } });

    archive.pipe(zipStream);

    const chunks = [];
    zipStream.on("data", (chunk) => chunks.push(chunk));

    // 🔁 LOOP FOR EMPLOYEES
    for (const emp of employeesData) {
      try {
        const html = await ejs.renderFile(
          path.join(__dirname, "views/payslip.ejs"),
          { ...emp, month: month || excelMonth, logoBase64 },
        );

        const page = await browser.newPage();
        await page.setContent(html, { waitUntil: "load", timeout: 0 });
        await page.emulateMediaType("print");

        const pdf = await page.pdf({
          format: "A4",
          printBackground: true,
        });

        await page.close();

        // ✅ ADD TO ZIP
        archive.append(pdf, {
          name: `${emp.empName}.pdf`,
        });

        // ✅ SEND TO EMPLOYEE
        if (emp.email && emp.email.trim() !== "") {
          await transporter.sendMail({
            from: `"Jhaishna Technologies Pvt Ltd" <${process.env.EMAIL_USER}>`,
            to: emp.email.trim(),
            subject: `Payslip for ${month || excelMonth}`,
            html: `
              <p>Dear ${emp.empName},</p>
              <p>Please find attached your payslip for <b>${month || excelMonth}</b>.</p>
              <br>
              <p>Regards,<br>HR Department</p>
            `,
            attachments: [
              {
                filename: `${emp.empName}.pdf`,
                content: pdf,
              },
            ],
          });

          // ================= Upload PDF to Flask =================

          try {
            const form = new FormData();

            const monthYear = (month || excelMonth).split(" ");

            form.append("email", emp.email.trim());
            form.append("month", monthYear[0]);
            form.append("year", monthYear[1]);
            form.append("gross_salary", emp.totalEarnings);
            form.append("deductions", emp.totalDeductions);
            form.append("net_salary", emp.netPay);

            form.append("pdf", pdf, {
              filename: `${emp.empName}.pdf`,
              contentType: "application/pdf",
            });

            await axios.post(process.env.HRMS_URL, form, {
              headers: form.getHeaders(),
            });

            console.log(`✅ Uploaded payslip for ${emp.empName}`);
          } catch (err) {
            console.log("Upload Error:", err.response?.data || err.message);
          }

          sent++;
        } else {
          skipped.push(emp.empName);
        }
      } catch (err) {
        console.error(`❌ Error for ${emp.empName}:`, err.message);
        failed.push(emp.empName);
      }
    }

    await archive.finalize();

    const zipBuffer = Buffer.concat(chunks);

    console.log(
      "📦 ZIP SIZE:",
      (zipBuffer.length / (1024 * 1024)).toFixed(2),
      "MB",
    );
    console.log("📧 HR EMAIL:", process.env.HR_EMAIL);

    // ✅ SEND ZIP TO HR (WITH ERROR LOG)
    try {
      await transporter.sendMail({
        from: `"Jhaishna Technologies Pvt Ltd" <${process.env.EMAIL_USER}>`,
        to: process.env.HR_EMAIL,
        subject: `All Payslips - ${month || excelMonth}`,
        html: `
          <p>Hello HR,</p>
          <p>All employee payslips are attached.</p>
          <p><b>Month:</b> ${month || excelMonth}</p>
          <p><b>Sent:</b> ${sent}</p>
          <p><b>Skipped:</b> ${skipped.length}</p>
          <p><b>Failed:</b> ${failed.length}</p>
        `,
        attachments: [
          {
            filename: `All_Payslips_${month || excelMonth}.zip`,
            content: zipBuffer,
          },
        ],
      });

      console.log("✅ HR MAIL SENT SUCCESSFULLY");
    } catch (err) {
      console.error("❌ HR MAIL ERROR:", err);
    }

    // ✅ CLOSE BROWSER
    await browser.close();

    res.json({
      success: true,
      sent,
      skipped,
      failed,
      hr: "ZIP sent to HR successfully", // ✅ new message
    });
  } catch (err) {
    if (browser) await browser.close();
    console.error("❌ SEND ALL ERROR:", err);
    res.status(500).json({ success: false });
  }
});
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Server running at http://localhost:${PORT}`);
});

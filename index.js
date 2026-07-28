import express from "express";
import multer from "multer";
import ExcelJS from "exceljs";
import moment from "moment";
import Holidays from "date-holidays";
import { google } from "googleapis";
import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";
import dotenv from "dotenv";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

dotenv.config({ path: path.resolve(__dirname, ".env") });

const IS_FROM_RENDER = process.env.IS_FROM_RENDER === "TRUE";

const DEFAULT_SPREADSHEET_ID = process.env.DEFAULT_SPREADSHEET_ID;
const MASTER_RANGE = process.env.MASTER_RANGE;
const IZIN_RANGE = process.env.IZIN_RANGE;
const DEFAULT_INPUT_FILE = path.resolve(
  __dirname,
  "data/PKM Rasanae Timur.xlsx",
);
const OUTPUT_DIR = path.resolve(__dirname, "output");
let DEFAULT_KEYFILE;
if (IS_FROM_RENDER) {
  DEFAULT_KEYFILE = "/etc/secrets/google.json";
} else {
  DEFAULT_KEYFILE = path.resolve(__dirname, "google.json");
}
const PORT = process.env.PORT || 3000;
const upload = multer({ storage: multer.memoryStorage() });

const auth = new google.auth.GoogleAuth({
  keyFile: DEFAULT_KEYFILE,
  scopes: ["https://www.googleapis.com/auth/spreadsheets"],
});

const readSheet = async (spreadsheetId, range) => {
  const client = await auth.getClient();
  const sheets = google.sheets({ version: "v4", auth: client });
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range,
  });
  return res.data.values || [];
};

const ensureOutputDir = (outputDir) => {
  if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir, { recursive: true });
  }
};

const buildWorksheetColumns = () => {
  const baseColumns = [
    { header: "No", key: "no" },
    { header: "NIP", key: "nip" },
    { header: "Pegawai", key: "nama" },
    { header: "Tahun", key: "tahun" },
    { header: "Bulan", key: "bulan" },
  ];

  const tanggalColumns = [];
  for (let i = 1; i <= 31; i++) {
    tanggalColumns.push({ header: `Tanggal ${i} In`, key: `tanggal_${i}_in` });
    tanggalColumns.push({
      header: `Tanggal ${i} Rest`,
      key: `tanggal_${i}_rest`,
    });
    tanggalColumns.push({
      header: `Tanggal ${i} Out`,
      key: `tanggal_${i}_out`,
    });
    tanggalColumns.push({
      header: `Tanggal ${i} Status`,
      key: `tanggal_${i}_status`,
    });
    tanggalColumns.push({
      header: `Tanggal ${i} Late`,
      key: `tanggal_${i}_late`,
    });
  }

  return [...baseColumns, ...tanggalColumns];
};

const buildPresenceWorkbook = (data) => {
  const currentDate = new Date();
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("Sheet 1");

  worksheet.columns = [
    { header: "NO", key: "no", width: 10 },
    { header: "NIP", key: "nip", width: 20 },
    { header: "NAMA", key: "nama", width: 25 },
    { header: "PANGKAT/GOL. RUANG", key: "pangkatGolRu", width: 25 },
    { header: "JABATAN", key: "jabatan", width: 50 },
    { header: "H", key: "hadirNormal", width: 5 },
    { header: "TK", key: "tanpaKeterangan", width: 5 },
    { header: "TL", key: "akumulasiKeterlambatan", width: 5 },
    { header: "TAT", key: "tidakAbsenTengah", width: 5 },
    { header: "DL", key: "dinasLuar", width: 5 },
    { header: "TB", key: "tugasBelajar", width: 5 },
    { header: "CT", key: "cutiTahunan", width: 5 },
    { header: "CM", key: "cutiMelahirkan", width: 5 },
    { header: "CB", key: "cutiBesar", width: 5 },
    { header: "CS", key: "cutiSakit", width: 5 },
    { header: "CAP", key: "cutiAlasanPenting", width: 5 },
    { header: "CTLN", key: "cutiDiLuarTanggunganNegara", width: 5 },
    { header: "CH", key: "cutiHaji", width: 5 },
  ];

  const headerRow = worksheet.getRow(1);
  headerRow.eachCell((cell) => {
    cell.border = {
      top: { style: "thin", color: { argb: "FF000000" } },
      left: { style: "thin", color: { argb: "FF000000" } },
      bottom: { style: "medium", color: { argb: "FF000000" } },
      right: { style: "thin", color: { argb: "FF000000" } },
    };
    cell.font = { bold: true };
  });

  data.forEach((item, index) => {
    const row = worksheet.addRow({ ...item, no: index + 1 });
    row.eachCell((cell) => {
      cell.border = {
        top: { style: "thin", color: { argb: "FF000000" } },
        left: { style: "thin", color: { argb: "FF000000" } },
        bottom: { style: "thin", color: { argb: "FF000000" } },
        right: { style: "thin", color: { argb: "FF000000" } },
      };
    });
  });

  worksheet.addRow({});
  const keteranganDanLokasiRow = worksheet.addRow({
    no: "Keterangan",
    cutiTahunan: `Kota Bima, ${currentDate.toLocaleDateString()}`,
  });
  keteranganDanLokasiRow.getCell("no").font = {
    bold: true,
  };

  worksheet.addRow({
    no: "H",
    nip: ": Hadir",
    cutiTahunan: "Kepala Puskesmas Rasanae Timur",
  });
  worksheet.addRow({ no: "TK", nip: ": Tanpa Keterangan" });
  worksheet.addRow({ no: "TL", nip: ": Konversi Akumulasi Keterlambatan" });
  worksheet.addRow({ no: "DL", nip: ": Dinas Luar" });
  worksheet.addRow({ no: "TB", nip: ": Tugas Belajar" });
  worksheet.addRow({ no: "CT", nip: ": Cuti Tahunan" });
  worksheet.addRow({ no: "CM", nip: ": Cuti Melahirkan" });
  worksheet.addRow({
    no: "CB",
    nip: ": Cuti Besar",
    cutiTahunan: "H. ABDULLAH S.KM.",
  });
  worksheet.addRow({
    no: "CS",
    nip: ": Cuti Sakit",
    cutiTahunan: "NIP. 196907041989031005",
  });
  worksheet.addRow({ no: "CAP", nip: ": Cuti Alasan Penting" });
  worksheet.addRow({
    no: "CTLN",
    nip: ": Cuti Di Luar Tanggungan Negara",
  });
  worksheet.addRow({ no: "CH", nip: ": Cuti Haji" });

  return workbook;
};

const createExcel = async (data, filePath) => {
  const workbook = buildPresenceWorkbook(data);
  await workbook.xlsx.writeFile(filePath);
};

const createExcelBuffer = async (data) => {
  const workbook = buildPresenceWorkbook(data);
  return workbook.xlsx.writeBuffer();
};

const validateMonth = (month) => {
  const parsed = moment(month, ["YYYY-MM", "YYYY-MM-DD"], true);
  if (!parsed.isValid()) {
    throw new Error("Invalid month format. Use YYYY-MM or YYYY-MM-DD.");
  }
  return parsed;
};

const parseDate = (dateString) => {
  if (!dateString) return null;

  // Splits "25/12/2026" into ["25", "12", "2026"]
  const [day, month, year] = dateString.split("/");

  // Reconstructs into ISO format "YYYY-MM-DD"
  return new Date(`${year}-${month}-${day}`);
};

const calculatePresence = async ({
  month,
  spreadsheetId = DEFAULT_SPREADSHEET_ID,
  masterRange = MASTER_RANGE,
  izinRange = IZIN_RANGE,
  inputFile = DEFAULT_INPUT_FILE,
  outputDir = OUTPUT_DIR,
  workbookBuffer = null,
  writeExcel = false,
} = {}) => {
  if (!month) {
    throw new Error("Missing month parameter. Use ?month=YYYY-MM.");
  }

  const monthMoment = validateMonth(month);
  const start = monthMoment.startOf("month").toDate();
  const end = monthMoment.endOf("month").toDate();

  const workbook = new ExcelJS.Workbook();
  if (workbookBuffer) {
    await workbook.xlsx.load(workbookBuffer);
  } else {
    await workbook.xlsx.readFile(inputFile);
  }
  const worksheet = workbook.getWorksheet(1);
  worksheet.columns = buildWorksheetColumns();

  const readSheetValue = await readSheet(spreadsheetId, masterRange);
  const izinValues = await readSheet(spreadsheetId, izinRange);

  const izinList = izinValues.map((izin) => {
    const [nip, nama, jenisIzin, startDate, endDate] = izin;
    return {
      nip,
      nama,
      jenisIzin,
      startDate: parseDate(startDate),
      endDate: parseDate(endDate),
    };
  });

  const hd = new Holidays("ID");
  const holidays = hd
    .getHolidays()
    .filter((h) => new Date(h.date) >= start && new Date(h.date) <= end);

  const results = [];

  for (const empployee of readSheetValue) {
    if (!empployee || empployee.length < 2) {
      continue;
    }

    const [nip, nama, pangkatGolRu, jabatan] = empployee;
    const colValues = worksheet.getColumn("nip").values;
    const rowIndex = colValues.indexOf(nip);

    if (rowIndex <= 0) {
      continue;
    }

    const targetRow = worksheet.getRow(rowIndex);
    let hadirNormal = 0;
    let tanpaKeterangan = 0;
    let totalTerlambat = 0;
    let tidakAbsenTengah = 0;
    let dinasLuar = 0;
    let tugasBelajar = 0;
    let cutiTahunan = 0;
    let cutiMelahirkan = 0;
    let cutiBesar = 0;
    let cutiSakit = 0;
    let cutiAlasanPenting = 0;
    let cutiDiLuarTanggunganNegara = 0;
    let cutiHaji = 0;

    for (let dt = new Date(start); dt <= end; dt.setDate(dt.getDate() + 1)) {
      const dateNumber = dt.getDate();
      const isHoliday = holidays.some(
        (h) => new Date(h.date).toDateString() === dt.toDateString(),
      );
      const isSunday = dt.getDay() === 0;

      const inCell = targetRow.getCell(`tanggal_${dateNumber}_in`);
      const restCell = targetRow.getCell(`tanggal_${dateNumber}_rest`);
      const lateCell = targetRow.getCell(`tanggal_${dateNumber}_late`);

      const status =
        inCell.value || restCell.value || lateCell.value ? "HN" : "TK";
      const lateMinutes = Number(lateCell.value || 0);

      if (!isHoliday && !isSunday) {
        const findIzinList = izinList.find(
          (izin) =>
            izin.nip === nip && dt >= izin.startDate && dt <= izin.endDate,
        );

        if (findIzinList) {
          if (findIzinList.jenisIzin === "CUTI_TAHUNAN") cutiTahunan++;
          else if (findIzinList.jenisIzin === "DINAS_LUAR") dinasLuar++;
          else if (findIzinList.jenisIzin === "TUGAS_BELAJAR") tugasBelajar++;
          else if (findIzinList.jenisIzin === "CUTI_MELAHIRKAN")
            cutiMelahirkan++;
          else if (findIzinList.jenisIzin === "CUTI_BESAR") cutiBesar++;
          else if (findIzinList.jenisIzin === "CUTI_SAKIT") cutiSakit++;
          else if (findIzinList.jenisIzin === "CUTI_ALASAN_PENTING")
            cutiAlasanPenting++;
          else if (findIzinList.jenisIzin === "CUTI_DI_LUAR_TANGGUNGAN_NEGARA")
            cutiDiLuarTanggunganNegara++;
          else if (findIzinList.jenisIzin === "CUTI_HAJI") cutiHaji++;
        } else {
          if (status === "HN") hadirNormal++;
          else tanpaKeterangan++;

          totalTerlambat += lateMinutes;

          if (!restCell.value) {
            tidakAbsenTengah++;
          }
        }
      }
    }

    const kalkulasiKeterlambatan = totalTerlambat / 60;
    const desimal = kalkulasiKeterlambatan % 1;
    const akumulasiKeterlambatan =
      desimal >= 0.75
        ? Math.floor(kalkulasiKeterlambatan) + 1
        : Math.floor(kalkulasiKeterlambatan);

    results.push({
      nama,
      nip,
      pangkatGolRu,
      jabatan,
      hadirNormal,
      tanpaKeterangan,
      akumulasiKeterlambatan,
      tidakAbsenTengah,
      dinasLuar,
      tugasBelajar,
      cutiTahunan,
      cutiMelahirkan,
      cutiBesar,
      cutiSakit,
      cutiAlasanPenting,
      cutiDiLuarTanggunganNegara,
      cutiHaji,
    });
  }

  let filePath;
  if (writeExcel) {
    ensureOutputDir(outputDir);
    filePath = path.join(
      outputDir,
      `presence_summary_${monthMoment.format("YYYY-MM")}_${new Date()
        .toISOString()
        .slice(0, 10)}.xlsx`,
    );
    await createExcel(results, filePath);
  }

  return { results, filePath };
};

const app = express();
app.use(express.json());

app.get("/api/health", (req, res) => {
  res.json({ status: "ok", timestamp: new Date().toISOString() });
});

app.get("/api/presence", async (req, res) => {
  try {
    const month = req.query.month || moment().format("YYYY-MM");
    const { results } = await calculatePresence({ month, writeExcel: false });
    res.json({ month, results });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

app.post("/api/presence", upload.single("file"), async (req, res) => {
  try {
    if (!req.file) {
      return res
        .status(400)
        .json({ error: "Missing file upload. Use field name 'file'." });
    }

    const month =
      req.body.month || req.query.month || moment().format("YYYY-MM");
    const { results } = await calculatePresence({
      month,
      workbookBuffer: req.file.buffer,
      writeExcel: false,
    });

    const excelBuffer = await createExcelBuffer(results);
    const filename = `presence_summary_${month}.xlsx`;

    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    );
    res.setHeader("Content-Disposition", `attachment; filename="${filename}"`);
    res.send(excelBuffer);
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

app.get("/api/presence/download", async (req, res) => {
  try {
    const month = req.query.month || moment().format("YYYY-MM");
    const { filePath } = await calculatePresence({ month, writeExcel: true });
    if (!filePath) {
      return res.status(500).json({ error: "Failed to generate Excel file." });
    }
    res.download(filePath, path.basename(filePath));
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

const startServer = () => {
  app.listen(PORT, () => {
    console.log(`API server is running at http://localhost:${PORT}`);
    console.log(`GET /api/presence?month=YYYY-MM`);
    console.log(`GET /api/presence/download?month=YYYY-MM`);
  });
};

const runCli = async () => {
  const month = process.argv[2];
  if (!month) {
    console.error("Usage: node index.js YYYY-MM");
    process.exit(1);
  }
  const { filePath } = await calculatePresence({ month, writeExcel: true });
  console.log(`Saved presence summary to: ${filePath}`);
  process.exit(0);
};

if (process.argv.length > 2) {
  await runCli();
} else {
  startServer();
}

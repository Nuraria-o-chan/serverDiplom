import express from "express";
import multer from "multer";
import XLSX from "xlsx";
import cors from "cors";
import path from "path";
import { fileURLToPath } from "url";
import fs from "fs";

const app = express();
const port = 3001;

app.use(cors());

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

app.use(express.static(path.join(__dirname, "public")));

const storage = multer.diskStorage({
  destination: function (req, file, cb) {
    cb(null, "uploads/");
  },
  filename: function (req, file, cb) {
    cb(null, file.originalname);
  },
});

const upload = multer({ storage: storage });

const uploadDir = "./uploads";
if (!fs.existsSync(uploadDir)) {
  fs.mkdirSync(uploadDir);
}

function parseDate(dateString) {
  if (!dateString) return null;
  const parts = dateString.split(".");
  return new Date(parts[2], parts[1] - 1, parts[0]);
}

app.post("/upload", upload.single("excelFile"), (req, res) => {
  try {
    if (!req.file) {
      return res
        .status(400)
        .send(
          `<script>alert("Файл не был загружен."); window.location.href = '/';</script>`
        );
    }

    const filePath = req.file.path;
    const workbook = XLSX.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];

    function getLastUsedRow(worksheet) {
      if (!worksheet || !worksheet["!ref"]) return 0;
      const range = XLSX.utils.decode_range(worksheet["!ref"]);
      return range.e.r + 1;
    }

    const mergedRanges =
      worksheet["!merges"]?.filter((x) => x["s"]["r"] === 0) || [];

    if (!mergedRanges || mergedRanges.length === 0) {
      console.warn("No merged ranges found starting on row 0.");
      return res
        .status(400)
        .send(
          `<script>alert("No merged ranges found starting on row 0."); window.location.href = '/';</script>`
        );
    }

    const allGroups = [];

    for (let i = 0; i < mergedRanges.length; i++) {
      const mergedRange = mergedRanges[i];

      const groupName =
        worksheet[
          XLSX.utils.encode_cell({ r: mergedRange.s.r, c: mergedRange.s.c })
        ]?.v?.toString() || "Без названия";

      const newGroup = {
        Name: groupName,
        days: [],
      };

      const mergedDayNames =
        worksheet["!merges"]?.filter(
          (x) => x["s"]["c"] === mergedRange.s.c && x["s"]["r"] > 1
        ) || [];

      for (let j = 0; j < mergedDayNames.length; j++) {
        const element = mergedDayNames[j];

        const day = {
          data:
            worksheet[
              XLSX.utils.encode_cell({ r: element.s.r, c: element.s.c })
            ]?.v || "Дата не указана",
          paras: [],
        };

        for (let dayRow = element.s.r; dayRow <= element.e.r; dayRow++) {
          const para = {};
          para.number =
            worksheet[
              XLSX.utils.encode_cell({ r: dayRow, c: mergedRange.s.c + 1 })
            ]?.v || "";
          para.name =
            worksheet[
              XLSX.utils.encode_cell({ r: dayRow, c: mergedRange.s.c + 2 })
            ]?.v || "";
          para.disciplina =
            worksheet[
              XLSX.utils.encode_cell({ r: dayRow, c: mergedRange.s.c + 3 })
            ]?.v || "";
          para.prepod =
            worksheet[
              XLSX.utils.encode_cell({ r: dayRow, c: mergedRange.s.c + 4 })
            ]?.v || "";
          para.kab =
            worksheet[
              XLSX.utils.encode_cell({ r: dayRow, c: mergedRange.s.c + 5 })
            ]?.v || "";

          if (para.number || para.disciplina || para.prepod || para.kab) {
            day.paras.push(para);
          }
        }
        newGroup.days.push(day);
      }
      console.log(newGroup);

      newGroup.days.sort((a, b) => {
        const dateA = parseDate(a.data);
        const dateB = parseDate(b.data);
        if (dateA && dateB) {
          return dateA.getTime() - dateB.getTime();
        } else {
          return 0;
        }
      });

      allGroups.push(newGroup);
    }
    console.log(allGroups);

    fs.unlinkSync(filePath);

    res.json(allGroups);
  } catch (error) {
    console.error("Ошибка при обработке файла:", error);

    res
      .status(500)
      .send(
        `<script>alert("Ошибка при обработке файла."); window.location.href = '/';</script>`
      );
  }
});

app.listen(port, () => {
  console.log(`Сервер запущен на порту ${port}`);
});

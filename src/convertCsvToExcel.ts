import fs from "fs";
import path from "path";
import ExcelJS from "exceljs";
import iconv from "iconv-lite";
import { parse } from "csv-parse/sync";

function findCsvFiles(dir: string, fileList: string[] = []): string[] {
  const files = fs.readdirSync(dir);

  files.forEach((file) => {
    const filePath = path.join(dir, file);
    if (fs.statSync(filePath).isDirectory()) {
      fileList = findCsvFiles(filePath, fileList);
    } else if (file.endsWith(".csv") && fs.statSync(filePath).size > 0) {
      fileList.push(filePath);
    }
  });

  return fileList;
}

async function convertCsvToExcel(
  csvFolderPath: string,
  outputExcelFilePath: string,
) {
  const workbook = new ExcelJS.Workbook();
  const csvFiles = findCsvFiles(csvFolderPath);

  if (csvFiles.length === 0) {
    throw new Error(`No CSV files found in the directory: ${csvFolderPath}`);
  }

  for (const csvFilePath of csvFiles) {
    const csvContent = fs.readFileSync(csvFilePath);
    const csvContentUtf8 = iconv.decode(csvContent, "UTF-8");

    const records = parse(csvContentUtf8, {
      columns: true,
      skip_empty_lines: true,
    });

    if (records.length === 0) {
      console.log(`Skipping empty content in file: ${csvFilePath}`);
      continue;
    }

    // ファイルパスから一意のワークシート名を生成
    const worksheetName = csvFilePath
      .substring(csvFolderPath.length + 1) // ベースディレクトリのパスを取り除く
      .replace(/\.csv$/, "") // .csv拡張子を取り除く
      .replace(/[\/\\]/g, "_"); // ディレクトリ区切り文字をアンダースコアに置換

    const worksheet = workbook.addWorksheet(worksheetName);
    worksheet.columns = Object.keys(records[0]).map((key) => ({
      header: key,
      key,
    }));

    records.forEach((record: any) => {
      worksheet.addRow(record);
    });
  }

  await workbook.xlsx.writeFile(outputExcelFilePath);
}

// コマンドライン引数からディレクトリと出力ファイルパスを取得
const args = process.argv.slice(2);
if (args.length !== 2) {
  console.error(
    "Usage: node convertCsvToExcel.js <csv_directory_path> <output_excel_file_path>",
  );
  process.exit(1);
}

const [csvFolderPath, outputExcelFilePath] = args;

convertCsvToExcel(csvFolderPath, outputExcelFilePath)
  .then(() => console.log("Conversion completed."))
  .catch((err) => {
    console.error("Conversion failed:", err);
    process.exit(1);
  });

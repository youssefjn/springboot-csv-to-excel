package excel.csv.excelcsv.service;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.Reader;
import java.nio.file.Files;
import java.nio.file.Paths;
import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVParser;
import org.apache.commons.csv.CSVRecord;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;


@Service
public class ConverterService {
    private static String[] columns = { "id", "nom", "prenom", "poste", "salaire" };

    public void convertCsvToExcel(MultipartFile csvFile) throws IOException {
        File convFile = new File(System.getProperty("java.io.tmpdir") + "/" + csvFile.toString());
        csvFile.transferTo(convFile);
        String path = convFile.getPath();
        try (

                Reader reader = Files.newBufferedReader(Paths.get(path));
                CSVParser csvParser = new CSVParser(reader, CSVFormat.DEFAULT.withFirstRecordAsHeader().withTrim());) {
            Workbook workbook = new XSSFWorkbook();

            Sheet sheet = workbook.createSheet("Employee");
            Font headerFont = workbook.createFont();
            headerFont.setBold(true);
            headerFont.setFontHeightInPoints((short) 14);
            headerFont.setColor(IndexedColors.RED.getIndex());
            CellStyle headerCellStyle = workbook.createCellStyle();
            headerCellStyle.setFont(headerFont);
            Row headerRow = sheet.createRow(0);
            for (int i = 0; i < columns.length; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(columns[i]);
                cell.setCellStyle(headerCellStyle);
            }
            int rowNum = 1;
            for (CSVRecord csvRecord : csvParser) {
                String id = csvRecord.get("id");
                String nom = csvRecord.get("nom");
                String prenom = csvRecord.get("prenom");
                String poste = csvRecord.get("poste");
                String salaire = csvRecord.get("salaire");
                Row row = sheet.createRow(rowNum++);

                row.createCell(0)
                        .setCellValue(id);

                row.createCell(1)
                        .setCellValue(nom);
                row.createCell(2)
                        .setCellValue(prenom);
                row.createCell(3)
                        .setCellValue(poste);
                row.createCell(4)
                        .setCellValue(salaire);
            }
            for (int i = 0; i < columns.length; i++) {
                sheet.autoSizeColumn(i);
            }

            // Write the output to a file
            FileOutputStream fileOut = new FileOutputStream("poi-generated-file.xlsx");
            workbook.write(fileOut);
            fileOut.close();

            // Closing the workbook
            workbook.close();
        }
    }
}
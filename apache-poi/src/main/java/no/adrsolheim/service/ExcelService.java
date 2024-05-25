package no.adrsolheim.service;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.Map;

public class ExcelService {

    SXSSFWorkbook wb;
    public void generateExcelWorkbook(Map<String, Object> columnData) throws IOException {
        wb = new SXSSFWorkbook(-1);
        Sheet sheet = wb.createSheet();
        ((SXSSFSheet) sheet).trackAllColumnsForAutoSizing();
        createHeader(sheet, columnData.keySet().stream().toList());
        for (int i = 0; i < columnData.keySet().size(); i++) {
            sheet.autoSizeColumn(i);
        }

        ((SXSSFSheet) sheet).untrackAllColumnsForAutoSizing();
        FileOutputStream out = new FileOutputStream("sxssf.xlsx");
        wb.write(out);
        out.close();

        wb.dispose();
    }

    private void createHeader(Sheet sheet, List<String> columns) {
        Row row = sheet.createRow(0);
        Font font = wb.createFont();
        font.setFontName("Arial");
        font.setFontHeightInPoints((short) 12);
        font.setColor(IndexedColors.BLACK.getIndex());
        CellStyle cellStyle = wb.createCellStyle();
        cellStyle.setFont(font);

        for (int i = 0; i < columns.size(); i++) {
            Cell cell = row.createCell(i);
            cell.setCellValue(columns.get(i));
            cell.setCellStyle(cellStyle);
        }

    }
}

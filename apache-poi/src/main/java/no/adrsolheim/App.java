package no.adrsolheim;

import no.adrsolheim.service.ExcelService;

import java.io.IOException;
import java.util.LinkedHashMap;
import java.util.Map;

/**
 * Hello world!
 *
 */
public class App {
    public static void main( String[] args ) throws IOException {
        ExcelService excelService = new ExcelService();
        LinkedHashMap<String, Object> rows = new LinkedHashMap<>();
        rows.put("NAME", "Jon");
        rows.put("AGE", 20);
        rows.put("FILL_PATTERN", "Solid");

        excelService.generateExcelWorkbook(rows);
    }
}

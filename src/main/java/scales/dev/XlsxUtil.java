package scales.dev;

import org.apache.poi.ss.usermodel.*;

import java.io.IOException;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.Locale;
import java.util.Map;

// ai slop code, i'll rewrite myself whenever i have time ig
// for later: https://www.baeldung.com/java-microsoft-excel
public class XlsxUtil {
  public static ArrayList<Map<String, String>> readXlsx(Path file, boolean header) {
    try (Workbook workbook = WorkbookFactory.create(file.toFile())) {

      Sheet sheet = workbook.getSheetAt(0);

      DataFormatter formatter = new DataFormatter(Locale.ROOT);
      FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();

      ArrayList<Map<String, String>> excel = new ArrayList<>();

      int startRow = header ? 1 : 0;
      int columnCount = 0;

      if (header) {
        Row headerRow = sheet.getRow(0);
        if (headerRow == null) return excel;
        columnCount = headerRow.getLastCellNum();
      } else {
        // Find the maximum number of columns
        for (Row row : sheet) {
          columnCount = Math.max(columnCount, row.getLastCellNum());
        }
      }

      // Read headers
      ArrayList<String> headers = new ArrayList<>();

      if (header) {
        Row headerRow = sheet.getRow(0);

        for (int i = 0; i < columnCount; i++) {
          Cell cell = headerRow.getCell(i, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
          headers.add(cell == null ? "" : formatter.formatCellValue(cell, evaluator));
        }
      }

      // Read rows
      for (int r = startRow; r <= sheet.getLastRowNum(); r++) {
        Row row = sheet.getRow(r);

        if (row == null) continue;

        Map<String, String> record = new LinkedHashMap<>();

        for (int c = 0; c < columnCount; c++) {
          Cell cell = row.getCell(c, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);

          String value = cell == null
                  ? ""
                  : formatter.formatCellValue(cell, evaluator);

          record.put(
                  header ? headers.get(c) : String.valueOf(c),
                  value
          );
        }

        excel.add(record);
      }

      return excel;

    } catch (IOException e) {
      throw new RuntimeException(e);
    }
  }
}
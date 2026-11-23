package scales.dev;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.nio.file.Path;
import java.util.*;

// https://www.baeldung.com/java-microsoft-excel
public class XlsxUtil {
  private static final DataFormatter FORMATTER = new DataFormatter(Locale.ROOT);

  // some really ugly code :pensive:
  public static scales.dev.Workbook readXlsx(Path path, boolean hasHeaders) {
    try{
      FileInputStream file = new FileInputStream(path.toFile());
      Workbook workbook = new XSSFWorkbook(file);
      FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();

      Map<String, Worksheet> worksheets = new LinkedHashMap<>();

      int sheetCount = workbook.getNumberOfSheets();
      for(int i = 0; i < sheetCount; i++) {
        Sheet sheet = workbook.getSheetAt(i);

        ArrayList<Map<String, String>> rows = new ArrayList<>();
        ArrayList<String> headers = new ArrayList<>();
        int columns = 0;

        if (!hasHeaders) {
          for (Row row : sheet) {
            // "getLastCellNum" is a short, which is weird, cause im sure u can have more than 32k columns
            columns = Math.max(columns, row.getLastCellNum());
          }
        }

        for (Row row : sheet) {
          if (headers.isEmpty() && hasHeaders) {
            for (Cell cell : row) {
              headers.add(cell.toString());
            }

            columns = headers.size();
            continue;
          }

          Map<String, String> values = new LinkedHashMap<>();

          for(int column = 0; column < columns; column++) {
            Cell cell = row.getCell(column, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
            String header = headers.size() <= column
                    ? Integer.toString(column)
                    : headers.get(column);

            String value = cell == null
                    ? ""
                    : FORMATTER.formatCellValue(cell, evaluator);

            values.put(header, value);
          }

          rows.add(values);
        }

        worksheets.put(sheet.getSheetName(), new Worksheet(rows));
      }

      return new scales.dev.Workbook(worksheets);
    } catch (IOException e) {
      throw new RuntimeException(e);
    }
  }
}
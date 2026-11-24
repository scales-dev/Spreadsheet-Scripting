package scales.dev;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.file.Path;
import java.time.Instant;
import java.time.format.DateTimeParseException;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

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

  // todo: working out data types when writing
  public static void writeXlsx(Path file, scales.dev.Workbook workbook) {
    try{
      Workbook workbookWriter = new XSSFWorkbook();
      FileOutputStream outputStream = new FileOutputStream(file.toFile());

      Set<Map.Entry<String, Worksheet>> worksheets = workbook.worksheets.entrySet();
      for (Map.Entry<String, Worksheet> worksheet : worksheets) {
        int rows = 0;
        LinkedHashSet<String> headers = new LinkedHashSet<>();

        Sheet sheetWriter = workbookWriter.createSheet(worksheet.getKey());
        Worksheet sheet = worksheet.getValue();

        sheet.rows.forEach(product -> headers.addAll(product.keySet()));

        Row headerRow = sheetWriter.createRow(rows++);
        int headerCellIndex = 0;
        for (String header : headers) {
          Cell cell = headerRow.createCell(headerCellIndex++);
          cell.setCellValue(header);
        }

        for (Map<String, String> inputRow : sheet.rows) {
          Row outputRow = sheetWriter.createRow(rows++);

          int cellIndex = 0;
          for (String cellValue : inputRow.values()) {
            Cell cell = outputRow.createCell(cellIndex++);
            setCellValue(cellValue, cell);
          }
        }
      }

      workbookWriter.write(outputStream);
      workbookWriter.close();
    } catch (IOException e) {
      throw new RuntimeException(e);
    }
  }

  // todo: .'s and ,'s are swapped in some locales, i remember pokemon made that mistake once.
  private enum CellType {
    MONEY {
      // matches any unusual character followed by any amount of numbers, an optional decimal then
      // optional more numbers
      private static final Pattern MONEY_REGEX =
          Pattern.compile("[^\\s0-9a-zA-Z\\-.]([\\d,]*\\.?[\\d,]*)");

      @Override
      public boolean setCell(String s, Cell cell) {
        Matcher matcher = MONEY_REGEX.matcher(s);
        if (!matcher.matches()) return false;

        cell.setCellValue(Double.parseDouble(CellType.fixNumber(matcher.group(1))));
        return true;
      }
    },
    NUMBER {
      // ignores numbers starting with 0 if they dont have a decimal after
      private static final Pattern NUMBER_REGEX =
          Pattern.compile("-?(([1-9][\\d,]*(\\.[\\d,]*)?)|(0\\.[\\d,]++))");

      @Override public boolean setCell(String s, Cell cell) {
        Matcher matcher = NUMBER_REGEX.matcher(s);
        if (!matcher.matches()) return false;

        cell.setCellValue(Double.parseDouble(CellType.fixNumber(matcher.group(1))));
        return true;
      }
    },
    BOOLEAN {
      private static final Pattern BOOLEAN_REGEX = Pattern.compile("(?i)(true|false)");

      @Override public boolean setCell(String s, Cell cell) {
        Matcher matcher = BOOLEAN_REGEX.matcher(s);
        if (!matcher.matches()) return false;

        cell.setCellValue(Boolean.parseBoolean(matcher.group(1)));
        return true;
      }
    },
    FORMULA {
      @Override public boolean setCell(String s, Cell cell) {
        if (!s.startsWith("=")) return false;

        cell.setCellFormula(s.substring(1));
        return true;
      }
    },
    DATE {
      @Override public boolean setCell(String s, Cell cell) {
        try{
          Date.from(Instant.parse(s));
          return true;
        } catch (DateTimeParseException e) {
          return false;
        }
      }
    },
    STRING {
      /// assumes all other checks failed, is a fallback
      @Override public boolean setCell(String s, Cell cell) {
        cell.setCellValue(s);
        return true;
      }
    };

    public abstract boolean setCell(String s, Cell cell);

    private static String fixNumber(String match) {
      return match.replace(",", "");
    }
  }

  public static void setCellValue(String s, Cell cell) {
    for (CellType cellType : CellType.values()) {
      if (cellType.setCell(s, cell)) {
        System.out.println(s + " : " + cellType.name());
        return;
      }
    }
  }
}
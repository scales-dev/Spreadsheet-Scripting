package scales.dev;

import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Map;
import java.util.stream.Collectors;

public class SpreadsheetUtil {
  public enum SpreadsheetType {
    CSV {
      @Override
      public ArrayList<Map<String, String>> read(Path file, boolean header) {
        return CsvUtil.readCsv(file, header);
      }

      @Override
      public void write(Path file, ArrayList<Map<String, String>> csv) {
        CsvUtil.writeCsv(file, csv);
      }
    },
    XLSX {
      @Override
      public ArrayList<Map<String, String>> read(Path file, boolean header) {
        return XlsxUtil.readXlsx(file, header);
      }
    },
    XLS {
      @Override public ArrayList<Map<String, String>> read(Path file, boolean header) {
        return XlsxUtil.readXlsx(file, header);
      }
    };

    public abstract ArrayList<Map<String, String>> read(Path file, boolean header);
    public void write(Path file, ArrayList<Map<String, String>> csv) {
      throw new RuntimeException("Writing this type is not yet implemented!");
    }

    public static SpreadsheetType getType(Path file) {
      String fileExtension = file.getFileName().toString().split("\\.")[1].toUpperCase();
      for (SpreadsheetType type : SpreadsheetType.values()) {
        if (!type.name().equals(fileExtension)) continue;

        return type;
      }

      String validExtensions = Arrays.stream(SpreadsheetType.values()).map(Enum::name).collect(Collectors.joining(", "));
      throw new RuntimeException("Invalid file extension, please input a file with an extension of the following: "+validExtensions+". excel can be used to convert between these, go to File -> Export");
    }
  }

  // todo: reading whole workbooks, not just single worksheets
  public static ArrayList<Map<String, String>> read(Path file, boolean header) {
    return SpreadsheetType.getType(file).read(file, header);
  }

  // todo: writing whole workbooks, not just single worksheets
  public static void write(Path file, ArrayList<Map<String, String>> csv) {
    SpreadsheetType.getType(file).write(file, csv);
  }
}

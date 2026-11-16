package scales.dev;

import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Map;

public class Main {
  public static void main(String[] args) {
    Path CSV_PATH = Path.of(args[0]);
    ArrayList<Map<String, String>> csv = SpreadsheetUtil.read(CSV_PATH, true);

    CsvUtil.writeCsv(CSV_PATH.getParent().resolve("OUTPUT.csv"), csv);
  }
}
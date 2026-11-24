package scales.dev;

import org.apache.poi.ss.usermodel.*;

import java.nio.file.Path;

public class Main {
  public static void main(String[] args) {
    Path INPUT = Path.of(args[0]);
    Workbook workbook = SpreadsheetUtil.read(INPUT, true);

    SpreadsheetUtil.write(INPUT.getParent().resolve("OUTPUT.xlsx"), workbook);
  }
}
package scales.dev;

import java.util.Map;

public class Workbook {
  public Map<String, Worksheet> worksheets;

  public Workbook(Map<String, Worksheet> worksheets) {
    this.worksheets = worksheets;
  }
}

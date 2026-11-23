package scales.dev;

import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVParser;
import org.apache.commons.csv.CSVPrinter;
import org.apache.commons.csv.CSVRecord;
import org.apache.commons.io.input.BOMInputStream;

import java.io.*;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.*;

public class CsvUtil {
  public static void writeCsv(Path file, Workbook workbook) {
    Set<Map.Entry<String, Worksheet>> worksheets = workbook.worksheets.entrySet();
    Path parent = file.getParent();
    Path outputParent = worksheets.size() == 1
            ? parent
            : parent.resolve(file.getFileName().toString().split("\\.")[0]);
    if (new File(outputParent.toUri()).mkdirs()) {
      System.out.println("Created new path: " + outputParent);
    }


    for (Map.Entry<String, Worksheet> worksheet : worksheets) {
      LinkedHashSet<String> headers = new LinkedHashSet<>();

      String name = worksheet.getKey() + ".csv";
      Worksheet sheet = worksheet.getValue();

      sheet.rows.forEach(product -> headers.addAll(product.keySet()));

      try{
        BufferedWriter writer = Files.newBufferedWriter(outputParent.resolve(name), StandardCharsets.UTF_8);
        CSVPrinter csvPrinter = new CSVPrinter(writer, CSVFormat.DEFAULT);

        csvPrinter.printRecord(headers);

        for (Map<String, String> item : sheet.rows) {
          csvPrinter.printRecord(headers.stream().map(item::get).toList());
        }

        writer.flush();
        writer.close();
      } catch (IOException e) {
        throw new RuntimeException(e);
      }
    }
  }

  public static Workbook readCsv(Path file, boolean header) {
    try {
      BufferedReader reader = new BufferedReader(new InputStreamReader(
              // removes weird characters like zero width spaces which are somehow getting into my headers
              BOMInputStream.builder().setPath(file).get(),
              StandardCharsets.UTF_8
      ));

      CSVFormat.Builder formatBuilder = CSVFormat.DEFAULT.builder();

      if (header) {
        formatBuilder = formatBuilder.setHeader().setSkipHeaderRecord(true);
      }
      CSVParser csvParser = CSVParser.parse(reader, formatBuilder.get());

      ArrayList<Map<String, String>> csv = new ArrayList<>();

      for (CSVRecord record : csvParser.getRecords()) {
        if (header) {
          csv.add(record.toMap());
        }
        else {
          Map<String, String> recordMap = new HashMap<>();

          for(int i = 0; i < record.size(); i++) {
            recordMap.put(String.valueOf(i), record.get(i));
          }

          csv.add(recordMap);
        }
      }

      csvParser.close();

      return new Workbook(Map.of("output", new Worksheet(csv)));
    } catch (IOException e) {
      throw new RuntimeException(e);
    }
  }
}
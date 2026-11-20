# Spreadsheet Scripting
A Java library for myself, that makes reading and writing spreadsheets slightly easier across projects.
Mainly because I do a lot of projects with spreadsheets, and I don't like reusing code.
I think it makes sense for any script I write to read a spreadsheet as an array of maps of String : String.
This library relies on
[Apache Commons CSV](https://mvnrepository.com/artifact/org.apache.commons/commons-csv) for csv's
and
[Apache POI](https://mvnrepository.com/artifact/org.apache.poi/poi) for xlsx and xls files
to read spreadsheets and return them in the format:
```json
[
    {
        "name": "Jack",
        "id": "0"
    },
    {
        "name": "John",
        "id": "1"
    }
]
```

### Usage
Thanks to [JitPak](https://jitpack.io/#scales-dev/Spreadsheet-Scripting)!
Replace TAG with a commit tag found [here](https://jitpack.io/#scales-dev/Spreadsheet-Scripting) 
or -SNAPSHOT to be on the bleeding edge

#### Gradle
```groovy
dependencies {
    implementation 'com.github.scales-dev:Spreadsheet-Scripting:Tag'
}
```
#### Maven
```xml
<dependency>
    <groupId>com.github.scales-dev</groupId>
    <artifactId>Spreadsheet-Scripting</artifactId>
    <version>Tag</version>
</dependency>
```

### Example
```java
void spreadsheetToCsv(Path input) {
  ArrayList<Map<String, String>> spreadsheet = SpreadsheetUtil.read(input, true);
  SpreadsheetUtil.write(input.getParent().resolve("output.csv"), spreadsheet); 
}
```

Or an entire example project: [Compress-Buying-List](https://github.com/scales-dev/Compress-Buying-List)
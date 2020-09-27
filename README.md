# Simple Excel read demo

Using Apache POI (ooxml) the Java program `ExcelReader` read file `salary.xlsx`. \
The Excel file has 3 sheets.

The demo does the following:
* Prints sheet count (3)
* Prints sheet names (loops `->` 2018, 2019, 2020)
* Prints last sheet name (2020)
* Prints last sheets content

### Last sheets content
```
Name	Age	Salary	
Sara	35.0	47000.0	
John	24.0	39000.0	
Brett	46.0	53000.0
```

**Note!** Column `Age` is formatted as _Number_ and `Salary` is formatted as _Currency_ in _Excel_

## Dependency

```
<dependency>
    <groupId>org.apache.poi</groupId>
    <artifactId>poi-ooxml</artifactId>
    <version>4.1.2</version>
</dependency>
```
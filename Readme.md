## Mudassar Ansari Excel Automation Template which populate data from sheet1 to sheet2

### Result Generation Examples

### This is an automation example to generate result for School in any semester, which are printed in-house. This is a Marco based excel sheet which consists of students data with subject, marks and grade. There is `Clear` button on sheet2 named `ResultSem1` which has to be pressed before starting the Result Generation Process.  There is a `Generate` button on sheet1 when clicked picks data from 2nd row and fill that into result sheet which is on sheet2. After populating data teachers can save the sheet2 in pdf format and printout taken on blank mark sheet with school name and other details. To generate next result we have to click on clear button on sheet2 that will take us to sheet1. For creating new result again we have to click `Generate` button until last students result generated. There is counter in cell `B57` which keep record of the no. of nest result to be processed. This file can generate only 52 result as the counter reaches 52 it will reset cell `B57` to value 2 (the first student record). 

### Note : `ResultSem1.xlsm` file is an MS Excel with inbuilt VB code, this file created in MS OFFICE to populate data from sheet1 to sheet2. Screenshots attached for reference.

### Discalmer : Hypothectical data is used to generate Result.

### For any further queries contact : mudassar.ansari@gmail.com

## Authors

- [@mudassar-web](https://github.com/mudassar-web)

## License

[MIT License](LICENSE)

### xlsx-populate
---
https://github.com/dtjohnson/xlsx-populate

```
npm install xlsx-populate
bower install xlsx-populate

```

```js
const XlsxPopulate = require('xlsx-populate');
XlsxPopulate.fromBlankAsync()
  .then(workbook => {
    workbook.sheet("Sheet1").cell("A1").value("This is neat!");
    return workbook.toFileAsync("./out.xlsx");
  });
  
const XlsxPopulate = require('xlsx-populate');  
XlsxPopulate.fromFileAsync("./Book1.xlsx")
  .then(workbook => {
    const value = workbook.sheet("Sheet1").cell("A1").value();
    console.log(value);
  });
  
const r = workbook.sheet(0).range("A1:C3");
r.value(5);
r.value([
  [1, 2, 3],
  [4, 5, 6],
  [7, 8, 9]
]);
r.value((cell, ri, ci, range) => Math.random());

const values = workbook.sheet("Sheet1").usedRange().value();

workbook.sheet(0).cell("A1").value([
  [1, 2, 3],
  [4, 5, 6],
  [7, 8, 9]
]);

sheet.column("B").width(25).hidden(false);
const cell = sheet.row(5).cell(3);

const sheet1 = workbook.sheet(0);
const sheet2 = workbook.sheet("Sheet2");
const sheets = workbook.sheets();

const newSheet1 = workbook.addSheet('New 1');
const newSheet2 = workbook.addSheet('New 2', 1);
const newSheet3 = workbook.addSheet('New 3', 'Sheet1');
const sheet = workbook.sheet('Sheet1');
const newSheet4 = workbook.addSheet('New 4', sheet);

const sheet = workbook.sheet(0).name("new sheet name");
workbook.moveSheet("Sheet1");
workbook.moveSheet("Sheet1", 2);
workbook.moveSheet("Sheet1", "Sheet2");
sheet.move("Sheet2");

workbook.deleteSheet("Sheet1");
workbook.deleteSheet(2);
workbook.sheet(0).delete();

const sheet = workbook.activeSheet();
sheet.active()
sheet.active(true);
workbook.activeSheet("Sheet2");

workbook.definedName("some name").value(5);  
workbook.sheet(0).defineName("some other name").value("foo");

workbook.definedName("some name", "TRUE");
workbook.sheet(0).definedName("some name", null);

workbook.find("foo", "bar");
workbook.find("foo");
workbook.sheet(0).find("foo");
workbook.sheet("Sheet1").cell("A1").find("foo");

workbook.find(/[a-z]+/g, match => match.toUpperCase());






















```

```
```



### xlsx-populate
---
https://github.com/dtjohnson/xlsx-populate

```
npm install xlsx-populate
bower install xlsx-populate

npm install -g gulp
npm install
gulp
gulp build
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

cell.style("bold", true);
cell.style({ bold: true, italic: true });
const bold = cell.style("bold");
const styles = cell.style(["bold", "italic"]);

range.style("bold", true);
range.style("bold", [[true, false], [false, true]]);
range.style("bold", (cell, ri, ci, range) => Math.random() > 0.5);
range.style({
  bold: true,
  italic: [[true, false], [false, true]],
  underline: (cell, ri, ci, range) => Math.random() > 0.5
});

sheet.row(1).style("bold", true);
sheet.column("A").style({ bold: true, italic: true });
const bold = sheet.column(3).style("bold");
const styles = sheet.row(5).style(["bold", "italic"]);

cell.style("fill", {
  type: "pettern",
  pattern: "darkDown",
  foreground: {
    rgb: "ff0000"
  },
  background: {
    theme: 3,
    tint: 0.4
  }
});

cell.style("fill", "0000ff");
const fill = cell.style("fill");
/*
{
  type: "solid",
  color: {
    rgb: "0000ff"
  }
}
*/

cell.style("numberFormat", "0.00");

cell.value(new Date(2017, 1, 22)).style("numberFormat", "dddd, mmmm dd, yyyy");

const num = cell.value();
const date = XlsxPopulate.numberToDate(num);

cell.dataValidation({
  type: 'list',
  allowBlank: false,
  showInputMessage: false,
  prompt: false,
  promptTitle: 'String',
  showErrorMessage: false,
  error: 'String',
  errorTitle: 'String',
  operator: 'String',
  formula1: '$A:$A',
  formula2: 'String'
});
cell.dataValidation('$A:$A');
const obj = cell.dataBalidation();
cell.dataValidation(null);

range.dataValidation({
  type: 'list',
  allowBlank: false,
  showInputMessage: false,
  prompt: false,
  promptTitle: 'String',
  showErrorMessage: false,
  error: 'String',
  errorTitle: 'String',
  operator: 'String',
  formula1: 'Item1,Item2,Item3,Item4',
  formula2: 'String'
});
range.dataValidation('Item1,Item2,Item3,Item4');
const obj = range.dataValidation();
range.dataValidation(null);

workbook
  .sheet(0)
    .cell("A1")
      .value("foo")
      .style("bold", true)
    .relativeCell(1, 0)
      .formula("A1")
      .style("italic", true)
.workbook()
  .sheet(1)
    .range("A1:B3")
      .value(5)
    .cell(0, 0)
      .style("underline", "double");

cell.value("Link Text")
  .style({ fontColor: "056c1", underline: true })
  .hyperlink("http://example.com");
cell.value("Link Text")
  .style({ fontColor: "0563c1", underline: true });
  .hyperlink({ hyperlink: "http://example.com", tooltip: "example.com" });
const value = cell.hyperlink();
cell.value("Click to Email Jeff Bezos")
  .hyperlink({ email: "jeff@amazon.com", emailSubject: "I know you're a busy man Jeff, but..."});
cell.value("Click to go to an internal cell")
  .hyperlink("Sheet2!A1");
cell.value("Click to go to an internal cell")
  .hyperlink(workbook.sheet(0).cell("A1"))

sheet.printOptions('headings', true);
const headings = sheet.printOptions('headings');
sheet.printOptions('verticalCentered', undefined);
const verticalCentered = sheet.printOptions('verticalCentered');
sheet.printGridLines(true);
sheet.printOptions('gridLines') == sheet.printOptions('gridLinesSet') === true;
sheet.printOptions('gridLineSets', false);
const isPrintGridLinesEnabled = sheet.printGridLines();

sheet.pageMargins('top', 1.1);
const topPageMarginInInches = sheet.pageMargins('top');

router.get("/download", function (req, res, next){
  XlsxPopulate.fromFileAsync("input.xlsx")
    .then(workbook => {
      workbook.sheet(0).cell("A1").value("foo");
      return workbook.outputAsync();
    })
    .then(data => {
      res.attachment("output.xlsx");
      res.send(data);
    })
    .catch(next);
});

var file = document.getElementById("file-input").files[0];
XlsxPopulate.fromDataAsync(file)
  .then(funciton(workbook){
  });
  
var req = new XMLHttpRequest();
req.open("GET", "http://...", true);
req.responseType = "arraybuffer";
req.onreadystatechange = function(){
  if(req.readyState === 4 && req.status === 200){
    Xlsx.Populate.fromDataAsync(req.response)
      .then(function(workboo){
      });
  }
};
req.send();

workbook.outputAsync()
  .then(funciton(blob){
    if (window.navigator && window.nabigator.msSaveOrOpenBlob){
      window.navigator.msSaveOrOpenBlob(blob, "out.xlsx");
    }else{
      var url = window.URL.createObjectURL(blob);
      var a = document.createElement("a");
      document.body.appendChild(a);
      a.href = url;
      a.download = "out.xlsx";
      a.click();
      window.URL.revokeObjectURL(url);
      document.body.removeChild(a);
  }
});

workbook.outputAsync("base64")
  .then(function(base64){
    location.href = "data:" + XlsxPopulate.MIME_TYPE + ";base64," + base64;
  });

var Promise = XlsxPopulate.Promise;

const Promise = require("bluebird");
const XlsxPopulate = require('xlsx-populate');
XlsxPopulate.Promise = Promise;

XlsxPopulate.fromFileAsync("./Book1.xlsx", { password: "S3cret!" })
  .then(workbook => {
  });

workbook.toFileAsync("./out.xlsx", { password: "S3cret!" });

```

```
```



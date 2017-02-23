var Excel = require('exceljs');
var filename = __dirname + '\\fitnes.xlsx';
var color = require('./color.json');
var fs = require('fs');

var colums_col = 27;
var fitnes = {};
// var count = 1;

// Access an individual columns by key, letter and 1-based column number
// var idCol = worksheet.getColumn('id');
// var nameCol = worksheet.getColumn('B');
// var dobCol = worksheet.getColumn(3);

// columns

// iterate over all current cells in this column
// dobCol.eachCell(function(cell, rowNumber) {
//     // ...
// });

// Iterate over all non-null cells in a row
// row.eachCell(function(cell, colNumber) {
//     console.log('Cell ' + colNumber + ' = ' + cell.value);
// });

// iterate over all current cells in this column including empty cells
// dobCol.eachCell({ includeEmpty: true }, function(cell, rowNumber) {
//     // ...
// });

// console.log(filename);
// read from a file
var workbook = new Excel.Workbook();
workbook.xlsx.readFile(filename)
    .then(function() {
      console.log('read success');
      var worksheet = workbook.getWorksheet(1);
      get_data(worksheet);
        // //Определим сколько столбцов
        // for (var i = 2; i < 35; i++) {
        //   if (worksheet.getRow(3).getCell(i).value > 0) {
        //     colums_col = worksheet.getRow(3).getCell(i).value;
        //   }
        // }
        // console.log(colums_col);
    });

//Функция обозначение цветов
  function rename_color(color){
    // console.log(color);
      if (color == 'FF92D050') {
          return('green');
      }
      else if (color == 'FFFF0000') {
        return('red');
      }
      else if (color == 'FF0070C0') {
        return('blue');
      }
      else if (color == 'FFFFFF00') {
        return('yellow');
      }
      else if (color == 'FFFFC000') {
        return('orange');
      }
      // else if (color == undefined) {
      //   return('black');
      // }
      else {
        return('white');
      }
  }
//

//Функция перебора всей таблицы
function get_data(sheet) {
  // for (var col = 2; col < colums_col + 2; col++) {
  for (var col = 26; col < 27; col++) {
    var col_1 = sheet.getColumn(col);
    var fi = {};
    var count = 1;
    var ending = 1;
    var cell_before_color = 'white';
    col_1.eachCell({ includeEmpty: false }, function(cell, rowNumber) {
      if (rowNumber > 3 && rowNumber < 65) {
        var adr = cell.address;
        // console.log(adr);
        var cell_color = rename_color(cell.fill.fgColor.argb);
        // console.log(cell_color);
        if (fi[count - 1]) {
          if (cell_color != fi[count - 1].color && ending == 1) {
            var row_before = rowNumber - 1;
            var time_val_end = sheet.getCell('A' + row_before).value;
            // console.log(time_val_end);
            ending = 2;
            var p = sheet.getCell('A' + col).value;
            fi[count - 1].end = time_val_end.charAt(6) + time_val_end.charAt(7) + ':' + time_val_end.charAt(9) + time_val_end.charAt(10);
            // console.log(fi);
            // fitnes[col-1] = fi;
          }
        }

        if (cell_color != 'white' && cell_before_color != cell_color) {
          var time_val = sheet.getCell('A' + rowNumber).value;
          // console.log(time_val);
          fi[count] = {};
          fi[count].color = cell_color;
          fi[count].begin = time_val.charAt(0) + time_val.charAt(1) + ':' + time_val.charAt(3) + time_val.charAt(4);
          // console.log(fi);
          count += 1;
          ending = 1;
        }
        cell_before_color = cell_color;
      }
      // console.log(fi);
    });
    console.log(col);
    // console.log(fi);
    fitnes[col-1] = fi;

  }
  console.log('\nfitnes = ');
  console.log(fitnes);
  // console.log(fitnes);
}


  // function get_before(sh, ad, ce) {
  //   ad = 'AB14';
  //   console.log(ad);
  //   console.log('in get_before');
  //   console.log(ad.length);
  //   if (ad.length > 2) {
  //     console.log('in ad = ' + ad.length);
  //     var nomer_before = parseInt(ad.substring(2)) - 1;
  //     var adres_before = ad.charAt(0) + ad.charAt(1) + nomer_before;
  //     var cell_before = sh.getCell(adres_before).fill.fgColor.argb;
  //     var color_before = rename_color(cell_before);
  //     console.log(color_before);
  //   }
  //
  //   if (ad.length < 2) {
  //
  //   }
  // }

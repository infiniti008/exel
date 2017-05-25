var Excel = require('exceljs');
var filename = 'tren';
var color = require('./color.json');
var fs = require('fs');

var month = 3;//Номер месяца
var posledn_stroka = 33;//Номер последней строки
var row_date = 3; //Номер строки с датами в месяце
var first_col = 2;//Номер первого столбца с данными

// var colums_col = 28;
var fitnes = {};
var workbook = new Excel.Workbook();
workbook.xlsx.readFile(__dirname + '\\exel_file_' + month + '\\' + filename + '.xlsx')
    .then(function() {
        console.log('read success');
        var worksheet = workbook.getWorksheet(1);
        get_arr(worksheet)
          .then(
            data => {
              // console.log(data);
              get_data_new(worksheet, data);
            }
          );
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
function get_data_new(sheet, da) {
  var color_current;
  var color_before;
  var color_after;
  var time_before;
  var ti = da.time;
  var table = da.table;
  for (var col = 0; col < table.length; col++) {
    var fi = {};
    var count = 1;
    var ending = 1;
    // console.log(table[col].length);
    for (var i = 0; i < table[col].length; i++) {
      // color_current = rename_color(sheet.getCell(table[col][i]).fill.fgColor.argb);
      color_current = table[col][i];
      // console.log(table[col][i] + ' - current -  ' + color_current);
      var j = i;
      if (j == 0) {
        j = 1;
      }
      // color_before = rename_color(sheet.getCell(table[col][j-1]).fill.fgColor.argb);
      color_before = table[col][j-1];
      // console.log(table[col][i] + ' - before - ' + color_before);
      var u = i;
      if (u == 60) {
        u = 59;
      }
      // color_after = rename_color(sheet.getCell(table[col][u+1]).fill.fgColor.argb);
      color_after = table[col][u+1];
      // console.log(table[col][i] + ' - after - ' + color_after);
      if (fi[count - 1]) {
        if (color_current != fi[count - 1].color && ending == 1) {
           //console.log(time_val_end);
           var o = i;
           if (o == 0) {
             o = 1;
           }
          time_before = ti[o - 1];
          ending = 2;
          fi[count - 1].end = time_before.charAt(6) + time_before.charAt(7) + ':' + time_before.charAt(9) + time_before.charAt(10);
          // console.log(fi);
        }
      }
      if (color_current != 'white' && color_current != color_before) {
        var time_val = ti[i];
        // console.log(time_val);
        fi[count] = {};
        fi[count].color = color_current;
        fi[count].begin = time_val.charAt(0) + time_val.charAt(1) + ':' + time_val.charAt(3) + time_val.charAt(4);
        // console.log(fi);
        count += 1;
        ending = 1;
      }
    }
    fitnes[col + 1] = fi;
  }
  console.log(fitnes);

  var path_save = __dirname + '\\grafic_' + month + '\\';
  if (!fs.existsSync(path_save)){
      fs.mkdirSync(path_save);
  }

  var to_json = JSON.stringify(fitnes);

  console.log(path_save);
  fs.writeFile(path_save + filename  + '.json', to_json, function () {
    console.log('write success');
  });
}
//
function get_arr(shee){
  console.log('in get_arr');
  var first_row = row_date + 1;
  var arr = [];//Массив имен столбцев
  var table = [];// Массив таблицы с данными
  var dat = {};
  var row = shee.getRow(row_date);
  row.eachCell(function(cell, colNumber) {
    var ind = cell.address.match(row_date).index;
    var im = cell.address.substring(0, ind);
    arr[colNumber-first_col] = im;
  });
  var day = arr.length; //Количество дней в месяце
  // day = 28;
  for (var i = 0; i < day; i++) {
    var row = [];
    for (var j = first_row; j < posledn_stroka + 1; j++) {
      var ty = arr[i] + j;
      if (shee.getCell(ty).fill) {
        row[j - first_row] = rename_color(shee.getCell(ty).fill.fgColor.argb);
      }
      else {
        row[j - first_row] = 'white';
      }
      // row[j - first_row] = arr[i] + j;
    }
    table[i] = row;
  }
  var time = []; //Массив времени
  for (var o = first_row; o < first_row + 61; o++) {
    time[o - first_row] = shee.getCell('A' + o).value;
  }
  // console.log(table);
  dat.time = time;
  dat.table = table;
    return new Promise((resolve, reject) => {
      setTimeout(function () {
        resolve(dat);
      }, 1000);
    });
}

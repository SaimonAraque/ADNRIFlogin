'use strict';

var global_workbook = null;
var global_worksheet = null;
var rABS = true;

function getWorksheet(workbook, position) {
  return workbook.Sheets[workbook.SheetNames[position]];
}

function getCell(cell, worksheet) {
  return global_worksheet[cell] ? global_worksheet[cell].v : undefined;
}

function getColumn(column, worksheet) {
  var data = {};
  var values = Object.keys(worksheet).filter(function (key) {
    return key.slice(0, 1) == column && (worksheet[key].v[0] == 'V' || worksheet[key].v[0] == 'J' || worksheet[key].v[0] == 'G' || worksheet[key].v[0] == 'E' || worksheet[key].v[0] == 'C' || worksheet[key].v[0] == 'P') && Number.isInteger(parseInt(worksheet[key].v[4]));
  }).map(function (key) {
    return worksheet[key].v;
  });

  for (var i = 0; i < values.length; i++) {
    data[i] = values[i];
  }return JSON.stringify(data);
}

function loadExcelFile(e) {
  var files = e.target.files,
      f = files[0];
  var reader = new FileReader();

  reader.onload = function (e) {
    var data = e.target.result;
    if (!rABS) data = new Uint8Array(data);
    global_workbook = XLSX.read(data, { type: rABS ? 'binary' : 'array' });

    global_worksheet = getWorksheet(global_workbook, 0);
  };

  if (rABS) reader.readAsBinaryString(f);else reader.readAsArrayBuffer(f);
}
"use strict";

var column_file = document.querySelector('.select_excel_column');

var query_column = 'A';
column_file.addEventListener('change', function (e) {
    query_column = e.target.value;
});

var search = document.querySelector('#searchButton');

search.addEventListener('click', function () {
    if (document.querySelector('.form_input_text').value != "") document.forms["rifForm"].submit();
});

var upload = document.querySelector('.upload_excel_button');

upload.addEventListener('change', loadExcelFile, false);

document.querySelector('#sendButton').addEventListener('click', function () {
    document.querySelector("#hiddenInput").value = getColumn(query_column, global_worksheet);
    document.forms["sendForm"].submit();
});

var closeButton = document.querySelector('.upload_file_box__close');

closeButton.addEventListener('click', function () {
    document.querySelector('.upload_file_box').classList.add('upload_file_box__closed');
});

var openButton = document.querySelector('.upload_file_box__open');

openButton.addEventListener('click', function () {
    document.querySelector('.upload_file_box').classList.toggle('upload_file_box__closed');
});
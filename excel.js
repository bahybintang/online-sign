var xls = document.getElementById("xls");
var konsep = document.getElementById("konsep");
var signButton = document.getElementById("sign");
var workbook;

xls.addEventListener("change", async function (e) {
  workbook = new ExcelJS.Workbook();
  var reader = new FileReader();
  reader.readAsArrayBuffer(e.target.files[0]);
  reader.onloadend = async function () {
    await workbook.xlsx.load(reader.result);
    if (konsep.value != "0") signButton.classList.remove("disabled");
    else signButton.classList.add("disabled");
  };
});

konsep.addEventListener("change", function (e) {
  if (e.target.value != "0" && xls.files.length > 0)
    signButton.classList.remove("disabled");
  else signButton.classList.add("disabled");
});

signButton.addEventListener("click", async function (e) {
  var imageData = signaturePad.toDataURL("image/jpeg");
  const imageId = workbook.addImage({
    base64: imageData,
    extension: "jpeg",
  });
  var worksheet = workbook.getWorksheet("Sheet1");
  var pos = getInsertPosition(worksheet);
  addSign(worksheet, imageId, pos);
  const buffer = await workbook.xlsx.writeBuffer();
  var blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  var link = document.createElement("a");
  link.href = window.URL.createObjectURL(blob);
  link.download = "Result.xlsx";
  link.click();
});

function getInsertPosition(worksheet) {
  var row = 20 + parseInt(konsep.value),
    column = "A".charCodeAt(0);
  var currRow = worksheet.getRow(row);
  for (var i = 3; ; i++) {
    if (!currRow.getCell(i).value) {
      column += i;
      break;
    }
  }
  column = String.fromCharCode(column - 1);
  return `${column}${row}:${column}${row}`;
}

function addSign(worksheet, imageId, pos) {
  var cellNum = pos.split(":")[0];
  var rowNum = pos.split(/[a-zA-Z:]+/)[1] - 1;
  var colName = pos.split(/\d+/)[0];
  var colNum = colName.charCodeAt(0) - "A".charCodeAt(0);
  var cellWidth = worksheet.getColumn(colName).width;
  var imageSize = getImageSizeInCell(400, 200, cellWidth);
  worksheet.addImage(imageId, {
    tl: { col: colNum + 0.1, row: rowNum + 0.1 },
    ext: imageSize,
    editAs: "oneCell",
  });
  worksheet.getRow(rowNum + 1).height = 50;
  var date = new Date();
  worksheet.getCell(cellNum).value = new Date(
    Date.UTC(
      date.getFullYear(),
      date.getMonth(),
      date.getDate(),
      date.getHours(),
      date.getMinutes(),
      date.getSeconds()
    )
  );
}

function getImageSizeInCell(width, height, cellWidth) {
  height = height * (cellWidth / width) * 6.5;
  width = cellWidth * 6.5;
  return { width: parseInt(width), height: parseInt(height) };
}

function getDateIndonesia(date) {
  date.setHours(date.getHours() + 7);
}

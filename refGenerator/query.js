//create file system object
var fso = new ActiveXObject("Scripting.FileSystemObject");
//get inputs
var year = document.getElementById("year");
var month = document.getElementById("month");
//get button
var button = document.getElementById("button");
//get output
var outputid = document.getElementById("output");
var yearid, monthid, xelObj, thisyear, thismonth, fso;
var localpath = fso.GetAbsolutePathName(".");
var thisyear = new Date().getFullYear();
var thismonth = new Date().getMonth() + 1;
year.value = thisyear;
month.value = thismonth;
function submit() {
  //get current year and month
  //create xml object
  xelObj = new ActiveXObject("Excel.Application");
  //get input value
  yearid = year.value;
  monthid = month.value;
  //check if input is valid
  if (yearid > thisyear || yearid < 2000) {
    alert("Invalid year");
  } else if (monthid > 12 || monthid < 1) {
    alert("Invalid month");
  } else {
    //get file
    path = localpath;
    if (fso.fileexists(localpath + "/appdata.txt")) {
      var file = fso.opentextfile(localpath + "/appdata.txt", 1);
      var pathdir = file.readline();
      file.close();
      if (fso.folderexists(pathdir)) {
        path = pathdir;
      }
    }
    if (!fso.folderexists(localpath + "/data")) {
      fso.createFolder(localpath + "/data");
    }
    if (fso.folderexists(localpath + "/data/query")) {
      fso.createFolder(localpath + "/data/query");
    }
    var dfilename = "data-" + yearid + "-" + monthid + ".txt";
    var data = [];
    if (fso.fileexists(path + "/data/history/" + dfilename)) {
      var file = fso.opentextfile(path + "/data/history/" + dfilename, 1);
      while (!file.atendofstream) {
        var linedata = file.readline();
        data.push(linedata.split("|"));
      }
      file.close();
      excel = xelObj.Workbooks.Add();
      xelObj.Visible = true;
      createExl(excel, data);
    } else {
      alert("No data file found for this month");
    }
  }
  //get year and month
}

function createExl(excel, data) {
  var wrks = excel.Worksheets(1);
  wrks.activate;
  wrks.visible = true;
  wrks.columns(1).columnwidth = 5;
  wrks.columns(2).columnwidth = 20;
  wrks.columns(3).columnwidth = 30;
  wrks.columns(4).columnwidth = 10;
  wrks.columns(5).columnwidth = 10;
  wrks.columns(6).columnwidth = 10;
  wrks.columns(6).numberformat = "[$-en-US]h:mm:ss AM/PM";
  wrks.name = "query";
  wrks.cells(1, 1).value = "S. no";
  wrks.cells(1, 2).value = "Name";
  wrks.cells(1, 3).value = "Email";
  wrks.cells(1, 4).value = "Reference";
  wrks.cells(1, 5).value = "Date";
  wrks.cells(1, 6).value = "Time";
  for (var i = 1; i <= 6; i++) {
    wrks.cells(1, i).interior.colorindex = 15;
    wrks.cells(1, i).font.bold = true;
    wrks.cells(1, i).font.size = 12;
    wrks.cells(1, i).font.name = "Arial";
  }
  for (var i = 2; i <= data.length + 1; i++) {
    wrks.cells(i, 1).value = i - 1;
    wrks.cells(i, 2).value = data[i - 2][0];
    wrks.cells(i, 3).value = data[i - 2][1];
    wrks.cells(i, 4).value = data[i - 2][2];
    wrks.cells(i, 5).value = data[i - 2][3];
    wrks.cells(i, 6).value = data[i - 2][4];
  }
}

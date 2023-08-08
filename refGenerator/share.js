var fso = new ActiveXObject("Scripting.FileSystemObject"),
  ref,
  refPath,
  path = "refGenerator",
  refPrimary = "KBS",
  refSecondary = 0,
  localpath = fso.GetAbsolutePathName(".").replace(/\\/g, "/"),
  Name = document.getElementById("name"),
  Email = document.getElementById("email"),
  date = new Date(),
  lrefPath,
  DomString = "",
  dom = document.getElementById("log");

if (document.getElementById("result").innerHTML == "") {
  document.getElementById("copy").style.display = "none";
  document.getElementById("generate").style.display = "block";
} else {
  document.getElementById("copy").style.display = "block";
  document.getElementById("generate").style.display = "none";
}
var connection = document.getElementById("connection");
var connect = document.getElementById("connect");
if (!fso.FileExists(localpath + "/appdata.txt")) {
  path = prompt("Enter your folder name", "F:/refGenerator");
  path = path
    .replace(/\\/g, "/")
    .replace(/\/$/, "")
    .replace(/file:\/\//, "");
  var appdata = fso.CreateTextFile(localpath + "/appdata.txt");
  appdata.WriteLine(path);
  appdata.Close();
}
var appdata = fso.OpenTextFile(localpath + "/appdata.txt", 1, true);
path = appdata.ReadLine();
appdata.Close();
var oldpath = checkConnection("return");
clearLog();
function myfunction() {
  //check if folder exists
  var currentpath = checkConnection("return");
  var inp;
  //check if file exists
  if (fso.FolderExists(path)) {
    if (
      !fso.FileExists(path + "/ref.txt") &&
      !fso.FileExists(localpath + "/ref.txt")
    ) {
      inp = prompt("Enter last reference number", "KBS1000");
      refSecondary = parseInt(inp.replace(refPrimary, ""));
      refSecondary++;
      refPath = fso.CreateTextFile(path + "/ref.txt");
      refPath.Write(refSecondary);
      refPath.Close();
      lrefPath = fso.CreateTextFile(localpath + "/ref.txt");
      lrefPath.Write(refSecondary);
      lrefPath.Close();
    } else if (
      fso.FileExists(path + "/ref.txt") &&
      !fso.FileExists(localpath + "/ref.txt")
    ) {
      refPath = fso.OpenTextFile(path + "/ref.txt", 1, true);
      ref = refPath.ReadAll();
      refPath.Close();
      refSecondary = parseInt(ref);
      refSecondary++;
      lrefPath = fso.CreateTextFile(localpath + "/ref.txt");
      lrefPath.Write(refSecondary);
      lrefPath.Close();
      refPath = fso.OpenTextFile(path + "/ref.txt", 2, true);
      refPath.Write(refSecondary);
      refPath.Close();
    } else if (
      !fso.FileExists(path + "/ref.txt") &&
      fso.FileExists(localpath + "/ref.txt")
    ) {
      lrefPath = fso.OpenTextFile(localpath + "/ref.txt", 1, true);
      lref = lrefPath.ReadAll();
      lrefPath.Close();
      refSecondary = parseInt(lref);
      refSecondary++;
      refPath = fso.CreateTextFile(path + "/ref.txt");
      refPath.Write(refSecondary);
      refPath.Close();
      localpath = fso.OpenTextFile(localpath + "/ref.txt", 2, true);
      localpath.Write(refSecondary);
      localpath.Close();
    } else {
      refPath = fso.OpenTextFile(path + "/ref.txt", 1, true);
      ref = parseInt(refPath.ReadAll());
      refPath.Close();
      lrefPath = fso.OpenTextFile(localpath + "/ref.txt", 1, true);
      var lref = parseInt(lrefPath.ReadAll());
      lrefPath.Close();
      if (ref > lref) {
        refSecondary = ref;
      } else {
        refSecondary = lref;
      }
      refSecondary++;
      refPath = fso.OpenTextFile(path + "/ref.txt", 2, true);
      refPath.Write(refSecondary);
      refPath.Close();
      lrefPath = fso.OpenTextFile(localpath + "/ref.txt", 2, true);
      lrefPath.Write(refSecondary);
      lrefPath.Close();
    }
  } else {
    //save in local
    if (!fso.fileExists(localpath + "/ref.txt")) {
      inp = prompt("Enter last reference number", "KBS1000");
      refSecondary = parseInt(inp.replace(refPrimary, ""));
      lrefPath = fso.CreateTextFile(localpath + "/ref.txt");
      refSecondary++;
      lrefPath.Write(refSecondary);
      lrefPath.Close();
    } else {
      lrefPath = fso.OpenTextFile(localpath + "/ref.txt", 1, true);
      lref = lrefPath.ReadAll();
      lrefPath.Close();
      refSecondary = parseInt(lref);
      refSecondary++;
      lrefPath = fso.OpenTextFile(localpath + "/ref.txt", 2, true);
      lrefPath.Write(refSecondary);
      lrefPath.Close();
    }
  }
  //check if name is empty
  if (trim(Name.value) == "") Name.value = "Anonymous";
  //check if email is empty
  if (trim(Email.value) == "") Email.value = "Unknown";
  //create txt file
  var month = date.getMonth() + 1;
  var year = date.getFullYear();
  var tdate = date.getDate() + "-" + month + "-" + year;
  var time =
    date.getHours() + ":" + date.getMinutes() + ":" + date.getSeconds();
  var historyDt =
    Name.value +
    "|" +
    Email.value +
    "|" +
    refPrimary +
    refSecondary +
    "|" +
    tdate +
    "|" +
    time;
  if (!fso.FolderExists(currentpath + "/data")) {
    fso.CreateFolder(currentpath + "/data");
  }
  if (!fso.FolderExists(currentpath + "/data/history")) {
    fso.CreateFolder(currentpath + "/data/history");
  }
  if (currentpath == path) {
    saveInLog(historyDt, currentpath);
  }
  if (
    !fso.FileExists(
      currentpath + "/data/history/" + "data-" + year + "-" + month + ".txt"
    )
  ) {
    var historyFile = fso.CreateTextFile(
      currentpath + "/data/history/" + "data-" + year + "-" + month + ".txt"
    );
    historyFile.WriteLine(historyDt);
    historyFile.Close();
  } else {
    var historyFile = fso.OpenTextFile(
      currentpath + "/data/history/" + "data-" + year + "-" + month + ".txt",
      8,
      true
    );
    historyFile.WriteLine(historyDt);
    historyFile.Close();
  }
  document.getElementById("result").innerHTML = refPrimary + refSecondary;
  document.getElementById("copy").style.display = "block";
  document.getElementById("generate").style.display = "none";
  document.getElementById("name").value = "";
  document.getElementById("email").value = "";
}
function reset() {
  var dissition = confirm("Are you sure?");
  if (dissition) {
    var inp = prompt(
      "Enter your last reference number. It's for the first time only.",
      "KBS10000"
    );
    refSecondary = parseInt(inp.replace(refPrimary, ""));
    lrefPath = fso.OpenTextFile(localpath + "/ref.txt", 2, true);
    lrefPath.Write(refSecondary);
    lrefPath.Close();
    if (fso.FolderExists(path)) {
      refPath = fso.OpenTextFile(path + "/ref.txt", 2, true);
      refPath.Write(refSecondary);
      refPath.Close();
    }
    document.getElementById("result").innerHTML = "";
  }
}
function trim(str) {
  if (str) {
    return str.replace(/^\s+|\s+$/gm, "");
  } else {
    return "";
  }
}
function checkConnection(ifr) {
  var txt = "",
    color = "";
  if (fso.FolderExists(path)) {
    returnVal = path;
    appdataPATH = path + "/data";
    if (!fso.FolderExists(appdataPATH)) {
      fso.CreateFolder(appdataPATH);
    }
    if (!fso.FolderExists(appdataPATH + "/history")) {
      fso.CreateFolder(appdataPATH + "/history");
    }
    if (fso.FolderExists(localpath + "/data/history")) {
      var folder = fso.GetFolder(localpath + "/data/history");
      var files = new Enumerator(folder.Files);
      while (!files.atEnd()) {
        var file = files.item();
        //move file to new folder
        if (!fso.fileExists(path + "/data/history/" + file.Name)) {
          file.Move(path + "/data/history/" + file.Name);
        } else {
          var file2 = fso.OpenTextFile(
            path + "/data/history/" + file.Name,
            8,
            true
          );
          var file1 = fso.OpenTextFile(file, 1, true);
          while (!file1.AtEndOfStream) {
            file2.WriteLine(file1.ReadLine());
          }
          file1.Close();
          file2.Close();
          fso.DeleteFile(file);
        }
        files.moveNext();
      }
    } else {
      if (!fso.FolderExists(localpath + "/data")) {
        fso.CreateFolder(localpath + "/data");
      }
      fso.CreateFolder(localpath + "/data/history");
    }
    txt = "It's Connected to " + path + " successfully";
    color = "0a0";
  } else if (fso.FolderExists(localpath)) {
    returnVal = localpath;
    appdataPATH = localpath + "/data";
    if (!fso.FolderExists(appdataPATH)) {
      fso.CreateFolder(appdataPATH);
    }
    if (!fso.FolderExists(appdataPATH + "/history")) {
      fso.CreateFolder(appdataPATH + "/history");
    }
    txt =
      "Can't connect to " +
      path +
      ". It's connected to " +
      localpath +
      " instead";
    color = "fa0";
  } else {
    txt = "Can't connect to any folder. Something is wrong!";
    color = "f00";
  }
  connection.style.color = color;
  connection.innerHTML = txt;
  updateInlog();
  if (ifr) return returnVal;
}
function runQ() {
  var wsh = new ActiveXObject("WScript.Shell");
  wsh.run("query.hta", 1, true);
}
function saveInLog(data, path) {
  var logFile;
  if (fso.FileExists(path + "/data/log.txt")) {
    logFile = fso.OpenTextFile(path + "/data/log.txt", 8, true);
    logFile.WriteLine(data);
    logFile.Close();
  } else {
    logFile = fso.CreateTextFile(path + "/data/log.txt");
    logFile.WriteLine(data);
    logFile.Close();
  }
}
function deleteData() {
  var dissition = confirm("Are you sure you want to delete all data?");
  var datapath = checkConnection("return");
  if (dissition) {
    fso.DeleteFolder(datapath + "/data", true);
    fso.DeleteFolder(localpath + "/data", true);
  }
}
function clearChache() {
  var chech = checkConnection("return");
  fso.DeleteFile(chech + "/data/log.txt");
}
function updateInlog() {
  DomString =
    '<div class="th"><span class="td no0">No.</span><span class="td no1">Name</span><span class="td no2">Email</span><span class="td no3">Referance</span><span class="td no4">Date</span><span class="td no5">Time</span></div>';
  if (fso.FolderExists(path)) {
    if (!fso.FolderExists(path + "/data")) fso.CreateFolder(path + "/data");
    if (!fso.FileExists(path + "/data/log.txt"))
      fso.CreateTextFile(path + "/data/log.txt");
    if (fso.FolderExists(path + "/data")) {
      var logFile = fso.OpenTextFile(path + "/data/log.txt", 1, true);
      var i = 0;
      var lines = [];
      while (!logFile.AtEndOfStream) {
        lines[i] = logFile.ReadLine();
        i++;
      }
      logFile.Close();
      var logFileData = [];
      for (var j = lines.length - 1; j > lines.length - 10 && j >= 0; j--) {
        DomString += '<div class="tr">';
        logFileData = lines[j].split("|");
        DomString += '<span class="td no0">' + (j + 1) + "</span>";
        for (var k = 0; k < logFileData.length; k++) {
          DomString +=
            '<span class="td no' + (k + 1) + '">' + logFileData[k] + "</span>";
        }
        DomString += "</div>";
      }
      dom.innerHTML = DomString;
    }
  } else {
    dom.innerHTML = "Not Connected with yet.";
  }
}
function clearLog() {
  if (
    fso.FolderExists(path + "/data") &&
    fso.FileExists(path + "/data/log.txt")
  ) {
    var f = fso.OpenTextFile(path + "/data/log.txt", 1, true);
    //keep last 15 lines
    var lines = f.ReadAll().split("\n");
    var newLines = lines.slice(lines.length - 15);
    f = fso.OpenTextFile(path + "/data/log.txt", 2, true);
    f.Write(newLines.join("\n"));
    f.Close();
  }
}
setInterval(updateInlog, 5000);

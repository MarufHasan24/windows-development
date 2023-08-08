window.resizeTo 1200, 700
Function refresh()
window.location.reload
End Function
Function Button()
    style = document.getElementById("buttonCon").className
    If style = "none" Then
    document.getElementById("buttonCon").className = "block"
    Else
    document.getElementById("buttonCon").className = "none"
    End If
End Function
    Function clipboard()
      Dim oClip
      Set oClip = CreateObject("htmlfile")
    oClip.write "<html><body>" & document.getElementById("result").innerHTML & "</body></html>"
    oClip.body.createTextRange().execCommand "Copy"
      Set oClip = Nothing
    document.getElementById("result").innerHTML = ""
    document.getElementById("copy").style.display = "none"
    document.getElementById("generate").style.display = "block"
    document.getElementById("name").value = ""
    document.getElementById("email").value = ""
    End Function
    
    Dim objFSO, sv
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    sv = objFSO.GetAbsolutePathName(".")
    Set objShell = CreateObject("WScript.Shell")
    If NOT(objFSO.FileExists(objShell.SpecialFolders("StartMenu") & "\Reference Generator.lnk")) Then
    CreateAShortcut
    End If

    Function CreateAShortcut()
        Dim objShortcut
        Set objShortcut = objShell.CreateShortcut(objShell.SpecialFolders("StartMenu") & "\Reference Generator.lnk")
    objShortcut.TargetPath = sv & "\index.hta"
    objShortcut.WorkingDirectory = sv
    objShortcut.Description = "Reference Generator"
    objShortcut.IconLocation = sv & "\icon.ico"
    objShortcut.Save
        Set objShortcut = Nothing
        Set objShell = Nothing
        Set objFSO = Nothing
    End Function
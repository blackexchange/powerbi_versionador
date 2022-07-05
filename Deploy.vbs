Function FileExists(FilePath)
  Set fso = CreateObject("Scripting.FileSystemObject")
  If fso.FileExists(FilePath) Then
    FileExists=CBool(1)
  Else
    FileExists=CBool(0)
  End If
End Function

set oShell = CreateObject("WScript.shell")


If FileExists("D:\dev\PowerBI\pbix\Capag.pbix") Then
  command= "cmd /c python D:\dev\py\GeradorPB\pbivcs.py -x --over-write D:\dev\PowerBi\pbix\Capag.pbix D:\dev\PowerBI\DevOps\pbix"
  command2= "cmd /c python D:\dev\py\GeradorPB\pbivcs.py -x --over-write D:\dev\PowerBi\pbit\Capag.pbit D:\dev\PowerBI\DevOps\pbit"
    
 
  oShell.run "cmd.exe /c " & command
  oShell.run "cmd.exe /c " & command2

Else
  command=  "cmd /c python D:\dev\py\GeradorPB\pbivcs.py -c --over-write D:\dev\PowerBI\DevOps\pbix D:\dev\PowerBi\pbix\Capag.pbix"
  command2= "cmd /c python D:\dev\py\GeradorPB\pbivcs.py -c --over-write D:\dev\PowerBI\DevOps\pbit D:\dev\PowerBi\pbit\Capag.pbit"
  
  oShell.run "cmd.exe /c " & command
  oShell.run "cmd.exe /c " & command2

  
End If


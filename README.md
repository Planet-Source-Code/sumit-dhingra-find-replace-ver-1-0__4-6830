<div align="center">

## Find & Replace \(ver 1\.0\)


</div>

### Description

This will find all occurences of a string in a text file and replace it with another string.
 
### More Info
 
Filename

String to search

String to replace

Comparison Method (optional)

Usage :

Save the code as a .vbs file.

findreplace.vbs mylog.txt Football Baseball 0

Note : CompareMethod can be either 0 (Binary comparison) or 1 (Text comparison).


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Sumit Dhingra](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/sumit-dhingra.md)
**Level**          |Intermediate
**User Rating**    |5.0 (20 globes from 4 users)
**Compatibility**  |VbScript \(browser/client side\)

**Category**       |[Strings](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/strings__4-26.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/sumit-dhingra-find-replace-ver-1-0__4-6830/archive/master.zip)





### Source Code

```
Option Explicit
Dim fso
Dim folder
Dim logfile
Dim logLine
Dim newfile
Dim deletefile
Dim filecontents
Dim finalcontents
Dim objArgs
Dim filename
Dim folderpath
Dim searchstr
Dim replacestr
Dim CompareMethod
Set objArgs = WScript.Arguments
Set fso = CreateObject("Scripting.FileSystemObject")
If objArgs.count >= 3 and objArgs.count <=4 then
	filename 	= objArgs(0)
	searchstr 	= objArgs(1)
	replacestr 	= objArgs(2)
	If objArgs.count = 4 then
		CompareMethod = objArgs(3)
		If CompareMethod <> 0 and CompareMethod <> 1 Then
			Wscript.Echo "CompareMethod can only be 0 or 1"
			Wscript.Quit(1)		'To indicate error.
		End If
	Else
		CompareMethod = 0	' Default to vbBinaryCompare.
	End If
Else
	wscript.echo "Usage: FindReplace.vbs [arguments..]" + vbcrlf + vbcrlf + "Arguments:" + vbcrlf + "File to be Searched" 		+ vbcrlf + "Searched string" + vbcrlf + "String to replace" + vbcrlf + "[Comparison Method]"
	 wscript.Quit (1)	'To indicate error.
End if
TextSearch(Filename)
Function TextSearch(Filename)
	Set logfile = fso.OpenTextFile(filename)
	If Err.Number <> 0 Then
		Wscript.echo Err.Description
		Wscript.Quit (Err.Number)
	End If
	filecontents = logfile.readall
	If CompareMethod = 0 Then
	 	finalcontents = Replace(filecontents, searchstr, replacestr, 1, -1, vbBinaryCompare)
	Else
		finalcontents = Replace(filecontents, searchstr, replacestr, 1, -1, vbTextCompare)
	End If
	logfile.Close
	Set deletefile = fso.getFile(filename)
	deletefile.delete
	set newfile = fso.CreateTextFile(filename, true)
	newfile.write FinalContents
	newfile.close
	If Err.Number <> 0 Then
		Wscript.echo Err.Description
		Wscript.Quit (Err.Number)
	End If
	set logfile = nothing
	set deletefile	= nothing
	set newfile = nothing
End Function
```


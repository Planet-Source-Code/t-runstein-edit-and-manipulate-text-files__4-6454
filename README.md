<div align="center">

## Edit and manipulate text files


</div>

### Description

This is a drawn out example of reading and writing with the FileScriptingObject. This is similar to copying a file, but allows rewriting specific line(s). It's intentionally overdone so that you can delete what you don't want. Includes extensive error handling. I've included lots of comments for newbies.

'T Runstein
 
### More Info
 
If using in VB, include reference to Microsoft Scripting Runtime

Make sure you change the file names (strpath and strFldr) or create a C:\FirstFile.txt before running the script.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[T Runstein](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/t-runstein.md)
**Level**          |Intermediate
**User Rating**    |4.3 (43 globes from 10 users)
**Compatibility**  |ASP \(Active Server Pages\), VbScript \(browser/client side\)

**Category**       |[Files](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files__4-2.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/t-runstein-edit-and-manipulate-text-files__4-6454/archive/master.zip)





### Source Code

```
option explicit
on error resume next
'Since this was written for Windows Scripting Host,
'it uses VBScript which doesn't use types.
'To use this with VB, as types to the declarations
dim objFSO 	'as FileSystemObject
dim fle1	'as file
dim fle2	'as file
dim strPath	'as string
dim strFldr	'as string
dim strLine	'as string
strPath = "C:FirstFile.txt" 	'Put in the file you want to edit
strFldr = "C:TempFile.txt"
Main 'This Calls the Main sub
sub Main()
 dim rtn 'as integer
	rtn = CopyStuff() 'This calls and runs the CopyStuff function
 if rtn = 1 then
	msgbox "Copy is complete"
 else
	msgbox "An error was found and the process was aborted. " & Cstr(rtn)
		'The & Cstr(rtn) will display the number returned by CopyStuff
		'After you've got your script running, you may want to remove this feature
 end if
'Cleanup
 if not fle1 is nothing then set fle1 = nothing
 if not fle2 is nothing then set fle2 = nothing
 if not objFSO is nothing then set objFSO = nothing
end sub
function CopyStuff()
 set objFSO = CreateObject("Scripting.FileSystemObject") 'This creates the FSO
	'I've included error handling after each step
	if err.number <> 0 then
		msgbox "Error in Creating Object: " & err.number & "; " & err.description
		CopyStuff = 0 'Returns this number
		exit function 'Stop processing, go back to Main
	end if
 if not objFSO.FileExists(strPath) then 'The file to copy is not present
	msgbox "The " & strPath & " file was not found on this computer"
	CopyStuff = 2
	exit function
 end if
 if objFSO.FileExists(strFldr) then
	objFSO.DeleteFile(strFldr) 'If the temp file is found, delete it
 end if
	set fle1 = objFSO.OpenTextFile(strPath) 'Open
		if err.number <> 0 then
			msgbox "Error opening " & strPath & ": " & err.number & "; " & err.description
			CopyStuff = 3
			exit function
		end if
	set fle2 = objFSO.CreateTextFile(strFldr) 'Create the temp file
		if err.number <> 0 then
			msgbox "Error creating temp ini: " & err.number & "; " & err.description
			CopyStuff = 4
			exit function
		end if
	'Here's the work horse that does the copying
	Do while not fle1.AtEndofStream 'Change this line, Change this one too
		strLine = fle1.ReadLine
		select Case strLine
			case "Change this line"
				'When the above line is found, it is replaced with the line below
				fle2.WriteLine "Changed"
			case "Change this one too"
				fle2.WriteLine "This line is changed"
			case else
				'This copies whatever was read in fle1
				fle2.WriteLine strLine
		end select
	loop
	if err.number <> 0 then
		msgbox "Error transfering data: " & err.number & "; " & err.description
		CopyStuff = 5
		fle1.close
		fle2.close
		exit function
	end if
	fle1.close
	 set fle1 = nothing
	fle2.close
	 set fle2 = nothing
	objFSO.DeleteFile strPath, true	'This deletes the original file
	objFSO.MoveFile strFldr, strPath 'This moves and renames the temp file, replacing the original
	if err.number <> 0 then
		msgbox "Error replacing " & strPath & " with new file: " & err.number & "; " & err.description
		CopyStuff = 6
	else
		CopyStuff = 1 'Remember that in Main, a 1 means successful
	end if
end function
```


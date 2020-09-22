<div align="center">

## Create a HTML\-file from ASP


</div>

### Description

To create a static HTML-page from an ASP, rather than repeatedly recreate the same info in the ASP-page
 
### More Info
 
The code is fully/heavily/well commented


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Basie van Heerden](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/basie-van-heerden.md)
**Level**          |Beginner
**User Rating**    |3.8 (19 globes from 5 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Server Side](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/server-side__4-31.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/basie-van-heerden-create-a-html-file-from-asp__4-6800/archive/master.zip)





### Source Code

```
'---------------------------------------------------------------------
' How to create a .txt-file. (or .html, .asp, whatever)
'
' In this example the file is only created when a condition, which
' YOU supply, is met. I use it to create a static .html-page once a
' day in order to diminish time spent on gathering yesterday's data,
' doing calculations on it, and then send it as html to the client.
' By doing this once and then store the info as a static html-page,
' I save on time and server-resources.
'
' I hope you find some use for this code!
'
' Basie van Heerden (basie.v.heerden@ast.co.za)
'---------------------------------------------------------------------
<%@ LANGUAGE=VBSCRIPT%>
<% OPTION EXPLICIT
'ALWAYS first declare the variables you're going to use!
Dim objFName
dim FName
Dim HTMLname
'This subroutine is where the file is created, but only if the set
'condition is met, else this page only redirects to the previously
'created page
Sub MakeHTML()
	On Error Resume Next
	'Put the path and name of the file to be created in a variable
	HTMLname = "c:\inetpub\whatever\TestPage.htm"
	'Tell the system that you want to create a file
	set objFName = server.CreateObject("Scripting.FileSystemObject")
	'Open the .html-file, overwrite the existing file, unicode
	set FName = objFName.CreateTextFile(HTMLname,8,True)
	'Write the whole file. It can be HTML, like in this example, or
	'it can be another ASP-page, a text file, whatever.
  FName.writeline("<HTML>")
  FName.writeline("<HEAD><TITLE>TEST</TITLE></HEAD>")
  FName.writeline("<BODY bgcolor=white>")
  FName.writeline("Hello world! Whatever...")
  FName.writeline("</BODY>")
  FName.writeline("</HTML>")
  'Close and clean up behind you...
	FName.close
	set FName = nothing
End sub
'Create the file on a given condition, eg when the page is accessed
'for the first time each day/week/month/whatever. After that and
'until the condition is met, just redirect to the created HTML-page.
if condition then
	MakeHTML()
end if
Response.Redirect "TestPage.htm"
Response.End
%>
```


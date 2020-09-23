<div align="center">

## ASP 8 Queens solution


</div>

### Description

To find the solutions to the 8 queens puzzle. The puzzle asks "How can you place 8 Queens on a chess board so that no Queen can attack any other queen?"
 
### More Info
 
All the user has to input is the number of solutions they would like to see. (There are 92 possible ways to arrange 8Queens on a chess board so that no Queen can attack any other queen.

A recursive Algorithm is used. I copied my own VB code (also available on PlanetSourceCode) Only minor changes were required to run it as ASP. I generally do the "Hello World" exercise then I use the 8 Queens puzzle as my way of learning any new programming language. It forces me to learn most of the basic language structures (multi dimentional arrays, Functions, Display and/or printing, getting input etc.) By the time I finish the 8 Queens puzzle I usually feel pretty comfortable with a new language.

The result is a series of graphical displays that show how to arrange 8 Queens on a chess board so that no Queen can attack any other queen.

There are nearly 17,000,000 different ways to arrange 8 Queens on a chess board so if you ask to see all 92 solutions this code has to try them all. On a slower computer that can take 2 or 3 minutes since this is a script not compiled code.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Richard Lindsay](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/richard-lindsay.md)
**Level**          |Beginner
**User Rating**    |5.0 (25 globes from 5 users)
**Compatibility**  |ASP \(Active Server Pages\), VbScript \(browser/client side\)

**Category**       |[Algorithims](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/algorithims__4-29.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/richard-lindsay-asp-8-queens-solution__4-6222/archive/master.zip)





### Source Code

```
<%@ LANGUAGE=VBScript %>
<HTML><BODY>
<% Server.ScriptTimeOut=200 %>
<% If Request.Form("NumSolutions") = "" Then %>
<FORM method=Post>
This web page can show you all the solutions to the 8-Queens puzzle. There are 92
possible solutions. Please enter the number of solutions you would like to see then
click the submit button.<BR><BR>
This demonstrates the use of functions in an ASP page. A recursive algorithim
is used to solve the problem. If you are familiar with VB and want to do a compairison
you can look at my 8 Queens solution in the VB section of www.planetsourcecode.com.
This is almost a pure cut and paste. It only took about 10 minutes to make it into
an ASP application. The only real changes are in the Sub UpdateDisplay() other than
that all I did was delete all the data typing since in VBScript everything is a Variant.<BR><BR>
<INPUT name=NumSolutions>
<INPUT type=submit value="Submit">
<% Else %>
<%
Dim NumToFind
Dim mblnBoard(7, 7)
Dim solutions
NumToFind = Request.Form ("NumSolutions")
bln8Queens(0)
Function blnCheckUp(ByVal intRow, ByVal intColumn)
blnCheckUp = True
Do Until intRow = 0 Or intColumn = 0
  intRow = intRow - 1
  intColumn = intColumn - 1
  If mblnBoard(intRow, intColumn) Then
    blnCheckUp = False
    Exit Function
  End If
Loop
End Function
Function blnCheckDown(ByVal intRow, ByVal intColumn)
blnCheckDown = True
Do Until intRow = 7 Or intColumn = 0
  intRow = intRow + 1
  intColumn = intColumn - 1
  If mblnBoard(intRow, intColumn) Then
    blnCheckDown = False
    Exit Function
  End If
Loop
End Function
Function blnCheckAcross(ByVal intRow, ByVal intColumn)
blnCheckAcross = True
Do Until intColumn = 0
  intColumn = intColumn - 1
  If mblnBoard(intRow, intColumn) Then
    blnCheckAcross = False
    Exit Function
  End If
Loop
End Function
Function blnIsSafe(ByVal intRow, ByVal intColumn)
  blnIsSafe = blnCheckUp(intRow, intColumn) And blnCheckDown(intRow, intColumn) And blnCheckAcross(intRow, intColumn)
End Function
Sub UpdateDisplay()
	Dim intRow
	Dim intColumn
	Response.Write "<HR>Solution #" & solutions & "<BR>"
	Response.Write "<table border=1 align=Left width=190 cellborder=1 cellspacing=0>"
	For intRow = 0 to 7
		Response.Write "<TR>"
		For intColumn = 0 to 7
			If (intColumn + intRow) MOD 2 = 1 Then
				Response.Write "<td bgcolor =""#000000"""
			Else
				Response.Write "<td"
			End If
			If mblnBoard(intRow,intColumn) = True Then
				Response.Write " align=center><font color =""#FF0066""> Q </font></td>"
			Else
				Response.Write "> &nbsp </td>"
			End If
		Next
		Response.Write "</tr>"
	Next
	Response.Write "</table><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR>"
End Sub
Function bln8Queens(ByVal intColumn)
  Dim intRow
  Do
    If intColumn = 7 And blnIsSafe(intRow, intColumn) Then
      mblnBoard(intRow, intColumn) = True
	  solutions = solutions + 1
      UpdateDisplay
      If CInt(solutions) >= CInt(Request.Form ("NumSolutions")) Then
		bln8Queens = True
        mblnBoard(intRow, intColumn) = False
        Exit Function
      End If
      mblnBoard(intRow, intColumn) = False
      intRow = intRow + 1
    Else
      If blnIsSafe(intRow, intColumn) Then
        mblnBoard(intRow, intColumn) = True
        bln8Queens = bln8Queens(intColumn + 1)
        mblnBoard(intRow, intColumn) = False
      End If
      intRow = intRow + 1
    End If
  Loop Until bln8Queens = True Or intRow = 8
End Function
%>
<% End If %>
</BODY></HTML>
```


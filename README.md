<div align="center">

## GetToken


</div>

### Description

The following code is a Visual Basic function that returns a specific "token" (section/substring of data) from a delimited string list. The function accepts the index of the desired token and also the delimiter as specified by the programmer.
 
### More Info
 
Requires : [string] delimited data, [integer] index of desired section, [string] delimiter (1 or more chars)

Examples:

GetToken("steve@hotmail.com", 2, "@") returns "hotmail.com"

GetToken("first,second,third", 2, ",") returns "second"

GetToken("111, 222, 333", 3, ", ") returns "333"

GetToken("line1" + vbCrLf + "line2" + vbCrLf + "line3", 2, vbCrLf) returns "line2"

Returns : [string] "Token" (section of data) from a list of delimited string data


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Troy DeMonbreun](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/troy-demonbreun.md)
**Level**          |Unknown
**User Rating**    |5.0 (25 globes from 5 users)
**Compatibility**  |VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/troy-demonbreun-gettoken__1-1280/archive/master.zip)





### Source Code

```
Function GetToken(ByVal strVal As String, intIndex As Integer, _
	strDelimiter As String) As String
'------------------------------------------------------------------------
' Author  : Troy DeMonbreun (vb@8x.com)
'
' Returns : [string] "Token" (section of data) from a list of
'      delimited string data
'
' Requires : [string] delimited data,
'      [integer] index of desired section,
'      [string] delimiter (1 or more chars)
'
' Examples : GetToken("steve@hotmail.com", 2, "@") returns "hotmail.com"
'      GetToken("123-45-6789", 2, "-") returns "45"
'      GetToken("first,middle,last", 3, ",") returns "last"
'
' Revised : 12/22/1998
'------------------------------------------------------------------------
	Dim strSubString() As String
	Dim intIndex2 As Integer
	Dim i As Integer
	Dim intDelimitLen As Integer
	intIndex2 = 1
	i = 0
	intDelimitLen = Len(strDelimiter)
	Do While intIndex2 > 0
		ReDim Preserve strSubString(i + 1)
		intIndex2 = InStr(1, strVal, strDelimiter)
		If intIndex2 > 0 Then
			strSubString(i) = Mid(strVal, 1, (intIndex2 - 1))
			strVal = Mid(strVal, (intIndex2 + intDelimitLen), Len(strVal))
		Else
			strSubString(i) = strVal
		End If
		i = i + 1
	Loop
	If intIndex > (i + 1) Or intIndex < 1 Then
		GetToken = ""
	Else
		GetToken = strSubString(intIndex - 1)
	End If
End Function
```


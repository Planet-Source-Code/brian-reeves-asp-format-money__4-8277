<div align="center">

## ASP Format Money


</div>

### Description

This function properly formats a number to viewable currency ($x,xxx.xx) to be displayed on a webpage (or saved internally). Also performs rounding based on the constant. Will round up, down, and normal (5-9 = +1; else drop digits)
 
### More Info
 
Number (to be converted into money)

String representing the money in the format of

$xx,xxx.xx (where x is a number)

Takes into account user errors (entering in < 2 digits after period).


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Brian Reeves](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/brian-reeves.md)
**Level**          |Intermediate
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |ASP \(Active Server Pages\), VbScript \(browser/client side\)

**Category**       |[Strings](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/strings__4-26.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/brian-reeves-asp-format-money__4-8277/archive/master.zip)

### API Declarations

Open Source


### Source Code

```
'Returns a string from a number to be displayed in a $xx.xx format
Public Function FormatMoney(sString)
	'change cRound to Normal, Up, Down
	Const cRound = "Normal"
	If InStr(sString, ".") Then
		'Adding extra zero's at the end for error correction (User passes "23." and it
		'	will still display correctly)
		FormatMoney = sString & "000"
		Select Case cRound
			Case "Normal"
				If Mid(FormatMoney, InStr(FormatMoney, ".") + 3) > 4 Then
					FormatMoney = Left(FormatMoney, InStr(FormatMoney, ".") + 1) & Fix(Mid(FormatMoney, InStr(FormatMoney, ".") + 2, 1) + 1)
				End If
			Case "Up"
				If Mid(FormatMoney, InStr(FormatMoney, ".") + 3) > 0 Then
					FormatMoney = Left(FormatMoney, InStr(FormatMoney, ".") + 1) & Fix(Mid(FormatMoney, InStr(FormatMoney, ".") + 2, 1) + 1)
				End If
		End Select
		FormatMoney = Left(FormatMoney, InStr(FormatMoney, ".") + 2)
	Else: FormatMoney = sString & ".00" 'Appending cents to the dollar
	End If
	FormatMoney = "$" & FormatMoney	'Appending dollar sign to beginning
End Function
```


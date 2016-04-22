
Value = "@PIN@"

Dim matchMin, matchMax
matchMin = 1
matchMax = 10

'regex example from .chm
Dim regEx, Match, Matches, numMatch, numActuallyPrinting 'creates all variables
Set regEx = New RegExp
regEx.Pattern = "\d{4}-"
regEx.Global = True

Set Matches = regEx.Execute(Field("Sheet1$.Applications"))

numMatch = 0
numActuallyPrinting = 0
For Each Match in Matches  'loop through "Matches" collection
	numMatch = numMatch + 1
	If numMatch >= matchMin  And numMatch <= matchMax Then
		Value = Value & Left(Match.Value, 4) 'Left removes the "-" from regex match
		numActuallyPrinting = numActuallyPrinting + 1
	End If
Next

Do While numMatch < matchMax
	Value = Value & "####"
	numMatch = numMatch + 1
Loop


If numActuallyPrinting = 0 Then Value = "n/a"


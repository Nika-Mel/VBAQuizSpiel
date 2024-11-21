Private loesung As String

Private Sub CommandButton1_Click()
Dim eingabe As String
eingabe = Me.TextBox1.text
Call LoesungVergleichen(eingabe)
End Sub

Private Sub CommandButton2_Click()
Call FrageStellen
End Sub

Private Sub Label1_Click()
End Sub

Private Sub Label2_Click()
End Sub

Private Sub Label3_Click()
End Sub

Private Sub UserForm_Click()
End Sub

Private Sub UserForm_Initialize()
Me.CommandButton1.Caption = "LÖSE AUF"
Me.CommandButton1.Font.Size = 12
Me.CommandButton2.Caption = "NEUE FRAGE"
Me.Label2.TextAlign = fmTextAlignCenter
Me.Label2.Font.Size = 12
Me.Label1.Font.Size = 18
Me.Label1.Font.Underline = True
Me.Label1.TextAlign = fmTextAlignCenter
Me.Label3.TextAlign = fmTextAlignCenter
Me.Label3.Font.Size = 12
Me.TextBox1.Font.Size = 12
Call FrageStellen
End Sub

Sub LoesungVergleichen(text As String)
Dim loesungVergleich As String
loesungVergleich = Replace(loesung, " ", "")
text = Replace(text, " ", "")
If loesungVergleich = text Then
Me.Label3.Caption = "Die Antwort ist richtig."
Me.Label3.ForeColor = RGB(0, 205, 0)
Else
Me.Label3.Caption = "Die Antwort ist falsch." & vbNewLine & vbNewLine & "Die Antwort war:"
& vbNewLine & loesung
Me.Label3.ForeColor = RGB(238, 0, 0)
End If
End Sub

Sub FrageStellen()
Dim frage As String
Dim a As Integer
Dim b As Integer
Dim c As Integer
Dim ax As Integer
Dim axx As Integer 'ax²'
Dim bb As Integer 'b²'
Dim lowerBoundFrage As Integer
Dim upperBoundFrage As Integer
Dim randomNumberFrage As Integer
lowerBoundFrage = 1
upperBoundFrage = 3
randomNumberFrage = Int((upperBoundFrage - lowerBoundFrage + 1) * Rnd + lowerBoundFrage)
'Für Bestimmung der Werte a und b'
Dim lowerBoundZahl As Integer
Dim upperBoundZahl As Integer
lowerBoundZahl = 1
upperBoundZahl = 5
a = Int((upperBoundZahl - lowerBoundZahl + 1) * Rnd + lowerBoundZahl) 'Bestimme Wert a'
b = Int((upperBoundZahl - lowerBoundZahl + 1) * Rnd + lowerBoundZahl) 'Bestimme Wert b'
c = Int((upperBoundZahl - lowerBoundZahl + 1) * Rnd + lowerBoundZahl) 'Bestimme Wert c'
'Erstelle Frage'
If randomNumberFrage = 1 Then
frage = "(" & a & "x + " & b & ")² =" '(ax + b)²'
axx = a * a
ax = 2 * (a * b)
bb = b * b
loesung = axx & "x² + " & ax & "x + " & bb
ElseIf randomNumberFrage = 2 Then
frage = "(" & a & "x - " & b & ")² =" '(ax - b)²'
axx = a * a
ax = 2 * (a * b)
bb = b * b
loesung = axx & "x² - " & ax & "x + " & bb
Else
frage = "(" & a & "x + " & b & ")(" & a & "x - " & c & ") =" '(ax + b)(ax - c)'
axx = a * a
ax = (a * b) - (a * c)
bb = b * c
If ax < 0 Then
loesung = axx & "x² " & ax & "x - " & bb
Else
loesung = axx & "x² + " & ax & "x - " & bb
End If
End If
Me.Label2.Caption = frage
End Sub

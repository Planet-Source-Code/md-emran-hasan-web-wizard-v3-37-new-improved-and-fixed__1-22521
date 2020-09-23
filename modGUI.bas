Attribute VB_Name = "modGUI"
Public Sub ShowContent(Index As Integer)
Select Case Index

Case 0
frmMain.picStart(0).Top = 150
For i = 1 To 7
frmMain.picStart(i).Top = 9000
Next i

Case 1
frmMain.picStart(1).Top = 150
frmMain.picStart(0).Top = 9000
For i = 2 To 7
frmMain.picStart(i).Top = 9000
Next i

Case 2
frmMain.picStart(1).Top = 9000
frmMain.picStart(0).Top = 9000
frmMain.picStart(2).Top = 150
For i = 3 To 7
frmMain.picStart(i).Top = 9000
Next i
frmMain.Label1(Index - 1).FontBold = True
frmMain.shpStart(Index - 1).FillColor = vbGreen

Case 3
For i = 0 To 2
frmMain.picStart(i).Top = 9000
Next i
frmMain.picStart(3).Top = 150
For i = 4 To 7
frmMain.picStart(i).Top = 9000
Next i

Case 4
For i = 0 To 3
frmMain.picStart(i).Top = 9000
Next i
frmMain.picStart(4).Top = 150
For i = 5 To 7
frmMain.picStart(i).Top = 9000
Next i

Case 5
For i = 0 To 4
frmMain.picStart(i).Top = 9000
Next i
frmMain.picStart(5).Top = 200
For i = 6 To 7
frmMain.picStart(i).Top = 9000
Next i

Case 6
For i = 0 To 5
frmMain.picStart(i).Top = 9000
Next i
frmMain.picStart(6).Top = 150
frmMain.picStart(7).Top = 9000

Case 7
For i = 0 To 6
frmMain.picStart(i).Top = 9000
Next i
frmMain.picStart(7).Top = 150

End Select
End Sub

Public Sub ChangeColor(Index As Integer)
Select Case Index

Case 0
frmMain.Label1(0).FontBold = True
frmMain.shpStart(0).FillColor = vbGreen
For i = 1 To 7
frmMain.Label1(i).FontBold = False
frmMain.shpStart(i).FillColor = &HC0C0C0
Next i

Case 1
frmMain.Label1(1).FontBold = True
frmMain.shpStart(1).FillColor = vbGreen
frmMain.Label1(0).FontBold = False
frmMain.shpStart(0).FillColor = &HC0C0C0
For i = 2 To 7
frmMain.Label1(i).FontBold = False
frmMain.shpStart(i).FillColor = &HC0C0C0
Next i

Case 2
frmMain.Label1(2).FontBold = True
frmMain.shpStart(2).FillColor = vbGreen
For i = 0 To 1
frmMain.Label1(i).FontBold = False
frmMain.shpStart(i).FillColor = &HC0C0C0
Next i
For i = 3 To 7
frmMain.Label1(i).FontBold = False
frmMain.shpStart(i).FillColor = &HC0C0C0
Next i

Case 3
frmMain.Label1(3).FontBold = True
frmMain.shpStart(3).FillColor = vbGreen
For i = 0 To 2
frmMain.Label1(i).FontBold = False
frmMain.shpStart(i).FillColor = &HC0C0C0
Next i
For i = 4 To 7
frmMain.Label1(i).FontBold = False
frmMain.shpStart(i).FillColor = &HC0C0C0
Next i

Case 4
frmMain.Label1(4).FontBold = True
frmMain.shpStart(4).FillColor = vbGreen
For i = 0 To 3
frmMain.Label1(i).FontBold = False
frmMain.shpStart(i).FillColor = &HC0C0C0
Next i
For i = 5 To 7
frmMain.Label1(i).FontBold = False
frmMain.shpStart(i).FillColor = &HC0C0C0
Next i

Case 5
frmMain.Label1(5).FontBold = True
frmMain.shpStart(5).FillColor = vbGreen
For i = 0 To 4
frmMain.Label1(i).FontBold = False
frmMain.shpStart(i).FillColor = &HC0C0C0
Next i
For i = 6 To 7
frmMain.Label1(i).FontBold = False
frmMain.shpStart(i).FillColor = &HC0C0C0
Next i

Case 6
frmMain.Label1(6).FontBold = True
frmMain.shpStart(6).FillColor = vbGreen
For i = 0 To 5
frmMain.Label1(i).FontBold = False
frmMain.shpStart(i).FillColor = &HC0C0C0
Next i
frmMain.Label1(7).FontBold = False
frmMain.shpStart(7).FillColor = &HC0C0C0

Case 7
frmMain.Label1(7).FontBold = True
frmMain.shpStart(7).FillColor = vbGreen
For i = 0 To 6
frmMain.Label1(i).FontBold = False
frmMain.shpStart(i).FillColor = &HC0C0C0
Next i

End Select
End Sub


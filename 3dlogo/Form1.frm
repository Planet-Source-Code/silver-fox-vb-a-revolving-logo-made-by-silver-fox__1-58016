VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   4200
      Top             =   2280
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************************************************************
'*
'*   Do whatever you want with this code (Dont bother me!..  ;o)
'*   But if you do use it in something, send me a copy..
'*
'*silver_fox_vb@ yahoo.com
'*
'*   Created because I needed to display a 3D spinning logo
'*   without using DirectX.
'*
'*   BEWARE!!..
'*   The code is a mess (and as you can see I`m not big on comments..  ;o)
'*
'***********************************************************************************


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 39 Then strtxtX = strtxtX + 1
If KeyCode = 37 Then strtxtX = strtxtX - 1
If KeyCode = 38 Then strtxtY = strtxtY + 1
If KeyCode = 40 Then strtxtY = strtxtY - 1
If KeyCode = 46 Then strtxtZ = strtxtZ + 1
If KeyCode = 34 Then strtxtZ = strtxtZ - 1


'turn perspective on/off
If KeyCode = 36 And perssel = True Then
    perssel = False
ElseIf KeyCode = 36 And perssel = False Then
    perssel = True
End If

End Sub

Private Sub Form_Load()

perssel = True
KeyPreview = True
        
strtxtX = 0  ' speed for spin
strtxtY = 0
strtxtZ = 0
        
Call LoadPoints

For i = 1 To TotalPoints
    Rt(i, 1) = Po(i, 1)
    Rt(i, 2) = Po(i, 2)
    Rt(i, 3) = Po(i, 3)
Next i


End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then Call Command1_Click

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then Call Command1_Click

End Sub

Private Sub cmdPlot_Click()

iOrgX = Form1.ScaleWidth \ 2
iOrgY = Form1.ScaleHeight \ 2

Form1.Picture = LoadPicture("")

mainloopBOT = 1
mainloopTOP = 9

Form1.AutoRedraw = True

For mainloop = 1 To (TotalPoints / PointsSplit) ' how many seperate parts
    
    DrawMode = 5
    'DrawStyle = 0
    'DrawWidth = 3
    
    For i = mainloopBOT To mainloopTOP '5
        Form1.Line (iOrgX + Rt(i, 1) + MoveX, iOrgY + Rt(i, 2) + MoveY)-(iOrgX + Rt(i + 1, 1) + MoveX, iOrgY + Rt(i + 1, 2) + MoveY)
    Next i
    
    mainloopBOT = mainloopBOT + PointsSplit
    mainloopTOP = mainloopTOP + PointsSplit

Next mainloop


End Sub


Public Sub LoadPoints()

'    Set excel_app = CreateObject("Excel.Application")
'
'    On Error Resume Next
'    Set excel_app = GetObject("\\felix\grpdata\Applications\ContractorsDB\pointsCDB.xls")
'
'    If Val(excel_app.Application.Version) >= 8 Then
'        Set excel_sheet = excel_app.ActiveSheet
'    Else
'        Set excel_sheet = excel_app
'    End If
'
'    TotalPoints = excel_sheet.Cells(1, 4) ' get total points
'
'    If TotalPoints = 0 Then
'        MsgBox "No total points in your points list!", vbOKOnly, "ERROR!!"
'        End
'    ElseIf TotalPoints > 1999 Then
'        MsgBox "to many total points for this version!", vbOKOnly, "ERROR!!"
'        End
'    End If
'
'
'
'    ' Get the points.
'    For i3 = 2 To TotalPoints + 1
'        strtxtXKEEP = excel_sheet.Cells(i3, 1) * 15
'        strtxtYKEEP = excel_sheet.Cells(i3, 2) * 15
'        strtxtZKEEP = excel_sheet.Cells(i3, 3) * 15
'        Po(i3 - 1, 1) = strtxtXKEEP
'        Po(i3 - 1, 2) = strtxtYKEEP
'        Po(i3 - 1, 3) = strtxtZKEEP
'
'    Next i3
'
'    Set excel_sheet = Nothing
'    Set excel_app = Nothing
    
    
    
Call pointsCDBLoad
    
'Call pointsBlockLoad
    
'Call pointsCircleLoad

'Call pointsGrassLoad
    
    
End Sub

Private Sub Timer1_Timer()
Dim SinA As Single
Dim CosA As Single
Dim SinB As Single
Dim CosB As Single
Dim SinC As Single
Dim CosC As Single

Dim X As Integer
Dim Y As Integer
Dim z As Integer

Dim xx As Integer
Dim yy As Integer
Dim zz As Integer

Dim i As Integer


' timer for bounce move

MoveX = X
MoveY = Y

'----------------------

ZoomINT = ZoomINT + 10

RotA = RotA + Val(strtxtX)
If RotA > 360 Then RotA = RotA - 360

RotB = RotB + Val(strtxtY)
If RotB > 360 Then RotB = RotB - 360

RotC = RotC + Val(strtxtZ)
If RotC > 360 Then RotC = RotC - 360

' Calc the angles
riaX = RotA
riaY = RotB
riaZ = RotC
riStart = 1
riEnd = TotalPoints

SinA = Sin(riaX * 3.1415926 / 180)
SinB = Sin(riaY * 3.1415926 / 180)
SinC = Sin(riaZ * 3.1415926 / 180)

CosA = Cos(riaX * 3.1415926 / 180)
CosB = Cos(riaY * 3.1415926 / 180)
CosC = Cos(riaZ * 3.1415926 / 180)

For i = riStart To riEnd
    
    X = Po(i, 1)
    Y = Po(i, 2)
    z = Po(i, 3)
    
    xx = Int(X * CosA * CosC + Y * SinA * SinB * CosC - z * SinA * CosB * CosC - Y * CosB * SinC - z * SinB * SinC)
    yy = Int(X * CosA * SinC + Y * SinA * SinB * SinC - z * SinA * CosB * SinC + Y * CosB * CosC + z * SinB * CosC)
    zz = Int(X * SinA - Y * CosA * SinB + z * CosA * CosB)

    'perssel = True
    
    If perssel = True Then
        '*********** Perspective bit ***********
        
        Z2 = 1 + zz / 3250
        If Z2 = 0 Then Z2 = 0.1
    
        xx = xx / Z2
        yy = yy / Z2
        zz = zz / Z2
        Label1.Caption = Format(Z2, "0.000")
        
        '***************************************
    End If


    Rt(i, 1) = xx
    Rt(i, 2) = yy
    Rt(i, 3) = zz

Next i


cmdPlot_Click
End Sub

VERSION 5.00
Begin VB.Form formLife 
   BackColor       =   &H00000000&
   Caption         =   "Evolving Artificial Life"
   ClientHeight    =   6435
   ClientLeft      =   1170
   ClientTop       =   915
   ClientWidth     =   8160
   FillColor       =   &H008080FF&
   ForeColor       =   &H000000FF&
   Icon            =   "formLife.frx":0000
   LinkTopic       =   "formLife"
   ScaleHeight     =   6435
   ScaleWidth      =   8160
   Begin VB.CommandButton Command1 
      Caption         =   "Remove Animals"
      Height          =   492
      Left            =   6480
      TabIndex        =   9
      Top             =   5880
      Width           =   732
   End
   Begin VB.CommandButton cmdDelPlants 
      Caption         =   "Remove Plants"
      Height          =   492
      Left            =   960
      TabIndex        =   8
      Top             =   5880
      Width           =   732
   End
   Begin VB.Timer tmrMultiTask 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   120
      Top             =   5760
   End
   Begin VB.PictureBox picLife 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   5172
      Left            =   120
      ScaleHeight     =   341
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   525
      TabIndex        =   7
      Top             =   120
      Width           =   7932
   End
   Begin VB.CommandButton cmdAddAnts 
      Caption         =   "Add Animals"
      Height          =   252
      Left            =   5040
      TabIndex        =   6
      Top             =   5880
      Width           =   1212
   End
   Begin VB.CommandButton cmdAddPlants 
      Caption         =   "Add Plants"
      Height          =   252
      Left            =   1920
      TabIndex        =   5
      Top             =   5880
      Width           =   1212
   End
   Begin VB.CommandButton cmdRestart 
      Caption         =   "Restart"
      Height          =   372
      Left            =   3600
      TabIndex        =   4
      Top             =   5880
      Width           =   972
   End
   Begin VB.Label lblAntCountLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Number of animals:"
      ForeColor       =   &H00C0C0FF&
      Height          =   252
      Left            =   4320
      TabIndex        =   3
      Top             =   5520
      Width           =   2052
   End
   Begin VB.Label lblPlantCountLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Number of plants:"
      ForeColor       =   &H00C0C0FF&
      Height          =   252
      Left            =   480
      TabIndex        =   2
      Top             =   5520
      Width           =   2052
   End
   Begin VB.Label lblAntCount 
      BackColor       =   &H00000000&
      Caption         =   "0"
      ForeColor       =   &H00C0C0FF&
      Height          =   252
      Left            =   6480
      TabIndex        =   1
      Top             =   5520
      Width           =   1212
   End
   Begin VB.Label lblPlantCount 
      BackColor       =   &H00000000&
      Caption         =   "0"
      ForeColor       =   &H00C0C0FF&
      Height          =   252
      Left            =   2640
      TabIndex        =   0
      Top             =   5520
      Width           =   1212
   End
   Begin VB.Menu mOptions 
      Caption         =   "&Options"
      Begin VB.Menu mChkInvertRed 
         Caption         =   "Invert red level"
      End
      Begin VB.Menu mChkInvertGreen 
         Caption         =   "Invert green level"
      End
      Begin VB.Menu mChkInvertBlue 
         Caption         =   "Invert blue level"
      End
      Begin VB.Menu mChkNoCloneAnt 
         Caption         =   "Prohibit self-reproducing animals"
      End
      Begin VB.Menu mChkNoClonePlant 
         Caption         =   "Prohibit self-reproducing plants"
      End
      Begin VB.Menu mChkNoMateAntAnt 
         Caption         =   "Prohibit animals from mating with animals"
      End
      Begin VB.Menu mChkNoMateAntPlant 
         Caption         =   "Prohibit animals from mating with plants"
      End
      Begin VB.Menu mChkNoMatePlantAnt 
         Caption         =   "Prohibit plants from mating with animals"
      End
      Begin VB.Menu mChkNoMatePlantPlant 
         Caption         =   "Prohibit plants from mating with plants"
      End
      Begin VB.Menu mChkNoAntEatAnt 
         Caption         =   "Prohibit animals from eating animals"
      End
      Begin VB.Menu mChkNoAntEatPlant 
         Caption         =   "Prohibit animals from eating plants"
      End
      Begin VB.Menu mChkNoPlantEatAnt 
         Caption         =   "Prohibit plants from eating animals"
         Checked         =   -1  'True
      End
      Begin VB.Menu mChkNoPlantEatPlant 
         Caption         =   "Prohibit plants from eating plants"
         Checked         =   -1  'True
      End
      Begin VB.Menu mChkNoAntEatCarion 
         Caption         =   "Prohibit animals from scavenging"
      End
      Begin VB.Menu mChkNoPlantEatCarion 
         Caption         =   "Prohibit plants from scavenging"
      End
      Begin VB.Menu mChkNoAntEatDirt 
         Caption         =   "Prohibit ambient feeding for animals"
      End
      Begin VB.Menu mChkAllowSpores 
         Caption         =   "Allow animals to mate over a distance"
      End
      Begin VB.Menu mChkAllowPollen 
         Caption         =   "Allow plants to mate over a distance"
         Checked         =   -1  'True
      End
      Begin VB.Menu mChkAllowPlagues 
         Caption         =   "Allow plagues"
         Checked         =   -1  'True
      End
      Begin VB.Menu mChkQuickStart 
         Caption         =   "Quick-start evolution"
         Checked         =   -1  'True
      End
      Begin VB.Menu mChkFastEvolve 
         Caption         =   "High rate of mutation"
         Checked         =   -1  'True
      End
   End
End
Attribute VB_Name = "formLife"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetTickCount Lib "kernel32" () As Long

Dim initialHeight As Single
Dim initialWidth As Single
Private ExitForm As Boolean

Private Sub cmdAddAnts_Click()
    'Add a user specified number of primateve live animals.
    Dim s As String
    Dim n
    s = InputBox("Add how many random primative animals?" + vbCrLf + vbCrLf + "0 = cancel", "Add Animals", "1")
    If s = "" Then Exit Sub
    On Error Resume Next
    n = -Val(s)
    On Error GoTo 0
    If n = 0 Then Exit Sub
    firstLife n
End Sub

Private Sub cmdAddPlants_Click()
    'Add a user specified number of primateve live plants.
    Dim s As String
    Dim n
    s = InputBox("Add how many primative plants?" + vbCrLf + vbCrLf + "0 = cancel", "Add Plants", "1")
    If s = "" Then Exit Sub
    On Error Resume Next
    n = Val(s)
    On Error GoTo 0
    If n = 0 Then Exit Sub
    firstLife n
End Sub

Private Sub cmdDelPlants_Click()
    'Delete a user specified number of plants.
    Dim s As String
    Dim n
    s = InputBox("Remove how many plants?" + vbCrLf + vbCrLf + "0 = cancel", "Delete Plants", "1")
    If s = "" Then Exit Sub
    On Error Resume Next
    n = Val(s)
    On Error GoTo 0
    If n = 0 Then Exit Sub
    deleteLife n
End Sub

Private Sub cmdRestart_Click()
    'Start over with a user specified number of primateve live plants.  (choosing a negative number will start with animals.)
    Dim s As String
    Dim n
    picLife.BorderStyle = 1 'show the pictureBox border.
    s = InputBox("Restart with how many primative plants?" + vbCrLf + vbCrLf + "0 = random", "Restart", "1")
    If s = "" Then
        picLife.BorderStyle = 0 'hide the pictureBox border.
        Exit Sub
    End If
    On Error Resume Next
    n = Val(s)
    On Error GoTo 0
    If n = 0 Then n = Int(Rnd * Rnd * 3333 + 1) 'pick a random number of plants and finish re-initializing.
    postInitialize n 're-initialize life form environment.
End Sub

Private Sub Command1_Click()
    'Delete a user specified number of Animals.
    Dim s As String
    Dim n
    s = InputBox("Remove how many animals?" + vbCrLf + vbCrLf + "0 = cancel", "Delete Animals", "1")
    If s = "" Then Exit Sub
    On Error Resume Next
    n = Val(s)
    On Error GoTo 0
    If n = 0 Then Exit Sub
    deleteLife -n
End Sub

Private Sub Form_Activate()
    Static postinitialized As Boolean 'initially false
    
    Dim LastTick As Long
    
    If Not postinitialized Then
        postInitialize Int(Rnd * Rnd * 3333 + 1) 'pick a random number of plants and finish initializing.
        postinitialized = True
    
    
        Do
            'multitasking core.
            Static taskNumber As Integer
            Static X As Single
            Static Y As Single
            Dim XX As Single
            Dim YY As Single
            Dim i As Integer 'loop counter
            
            For i = 0 To 25 'arbitrary number of loops.  Low numbers are slower. High numbers are less responsive.
                taskNumber = taskNumber + 1
                If taskNumber > 8 Then taskNumber = 0
                
                Select Case taskNumber
                Case 0
                    GetNextItemXY X, Y
                Case 1
                    moveItemXY X, Y
                Case 2
                    mateItemXY X, Y
                Case 3
                    cloneItemXY X, Y
                Case 4
                    feedItemXY X, Y
                Case 5
                    ageItemXY X, Y
                Case 6
                    processNextGeneXY X, Y
                Case 7
                    If (PlantCount < 0) Or (AntCount < 0) Or (lifeCount < 0) Then needReDraw = True
                Case Else
                    'if more than half a second has passed, then allow processing events
                    If ((GetTickCount - LastTick) > 500) Then
                        DoEvents
                        If needReDraw Then
                            needReDraw = False
                        End If
                        
                        LastTick = GetTickCount
                    End If  'has more than half a second passed
                End Select
            Next i
            
            If lifeCount < 1 Then 'Odds are "supposed to be" against all of this...
                'Simulate long term interactions of biochemicals in "lifeless" organic materials.
                If Grid(X, Y).Energy < 0 Then
                    If Not Grid(X, Y).Alive Then
                        XX = GetRndInt(GridSizeX - 1)
                        YY = GetRndInt(GridSizeY - 1)
                        If Sqr((XX - X) ^ 2 + (YY - Y) ^ 2) < Abs(Grid(X, Y).Energy) Then
                            If (Grid(X, Y).RGB = Grid(XX, YY).RGB) And Not Grid(XX, YY).Alive Then
                                Grid(X, Y).Energy = -Grid(X, Y).Energy
                                BirthItemXY XX, YY
                            End If
                        End If
                    End If
                End If
            End If
            
            If mChkAllowPlagues.Checked Then
                If lifeCount > 1000 Then 'Don't have plagues when the population is low.
                    If neighborCount(X, Y) > 5 Then '(over-crowded)
                        plagueItemXY X, Y
                    End If
                End If
            End If
        Loop Until ExitForm
        Unload Me
    End If
End Sub

Sub postInitialize(n) 'Initialization of environment after the form has been initialized.
    resetGrid 'Prepare the grid for life to grow in it.
    picLife.Cls 'Clear the pictureBox.
    picLife.BorderStyle = 0 'hide the pictureBox border.
    firstLife n 'Seed first n primative life forms.
    tmrMultiTask.Enabled = True 'Activate the multi-tasking core.
End Sub

Private Sub Form_Load()
    Dim ctrl As Object
    Dim sTmp As String
    
    formLife.Caption = "AntieLife - v." + Trim(Str(App.Major)) + "." + Trim(Str(App.Minor)) + ".0." + Trim(Str(App.Revision)) + " - TechnoZeus"
    initialHeight = formLife.ScaleHeight
    initialWidth = formLife.ScaleWidth
    
    On Error Resume Next
    For Each ctrl In formLife.Controls
        If ctrl.Tag > "" Then Stop
        sTmp = ""
        sTmp = sTmp + CStr(ctrl.Top)
        sTmp = sTmp + vbCrLf
        sTmp = sTmp + CStr(ctrl.Left)
        sTmp = sTmp + vbCrLf
        sTmp = sTmp + CStr(ctrl.Height)
        sTmp = sTmp + vbCrLf
        sTmp = sTmp + CStr(ctrl.Width)
        sTmp = sTmp + vbCrLf
        sTmp = sTmp + CStr(ctrl.Font.Size)
        ctrl.Tag = sTmp
    Next ctrl
    Randomize
End Sub

Private Sub Form_Paint()
    needReDraw = True  'Avoids a redraw being started durring another redraw.
End Sub

Sub drawGrid()
    'draw all life forms.
    Dim X As Single
    Dim Y As Single
    Dim LastX As Single
    Dim LastY As Single
    Dim W As Single
    
    GetNextItemXY LastX, LastY
    lifeCount = 0
    PlantCount = 0
    AntCount = 0
    
    Do
        GetNextItemXY X, Y
        DrawItemXY X, Y
        If Grid(X, Y).Alive Then
            lifeCount = lifeCount + 1
            If Grid(X, Y).Speed > 0 Then
                AntCount = AntCount + 1
            Else
                PlantCount = PlantCount + 1
            End If
        End If
    Loop Until (X = LastX) And (Y = LastY)
End Sub

Sub DrawItemXY(X As Single, Y As Single, Optional unDraw As Boolean = False)
    'draw a single life form.
    Dim W As Single
    With Grid(X, Y)
        If .Alive Then
            W = .Width
            If W < 1 Then W = 1
            picLife.DrawWidth = W  ' Set DrawWidth.
            If unDraw Then
                picLife.PSet (X * mulX + OffsetX, Y * mulY + OffsetY), 0
                picLife.Line -(.NextX * mulX + OffsetX, .NextY * mulY + OffsetY), 0
            Else
                picLife.PSet (X * mulX + OffsetX, Y * mulY + OffsetY), .RGB
                picLife.Line -(.NextX * mulX + OffsetX, .NextY * mulY + OffsetY), .RGB
            End If
        End If
    End With
End Sub

Sub DrawOverXY(X As Single, Y As Single)
    'draw life forms near x,y
    Dim XXp As Single, XXm As Single, YYp As Single, YYm As Single
    Dim Xt As Integer
    Dim Yt As Integer
    Dim W As Single
    W = CLng(Grid(X, Y).Width / 2) + 1
    If W < 1 Then W = 1
    For Xt = 0 To W
        For Yt = 0 To W
            XXp = inGridX(X + Xt)
            YYp = inGridY(Y + Yt)
            XXm = inGridX(X - Xt)
            YYm = inGridY(Y - Yt)
            If Xt > 0 Then DrawItemXY XXm, YYp
            If Yt > 0 Then DrawItemXY XXp, YYm
            If (Xt > 0) And (Yt > 0) Then DrawItemXY XXm, YYm
            DrawItemXY XXp, YYp
        Next Yt
    Next Xt
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'exit the form
    ExitForm = True
End Sub

Private Sub Form_Resize()
    'Make sure everything fits in the window.
    Dim ctrl As Object
    Dim tmpSa() As String
    Dim tmpHeight As Single
    Dim tmpWidth As Single
    Dim tmpSize As Single
    tmpHeight = heightRatio * 0.97
    tmpWidth = widthRatio * 0.99
    If tmpHeight < 0.25 Then tmpHeight = 0.25
    If tmpWidth < 0.25 Then tmpWidth = 0.25
    tmpSize = lessNotZero(tmpHeight, tmpWidth)
    On Error Resume Next
    For Each ctrl In formLife.Controls
        tmpSa = Split(ctrl.Tag, vbCrLf)
        ctrl.Top = (tmpHeight * Val(tmpSa(0)))
        ctrl.Left = (tmpWidth * Val(tmpSa(1)))
        ctrl.Height = (tmpHeight * Val(tmpSa(2)))
        ctrl.Width = (tmpWidth * Val(tmpSa(3)))
        ctrl.Font.Name = "Arial"
        ctrl.Font.Size = (tmpSize * Val(tmpSa(4)))
        ctrl.Refresh
    Next ctrl
    On Error GoTo 0
    mulX = picLife.ScaleWidth / GridSizeX
    mulY = picLife.ScaleHeight / GridSizeY
    picLife.Cls
    Form_Paint
End Sub

Private Sub mChkAllowPlagues_Click()
    mChkAllowPlagues.Checked = Not mChkAllowPlagues.Checked
End Sub

Private Sub mChkAllowPollen_Click()
    mChkAllowPollen.Checked = Not mChkAllowPollen.Checked
End Sub

Private Sub mChkAllowSpores_Click()
    mChkAllowSpores.Checked = Not mChkAllowSpores.Checked
End Sub

Private Sub mChkFastEvolve_Click()
    mChkFastEvolve.Checked = Not mChkFastEvolve.Checked
End Sub

Private Sub mChkInvertGreen_Click()
    mChkInvertGreen.Checked = Not mChkInvertGreen.Checked
    setBackgroundColor
    needReDraw = True
End Sub

Private Sub mChkInvertBlue_Click()
    mChkInvertBlue.Checked = Not mChkInvertBlue.Checked
    setBackgroundColor
    needReDraw = True
End Sub

Private Sub mChkInvertRed_Click()
    mChkInvertRed.Checked = Not mChkInvertRed.Checked
    setBackgroundColor
    needReDraw = True
End Sub

Private Sub mChkNoAntEatAnt_Click()
    mChkNoAntEatAnt.Checked = Not mChkNoAntEatAnt.Checked
End Sub

Private Sub mChkNoAntEatCarion_Click()
    mChkNoAntEatCarion.Checked = Not mChkNoAntEatCarion.Checked
End Sub

Private Sub mChkNoAntEatDirt_Click()
    mChkNoAntEatDirt.Checked = Not mChkNoAntEatDirt.Checked
End Sub

Private Sub mChkNoAntEatPlant_Click()
    mChkNoAntEatPlant.Checked = Not mChkNoAntEatPlant.Checked
End Sub

Private Sub mChkNoCloneAnt_Click()
    mChkNoCloneAnt.Checked = Not mChkNoCloneAnt.Checked
End Sub

Private Sub mChkNoClonePlant_Click()
    mChkNoClonePlant.Checked = Not mChkNoClonePlant.Checked
End Sub

Private Sub mChkNoMateAntAnt_Click()
    mChkNoMateAntAnt.Checked = Not mChkNoMateAntAnt.Checked
End Sub

Private Sub mChkNoMateAntPlant_Click()
    mChkNoMateAntPlant.Checked = Not mChkNoMateAntPlant.Checked
End Sub

Private Sub mChkNoMatePlantAnt_Click()
    mChkNoMatePlantAnt.Checked = Not mChkNoMatePlantAnt.Checked
End Sub

Private Sub mChkNoMatePlantPlant_Click()
    mChkNoMatePlantPlant.Checked = Not mChkNoMatePlantPlant.Checked
End Sub

Private Sub mChkNoPlantEatAnt_Click()
    mChkNoPlantEatAnt.Checked = Not mChkNoPlantEatAnt.Checked
End Sub

Private Sub mChkNoPlantEatCarion_Click()
    mChkNoPlantEatCarion.Checked = Not mChkNoPlantEatCarion.Checked
End Sub

Private Sub mChkNoPlantEatPlant_Click()
    mChkNoPlantEatPlant.Checked = Not mChkNoPlantEatPlant.Checked
End Sub

Private Sub mChkQuickStart_Click()
    mChkQuickStart.Checked = Not mChkQuickStart.Checked
End Sub

Function heightRatio() As Single
    heightRatio = formLife.ScaleHeight / initialHeight
End Function

Function widthRatio() As Single
    widthRatio = formLife.ScaleWidth / initialWidth
End Function

Function lessNotZero(ByVal A, ByVal B)
    'returns the lesser value, provided that value is not zero.
    'returns zero only if both values provided are zero.
    If (A = 0) Or ((B < A) And (B <> 0)) Then A = B
    lessNotZero = A
End Function

Private Sub tmrMultiTask_Timer()
    'update the display every tick
    lblPlantCount.Caption = Str(PlantCount)
    lblAntCount.Caption = Str(AntCount)
    drawGrid
End Sub

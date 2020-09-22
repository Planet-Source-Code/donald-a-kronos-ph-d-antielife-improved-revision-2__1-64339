Attribute VB_Name = "LifeMath"
Option Explicit

Private Declare Sub moveMem Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)

Public Const pi = 3.14159265358979 '(pi)
Public Const sqr2 = 1.4142135623731 '(sqr(2))
Public Const pi2 = 6.28318530717959 '(2 * pi)
Public Const invSqr2 = 0.707106781186545 '(1/sqr(2))

Public Function GetRndInt(ByVal Max As Integer, Optional ByVal Min As Integer = 0) As Integer
'This function will produce a random Integer in the specified range
    GetRndInt = Int(Rnd * (Abs(Max - Min) + 1)) * Sgn(Max - Min) + Min
End Function

Sub getRGB(ByVal sourceRGB As Long, ByRef getRed, ByRef getGreen, ByRef getBlue)
    getRed = sourceRGB And 255
    sourceRGB = sourceRGB \ 256
    getGreen = sourceRGB And 255
    sourceRGB = sourceRGB \ 256
    getBlue = sourceRGB And 255
    If formLife.mChkInvertRed.Checked Then getRed = Abs(255 - getRed)
    If formLife.mChkInvertGreen.Checked Then getGreen = Abs(255 - getGreen)
    If formLife.mChkInvertBlue.Checked Then getBlue = Abs(255 - getBlue)
End Sub

Function swapGenes(Genes As String) As String
    Dim G() As Byte
    Dim L As Long
    Dim i As Long
    Dim t
    Dim swap As Boolean
    swap = (Rnd < 0.5)
    L = Len(Genes)
    G() = Genes
    For i = 0 To L - 2 Step 2
        If Rnd > 0.95 Then swap = Not swap
'        If swap Then g(i) = (g(i) And 15) * 16 + (g(i) And 240) / 16
        If swap Then
            t = G(i + 1)
            G(i + 1) = G(i)
            G(i) = t
        End If
    Next i
    swapGenes = G()
End Function

Function breedGenes(ByVal dad As String, ByVal mom As String) As String
    dad = swapGenes(dad)
    mom = swapGenes(mom)
    breedGenes = combineGenes(dad, mom)
End Function

Function combineGenes(dad As String, mom As String) As String
    Dim gP() As Byte
    Dim gM() As Byte
    Dim G() As Byte
    Dim Lp As Long
    Dim Lm As Long
    Dim L As Long
    Dim i As Long
    Dim swap As Boolean
    If dad = "" Then dad = mom
    If mom = "" Then mom = dad
    swap = (Rnd < 0.5)
    Lp = Len(dad)
    gP() = dad
    Lm = Len(mom)
    gM() = mom
    G = gM
    L = lesserOf(Lp, Lm)
    For i = 0 To L - 2 Step 2
        G(i) = gM(i)
        G(i + 1) = gP(i + 1)
    Next i
    combineGenes = G()
End Function


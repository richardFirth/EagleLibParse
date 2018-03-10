VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EagleDeviceSetOb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'/T--EagleDeviceSetOb-------------------\
' Function Name   | Return    |  Description|
'-----------------|-----------|---------|
'setupDeviceSet   | String)   |  later  |
'getName          | String    |  later  |
'getData          | String()  |  later  |
'getPackagesUsed  | String()  |  later  |
'getSymbolsUsed   | String()  |  later  |
'~~getSymbolName  | String    |  later  |
'~~getPackageName | String    |  later  |
'\--------------------------------------/

Option Explicit

Private A_Name As String
Private B_Data() As String

Public Function setupDeviceSet(name As String, data() As String)
' later
    A_Name = name
    B_Data = data
End Function

Public Function getName() As String
' later
    getName = A_Name
End Function

Public Function getData() As String()
' later
    getData = B_Data
End Function

Public Function getPackagesUsed() As String()
' later
    Dim x As Integer
    Dim n As Integer: n = 1
    Dim thePackages() As String
    If Not arrayHasStuff(B_Data) Then Exit Function
    For x = LBound(B_Data) To UBound(B_Data)
        If Left(B_Data(x), 12) = "<device name" Then
            If Not "<device name="""">" = B_Data(x) Then
                ReDim Preserve thePackages(1 To n) As String
                thePackages(n) = getPackageName(B_Data(x))
                n = n + 1
            End If
        End If
    Next x
    getPackagesUsed = thePackages
End Function

Public Function getSymbolsUsed() As String()
' later
    Dim x As Integer
    Dim n As Integer: n = 1
    Dim theSymbols() As String
    If Not arrayHasStuff(B_Data) Then Exit Function
    For x = LBound(B_Data) To UBound(B_Data)
        If Left(B_Data(x), 10) = "<gate name" Then
            ReDim Preserve theSymbols(1 To n) As String
            theSymbols(n) = getSymbolName(B_Data(x))
            n = n + 1
        End If
    Next x
    getSymbolsUsed = theSymbols
End Function

Private Function getSymbolName(theSTR As String) As String
' later
    Dim SYM() As String
    SYM = Split(theSTR, """")
    getSymbolName = SYM(3)
End Function

Private Function getPackageName(theSTR As String) As String
' later
    Dim pkg() As String
    pkg = Split(theSTR, """")
    getPackageName = pkg(3)
End Function

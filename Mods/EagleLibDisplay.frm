VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} EagleLibDisplay 
   Caption         =   "Eagle Lib Display -"
   ClientHeight    =   6876
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   14592
   OleObjectBlob   =   "EagleLibDisplay.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "EagleLibDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'/T--EagleLibDisplay--------------------------\
' Function Name             | Return|  Description|
'---------------------------|-------|---------|
'initializeEagleLibDisplay  | Void  |  later  |
'~~CButton1_Click           | Void  |  later  |
'~~CommandButton1_Click     | Void  |  later  |
'~~deviceList_Click         | Void  |  later  |
'~~EagleLibDisplay_Click    | Void  |  later  |
'~~HighlightUnused_Click    | Void  |  later  |
'~~DelUnused_Click          | Void  |  later  |
'~~packageList_Click        | Void  |  later  |
'~~symbolList_Click         | Void  |  later  |
'~~TextBox1_Change          | Void  |  later  |
'~~UserForm_QueryClose      | Void  |  later  |
'\--------------------------------------------/

Option Explicit

Dim theEagleData As EagleLibObject

Sub initializeEagleLibDisplay(eDat As EagleLibObject)
' later
    Set theEagleData = eDat
    Me.Caption = "Display Library " & theEagleData.LibName
    Call PopulateListBoxWithStringArr(Me.deviceList, theEagleData.ZDeviceNames)
    Call PopulateListBoxWithStringArr(Me.packageList, theEagleData.ZPackageNames)
    Call PopulateListBoxWithStringArr(Me.symbolList, theEagleData.ZSymbolNames)
End Sub

Private Sub CButton1_Click()
' later
    Dim zzz As Integer
    For zzz = 0 To ListBox1.ListCount - 1
            If ListBox1.Selected(zzz) Then
            selectedOption = ListBox1.List(zzz)
            End If
    Next zzz
End Sub

Private Sub CommandButton1_Click()
' later
    If TextBox1.Value <> "" Then
        Dim fPath As String
        fPath = "C:\Users\rfirth1\Desktop\LibraryMacro\Output\" & TextBox1.Value & ".lbr"
        Call theEagleData.saveLibraryAsText(fPath)
        Me.Hide
    End If
End Sub

Private Sub deviceList_Click()
' later
    Dim theDevice As String
    theDevice = getSelectedItemsFromListBox(deviceList)(1)
    Dim tPKG() As String: tPKG = theEagleData.getPackagesUsedByDeviceName(theDevice)
    Dim tSym() As String: tSym = theEagleData.getSymbolsUsedByDeviceName(theDevice)
    packageList.MultiSelect = fmMultiSelectMulti
    symbolList.MultiSelect = fmMultiSelectMulti
    Call highlightSpecificItemsByArr(packageList, tPKG)
    Call highlightSpecificItemsByArr(symbolList, tSym)
   ' Call highlightSpecificItemsByArr
End Sub

Private Sub EagleLibDisplay_Click()
' later
End Sub

Private Sub HighlightUnused_Click()
' later
    Call deselectListBox(symbolList)
    Call deselectListBox(deviceList)
    Call deselectListBox(packageList)
    Call highlightSpecificItemsByArr(packageList, theEagleData.getUnusedPackages)
    Call highlightSpecificItemsByArr(symbolList, theEagleData.getUnusedSymbols)
End Sub

Private Sub DelUnused_Click()
' later
 Call theEagleData.removeByName(theEagleData.getUnusedSymbols, B_Symbol)
 Call theEagleData.removeByName(theEagleData.getUnusedPackages, A_Package)
 Call PopulateListBoxWithStringArr(Me.symbolList, theEagleData.ZSymbolNames)
 Call PopulateListBoxWithStringArr(Me.packageList, theEagleData.ZPackageNames)
End Sub

Private Sub packageList_Click()
' later
  '  Call deselectListBox(symbolList)
  '  Call deselectListBox(deviceList)
'
'    packageList.MultiSelect = fmMultiSelectSingle
'    symbolList.MultiSelect = fmMultiSelectSingle
End Sub

Private Sub symbolList_Click()
' later
 '   Call deselectListBox(packageList)
 '   Call deselectListBox(deviceList)
'
'   packageList.MultiSelect = fmMultiSelectSingle
'    symbolList.MultiSelect = fmMultiSelectSingle
End Sub

Private Sub TextBox1_Change()
' later
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
' later
    If CloseMode = 0 Then End
End Sub

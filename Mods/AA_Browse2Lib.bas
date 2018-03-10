Attribute VB_Name = "AA_Browse2Lib"

'/T--AA_Browse2Lib-------------------------------------\
' Function Name   | Return|  Description               |
'-----------------|-------|----------------------------|
'mainProcLib      | Void  |  main procedure            |
'GetAnEagleBoard  | Void  |  get an eagle board        |
'updateDislay     | Void  |  update the local display  |
'\-----------------------------------------------------/

Option Explicit

Dim theLib As New EagleLibObject

Sub mainProcLib()
' main procedure
    Dim theLBPths() As String
    theLBPths() = BrowseFilePaths(I_Lib)
    Dim tDATA() As String
    tDATA = TrimAndCleanArray(convertTXTDocumentToStringArr(theLBPths(1)))
Call theLib.initializeEagleLibFromLibrary(tDATA, nameFromPath(theLBPths(1)))
'theLib.displayEagleLibData
Dim NewEDisplay As New EagleLibDisplay
Call NewEDisplay.initializeEagleLibDisplay(theLib)
NewEDisplay.Show
'updateDislay
'Call theLib.populateLibrary("C:\Users\rfirth1\Desktop\Richard_Components2.lbr")
'Call showStrArr(theLib.getSymbolsUsedByDeviceName("RELAY"), "Symbols:")
End Sub

Sub GetAnEagleBoard()
' get an eagle board
    Dim brdPTH As String
    Dim schPTH As String
    schPTH = BrowseFilePath(K_SCH)
    brdPTH = Left(schPTH, Len(schPTH) - 3) & "brd"
   ' MsgBox schPTH & vbNewLine & brdPTH
    Dim brdDATA() As String:    brdDATA = TrimAndCleanArray(convertTXTDocumentToStringArr(brdPTH))
    Dim schDATA() As String:    schDATA = TrimAndCleanArray(convertTXTDocumentToStringArr(schPTH))
    Call theLib.initializeEagleLibFromBRD(brdDATA, schDATA, "BRD")
    Dim NewEDisplay As New EagleLibDisplay
    Call NewEDisplay.initializeEagleLibDisplay(theLib)
    NewEDisplay.Show
End Sub

Sub updateDislay()
' update the local display
    With ThisWorkbook.Sheets(1).Columns("A:G").ClearContents
        'Call printStringArrToColumn(theLib.origData, nShet, 1, "T Data")
        Call printStringArrToColumn(theLib.ZDeviceNames, ThisWorkbook.Sheets(1), 1, "Devices")
        Call printStringArrToColumn(theLib.ZPackageNames, ThisWorkbook.Sheets(1), 2, "Packages")
        Call printStringArrToColumn(theLib.ZSymbolNames, ThisWorkbook.Sheets(1), 3, "Symbols")
        Call printStringArrToColumn(theLib.getUnusedPackages, ThisWorkbook.Sheets(1), 4, "unused PKG")
        Call printStringArrToColumn(theLib.getUnusedSymbols, ThisWorkbook.Sheets(1), 5, "unused SYM")
    End With
End Sub


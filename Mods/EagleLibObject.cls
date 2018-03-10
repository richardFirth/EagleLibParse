VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EagleLibObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'/T--EagleLibObject-------------------------------------------------------------------------------------------\
' Function Name                  | Return            |  Description                                           |
'--------------------------------|-------------------|--------------------------------------------------------|
'initializeEagleLibFromLibrary   | Void              |  initialize the eagle library data from a library      |
'initializeEagleLibFromBRD       | Void              |  gets the eagle library gata from a board & schematic  |
'~~setEagleHeader                | Void              |  sets the eagle header                                 |
'getUnusedSymbols                | String()          |  gets the unused symbols from an eagle library         |
'getUnusedPackages               | String()          |  gets the unused packages from an eagle library        |
'getAllPackagesUsed              | String()          |  gets the packages used                                |
'getAllSymbolsUsed               | String()          |  gets the symbols used                                 |
'getSymbolsUsedByDeviceName      | String()          |  gets the symbols Used By a Device                     |
'getPackagesUsedByDeviceName     | String()          |  gets the packages Used By a Device                    |
'removeByName                    | Void              | im x As Integer                                        |
'~~removeDeviceByName            | Void              |  removes a device given its name                       |
'~~removeSymbolByName            | Void              |  removes a symbol given its name                       |
'~~removePackageByName           | Void              |  removes a package given its name                      |
'displayEagleLibData             | Void              |  display the data in a new sheet                       |
'-----get objects---------------------------------------------------------------------------------------------|
'~~getDeviceByName               | EagleDeviceSetOb  |  later                                                 |
'~~getSymbolByName               | EagleSymbolOb     |  later                                                 |
'~~getPackageByName              | EaglePackageOB    |  later                                                 |
'-----get lists of names--------------------------------------------------------------------------------------|
'ZDeviceNames                    | String()          |  later                                                 |
'ZPackageNames                   | String()          |  later                                                 |
'ZSymbolNames                    | String()          |  later                                                 |
'-----pop-----------------------------------------------------------------------------------------------------|
'saveLibraryAsText               | Void              | im newLibDat() As String                               |
'~~ZDeviceData                   | String()          |  later                                                 |
'~~ZPackageData                  | String()          |  later                                                 |
'~~ZSymbolData                   | String()          |  later                                                 |
'-----populate initial----------------------------------------------------------------------------------------|
'~~constructEagleDeviceSetObjects| Collection        |  later                                                 |
'~~constructEagleObjects         | Collection        |  later                                                 |
'~~constructEagleDeviceSet       | EagleDeviceSetOb  |  later                                                 |
'~~constructEaglePackageObject   | EaglePackageOB    |  later                                                 |
'~~constructEagleSymbolObject    | EagleSymbolOb     |  later                                                 |
'packageName                     | String            |  later                                                 |
'\------------------------------------------------------------------------------------------------------------/

Option Explicit

' http://www.patorjk.com/software/taag/#p=display&f=Soft&t=Type%20Something%20

Public Enum EagleTypes
    A_Package
    B_Symbol
    C_Device
End Enum

Public LibName As String

Private origData() As String
Private newData() As String

Private headerDat() As String
Private eaglePackages As Collection
Private eagleSymbols As Collection
Private eagleDevices As Collection

Public Sub initializeEagleLibFromLibrary(theLibData() As String, setlname As String)
' initialize the eagle library data from a library
    origData = theLibData
    LibName = setlname
    Call setEagleHeader(theLibData)
    Set eaglePackages = constructEagleObjects(theLibData, A_Package)
    Set eagleSymbols = constructEagleObjects(theLibData, B_Symbol)
    Set eagleDevices = constructEagleDeviceSetObjects(theLibData)
End Sub

Public Sub initializeEagleLibFromBRD(theBRD() As String, theSCH() As String, setlname As String)
' gets the eagle library gata from a board & schematic
   ' origData = theLibData
    LibName = setlname
    Call setEagleHeader(getArrayFromColumn(ThisWorkbook.Sheets("LibHeader"), 1))
    Set eaglePackages = constructEagleObjects(theBRD, A_Package)
    Set eagleSymbols = constructEagleObjects(theSCH, B_Symbol)
    Set eagleDevices = constructEagleDeviceSetObjects(theSCH)
End Sub

Private Sub setEagleHeader(hData() As String)
' sets the eagle header
    Dim locH() As String
    Dim n As Integer: n = 1
    Dim x As Integer
    For x = LBound(hData) To UBound(hData)
        ReDim Preserve locH(1 To n) As String
        locH(n) = hData(x)
        n = n + 1
        If hData(x) = "</layers>" Then Exit For
    Next x
    headerDat = locH
End Sub

Public Function getUnusedSymbols() As String()
' gets the unused symbols from an eagle library
    getUnusedSymbols = DifferenceBetweenSets(ZSymbolNames, getAllSymbolsUsed)
End Function

Public Function getUnusedPackages() As String()
' gets the unused packages from an eagle library
    getUnusedPackages = DifferenceBetweenSets(ZPackageNames, getAllPackagesUsed)
End Function

Public Function getAllPackagesUsed() As String()
' gets the packages used
    Dim pkgUsed() As String
    Dim theDevices() As String: theDevices = ZDeviceNames
    If Not arrayHasStuff(theDevices) Then Exit Function
    Dim x As Integer
    For x = LBound(theDevices) To UBound(theDevices)
        pkgUsed = ConcatenateArrays(pkgUsed, getPackagesUsedByDeviceName(theDevices(x)))
    Next x
    getAllPackagesUsed = removeDupesStringArray(pkgUsed)
End Function

Public Function getAllSymbolsUsed() As String()
' gets the symbols used
    Dim symUsed() As String
    Dim theDevices() As String: theDevices = ZDeviceNames
    If Not arrayHasStuff(theDevices) Then Exit Function
    Dim x As Integer
    For x = LBound(theDevices) To UBound(theDevices)
        symUsed = ConcatenateArrays(symUsed, getSymbolsUsedByDeviceName(theDevices(x)))
    Next x
    getAllSymbolsUsed = removeDupesStringArray(symUsed)
End Function

Public Function getSymbolsUsedByDeviceName(theDeviceName As String) As String()
' gets the symbols Used By a Device
    Dim theDev As EagleDeviceSetOb
    Set theDev = getDeviceByName(theDeviceName)
    getSymbolsUsedByDeviceName = theDev.getSymbolsUsed
End Function

Public Function getPackagesUsedByDeviceName(theDeviceName As String) As String()
' gets the packages Used By a Device
    Dim theDev As EagleDeviceSetOb
    Set theDev = getDeviceByName(theDeviceName)
    getPackagesUsedByDeviceName = theDev.getPackagesUsed
End Function

Public Sub removeByName(theNames() As String, tType As EagleTypes)
Dim x As Integer
If Not arrayHasStuff(theNames) Then Exit Sub
For x = LBound(theNames) To UBound(theNames)
    If tType = A_Package Then Call removePackageByName(theNames(x))
    If tType = B_Symbol Then Call removeSymbolByName(theNames(x))
    If tType = C_Device Then Call removeDeviceByName(theNames(x))
Next x
End Sub

Private Sub removeDeviceByName(theDeviceName As String)
' removes a device given its name
    Dim x As Integer
    For x = 1 To eagleSymbols.Count
        If eagleDevices.Item(x).getName = theDeviceName Then Call eagleDevices.Remove(x): Exit Sub
    Next x
End Sub

Private Sub removeSymbolByName(theSymbolName As String)
' removes a symbol given its name
    Dim x As Integer
    For x = 1 To eagleSymbols.Count
        If eagleSymbols.Item(x).getName = theSymbolName Then Call eagleSymbols.Remove(x): Exit Sub
    Next x
End Sub

Private Sub removePackageByName(thePackageName As String)
' removes a package given its name
    Dim x As Integer
    For x = 1 To eaglePackages.Count
        If eaglePackages.Item(x).getName = thePackageName Then Call eaglePackages.Remove(x): Exit Sub
    Next x
End Sub

'Private eaglePackages As Collection
'Private eagleSymbols As Collection
'Private eagleDevices As Collection

Public Sub displayEagleLibData()
' display the data in a new sheet
    Dim xWKBK As Workbook: Set xWKBK = Workbooks.Add
    Dim nShet As Worksheet: Set nShet = xWKBK.Sheets(1)
    Call printStringArrToColumn(origData, nShet, 1, "T Data")
    Call printStringArrToColumn(ZDeviceNames, nShet, 2, "Devices")
    Call printStringArrToColumn(ZPackageNames, nShet, 3, "Packages")
    Call printStringArrToColumn(ZSymbolNames, nShet, 4, "Symbols")
End Sub

'
'    ___     _       ___ _     _           _       .
'   / _ \___| |_    /___\ |__ (_) ___  ___| |_ ___
'  / /_\/ _ \ __|  //  // '_ \| |/ _ \/ __| __/ __|
' / /_\\  __/ |_  / \_//| |_) | |  __/ (__| |_\__ \
' \____/\___|\__| \___/ |_.__// |\___|\___|\__|___/
'                            |__/

'#get objects

Private Function getDeviceByName(tname As String) As EagleDeviceSetOb
' later
    Dim aDevice As EagleDeviceSetOb
    For Each aDevice In eagleDevices
        If aDevice.getName = tname Then Set getDeviceByName = aDevice: Exit Function
    Next aDevice
End Function

Private Function getSymbolByName(tname As String) As EagleSymbolOb
' later
    Dim aSym As EagleSymbolOb
    For Each aSym In eagleSymbols
        If aSym.getName = tname Then Set getSymbolByName = aSym: Exit Function
    Next aSym
End Function

Private Function getPackageByName(tname As String) As EaglePackageOB
' later
    Dim aPkg As EaglePackageOB
    For Each aPkg In eagleSymbols
        If aPkg.getName = tname Then Set getPackageByName = aPkg: Exit Function
    Next aPkg
End Function

'            _     _ _     _                __
'  __ _  ___| |_  | (_)___| |_ ___    ___  / _|  _ __   __ _ _ __ ___   ___  ___
' / _` |/ _ \ __| | | / __| __/ __|  / _ \| |_  | '_ \ / _` | '_ ` _ \ / _ \/ __|
'| (_| |  __/ |_  | | \__ \ |_\__ \ | (_) |  _| | | | | (_| | | | | | |  __/\__ \
' \__, |\___|\__| |_|_|___/\__|___/  \___/|_|   |_| |_|\__,_|_| |_| |_|\___||___/
' |___/
'
'#get lists of names

Public Function ZDeviceNames() As String()
' later
    Dim dNames() As String
    Dim aDEV As EagleDeviceSetOb
    Dim n As Integer: n = 1
    For Each aDEV In eagleDevices
        ReDim Preserve dNames(1 To n) As String
        dNames(n) = aDEV.getName
        n = n + 1
    Next aDEV
    ZDeviceNames = dNames
End Function

Public Function ZPackageNames() As String()
' later
    Dim pNames() As String
    Dim aPkg As EaglePackageOB
    Dim n As Integer: n = 1
    For Each aPkg In eaglePackages
        ReDim Preserve pNames(1 To n) As String
        pNames(n) = aPkg.getName
        n = n + 1
    Next aPkg
    ZPackageNames = pNames
End Function

Public Function ZSymbolNames() As String()
' later
    Dim sNames() As String
    Dim aSym As EagleSymbolOb
    Dim n As Integer: n = 1
    For Each aSym In eagleSymbols
        ReDim Preserve sNames(1 To n) As String
        sNames(n) = aSym.getName
        n = n + 1
    Next aSym
    ZSymbolNames = sNames
End Function

'.
'   ___                  _       _         _          _            _  .
'  / _ \___  _ __  _   _| | __ _| |_ ___  | |_ ___   | |_ _____  _| |_
' / /_)/ _ \| '_ \| | | | |/ _` | __/ _ \ | __/ _ \  | __/ _ \ \/ / __|
'/ ___/ (_) | |_) | |_| | | (_| | ||  __/ | || (_) | | ||  __/>  <| |_
'\/    \___/| .__/ \__,_|_|\__,_|\__\___|  \__\___/   \__\___/_/\_\\__|
'           |_|
'

'#pop
Public Sub saveLibraryAsText(saveTo As String)
    Dim newLibDat() As String
    newLibDat = ConcatenateArrays(newLibDat, headerDat)
        newLibDat = AddToStringArray(newLibDat, "<library>")
    newLibDat = ConcatenateArrays(newLibDat, ZPackageData)
    newLibDat = ConcatenateArrays(newLibDat, ZSymbolData)
    newLibDat = ConcatenateArrays(newLibDat, ZDeviceData)
    Dim dSTR() As String
    dSTR = Split("</library>,</drawing>,</eagle>", ",")
    newLibDat = ConcatenateArrays(newLibDat, dSTR)
    Call createTextFromStringArr(newLibDat, saveTo)
End Sub

    
    Private Function ZDeviceData() As String()
    ' later
        Dim dData() As String
        Dim aDEV As EagleDeviceSetOb
        dData = AddToStringArray(dData, "<devicesets>")
        For Each aDEV In eagleDevices
            dData = ConcatenateArrays(dData, aDEV.getData)
        Next aDEV
        dData = AddToStringArray(dData, "</devicesets>")
        ZDeviceData = dData
    End Function

    Private Function ZPackageData() As String()
    ' later
        Dim pData() As String
        Dim aPkg As EaglePackageOB
        pData = AddToStringArray(pData, "<packages>")
        For Each aPkg In eaglePackages
            pData = ConcatenateArrays(pData, aPkg.getData)
        Next aPkg
        pData = AddToStringArray(pData, "</packages>")
        ZPackageData = pData
    End Function

    Private Function ZSymbolData() As String()
    ' later
        Dim sData() As String
        Dim aSym As EagleSymbolOb
        sData = AddToStringArray(sData, "<symbols>")
        For Each aSym In eagleSymbols
            sData = ConcatenateArrays(sData, aSym.getData)
        Next aSym
        sData = AddToStringArray(sData, "</symbols>")
        ZSymbolData = sData
    End Function

'
'  _____                  _       _         _____       _ _   _       _      '
' |  __ \                | |     | |       |_   _|     (_) | (_)     | |
' | |__) |__  _ __  _   _| | __ _| |_ ___    | |  _ __  _| |_ _  __ _| |
' |  ___/ _ \| '_ \| | | | |/ _` | __/ _ \   | | | '_ \| | __| |/ _` | |
' | |  | (_) | |_) | |_| | | (_| | ||  __/  _| |_| | | | | |_| | (_| | |
' |_|   \___/| .__/ \__,_|_|\__,_|\__\___| |_____|_| |_|_|\__|_|\__,_|_|
'            | |
'            |_|
'
'#populate initial

Private Function constructEagleDeviceSetObjects(theLBR() As String) As Collection
' later
    Dim x As Long
    Dim theDeviceSet As Collection
    Set theDeviceSet = New Collection
        For x = LBound(theLBR) To UBound(theLBR)
            If Left(theLBR(x), 15) = "<deviceset name" Then
                Call theDeviceSet.Add(constructEagleDeviceSet(theLBR, x))
            End If
        Next x
    Set constructEagleDeviceSetObjects = theDeviceSet
End Function

Private Function constructEagleObjects(theLBR() As String, tType As EagleTypes) As Collection
' later
    Dim x As Long
    Dim thePKGS As Collection
    Set thePKGS = New Collection
        If tType = A_Package Then
            For x = LBound(theLBR) To UBound(theLBR)
                If Left(theLBR(x), 13) = "<package name" Then
                    Call thePKGS.Add(constructEaglePackageObject(theLBR, x))
                End If
            Next x
        End If
        If tType = B_Symbol Then
            For x = LBound(theLBR) To UBound(theLBR)
                If Left(theLBR(x), 12) = "<symbol name" Then
                    Call thePKGS.Add(constructEagleSymbolObject(theLBR, x))
                End If
            Next x
        End If
    Set constructEagleObjects = thePKGS
End Function

Private Function constructEagleDeviceSet(theLib() As String, theNum As Long) As EagleDeviceSetOb
' later
    Dim theEagDevSet As New EagleDeviceSetOb
    Dim theData() As String
    Dim y As Integer
    Dim n As Integer: n = 1
    For y = theNum To UBound(theLib)
        ReDim Preserve theData(1 To n) As String
        theData(n) = theLib(y)
        n = n + 1
        If theLib(y) = "</deviceset>" Then Exit For
    Next y
    'setupDeviceSet
   Call theEagDevSet.setupDeviceSet(packageName(theLib(theNum)), theData)
   Set constructEagleDeviceSet = theEagDevSet
End Function

Private Function constructEaglePackageObject(theLib() As String, theNum As Long) As EaglePackageOB
' later
    Dim theEagPKG As New EaglePackageOB
    Dim theData() As String
    Dim y As Integer
    Dim n As Integer: n = 1
    For y = theNum To UBound(theLib)
        ReDim Preserve theData(1 To n) As String
        theData(n) = theLib(y)
        n = n + 1
        If theLib(y) = "</package>" Then Exit For
    Next y
   Call theEagPKG.setupPackage(packageName(theLib(theNum)), theData)
   Set constructEaglePackageObject = theEagPKG
End Function

Private Function constructEagleSymbolObject(theLib() As String, theNum As Long) As EagleSymbolOb
' later
    Dim theEagPKG As New EagleSymbolOb
    Dim theData() As String
    Dim y As Integer
    Dim n As Integer: n = 1
    For y = theNum To UBound(theLib)
        ReDim Preserve theData(1 To n) As String
        theData(n) = theLib(y)
        n = n + 1
        If theLib(y) = "</symbol>" Then Exit For
    Next y
   Call theEagPKG.setupSymbol(packageName(theLib(theNum)), theData)
   Set constructEagleSymbolObject = theEagPKG
End Function

Function packageName(thePKGNM As String) As String
' later
    Dim thePName() As String
    thePName = Split(thePKGNM, """")
    packageName = thePName(1)
End Function


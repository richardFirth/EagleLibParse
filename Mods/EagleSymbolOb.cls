VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EagleSymbolOb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'/T--EagleSymbolOb-------------------------------------\
' Function Name| Return    |  Description               |
'-------------|-----------|----------------------------|
'setupSymbol  | String)   |  set up eagle symbol data  |
'getName      | String    |  gets the symbol name      |
'getData      | String()  |  gets the symbol data      |
'\-----------------------------------------------------/

Option Explicit

Private A_Name As String
Private B_Data() As String

Public Function setupSymbol(name As String, data() As String)
' set up eagle symbol data
    A_Name = name
    B_Data = data
End Function

Public Function getName() As String
' gets the symbol name
    getName = A_Name
End Function

Public Function getData() As String()
' gets the symbol data
    getData = B_Data
End Function


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EaglePackageOB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'/T--EaglePackageOB----------------------------------\
' Function Name| Return    |  Description            |
'--------------|-----------|-------------------------|
'setupPackage  | String)   |  gets the package data  |
'getName       | String    |  gets the name data     |
'getData       | String()  |  gets the name data     |
'\---------------------------------------------------/

Option Explicit

Private A_Name As String
Private B_Data() As String

Public Function setupPackage(name As String, data() As String)
' gets the package data
    A_Name = name
    B_Data = data
End Function

Public Function getName() As String
' gets the name data
    getName = A_Name
End Function

Public Function getData() As String()
' gets the name data
    getData = B_Data
End Function


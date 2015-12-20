Attribute VB_Name = "Module1"
Public Declare Function Inp Lib "inpout32.dll" Alias "Inp32" (ByVal PortAddress As Integer) As Integer
Public Declare Sub Out Lib "inpout32.dll" Alias "Out32" (ByVal PortAddress As Integer, ByVal Value As Integer)


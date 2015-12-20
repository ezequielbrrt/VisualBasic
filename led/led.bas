Attribute VB_Name = "Module1"
Public Declare Function Inp Lib "C:\INPOUT32.DLL" Alias "Inp32" (ByVal Portaddress As Integer) As Integer
Public Declare Sub Out Lib "C:\INPOUT32.DLL" Alias "Out32" (ByVal Portaddress As Integer, ByVal Value As Integer)










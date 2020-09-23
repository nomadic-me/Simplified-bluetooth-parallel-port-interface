Attribute VB_Name = "Module1"
Public Declare Sub PortOut Lib "io.dll" (ByVal Port As Integer, ByVal Value As Byte)
Public Declare Function PortIn Lib "io.dll" (ByVal Port As Integer) As Byte



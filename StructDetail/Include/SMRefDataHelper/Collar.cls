VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Collar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Const MODULE As String = "S:\StructDetail\Data\Include\SMRefDataHelper\Collar.cls"

'Function to find if object is a collar or not
Public Function IsCollar(oObject As Object) As Boolean

    On Error GoTo ErrorHandler
    
    IsCollar = False
    
    If oObject Is Nothing Then Exit Function
    
    If TypeOf oObject Is IJPlate Then
        Dim oPlateUtil As IJPlateAttributes
        Set oPlateUtil = New PlateUtils
    
        Dim plateType As StructPlateType
    
        Dim oPlate As IJPlate
        Set oPlate = oObject
        plateType = oPlate.plateType
    
        If TypeOf oObject Is IJPlatePart And plateType = CollarPlate Then
            IsCollar = True
        End If
    End If

    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "IsCollar").Number

End Function
' ** End CM **


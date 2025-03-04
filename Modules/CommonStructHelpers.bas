Attribute VB_Name = "CommonStructHelpers"
Option Explicit
' depends on StructGenericTools (IJStructSymbolTools)
' depends on StructGeneric (IJStructOperationPattern)

Private m_pStructSymbolTools As IJStructSymbolTools
Private m_oAcisHelper As AcisHelper
Private m_oPortHelper As PortHelper
Function GetStructSymbolTools() As IJStructSymbolTools
    If m_pStructSymbolTools Is Nothing Then
        Set m_pStructSymbolTools = New StructSymbolTools
    End If
    Set GetStructSymbolTools = m_pStructSymbolTools
End Function
Public Function GetAcisHelper() As AcisHelper
    If m_oAcisHelper Is Nothing Then
        Set m_oAcisHelper = New AcisHelper
    End If
    Set GetAcisHelper = m_oAcisHelper
End Function
Public Function GetPortHelper() As PortHelper
    If m_oPortHelper Is Nothing Then
        Set m_oPortHelper = New PortHelper
    End If
    Set GetPortHelper = m_oPortHelper
End Function

Function StructEntity_GetOperation(pStructOperationPattern As IJStructOperationPattern, sProgid As String)
    Dim pElementsOfOperators As IJElements
    Dim oOperation As Object
    Call pStructOperationPattern.GetOperationPattern(sProgid, pElementsOfOperators, oOperation)
    Set StructEntity_GetOperation = oOperation
End Function


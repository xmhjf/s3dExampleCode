Attribute VB_Name = "GCErrors"
Option Explicit
Const E_FAIL = &H80004005
Public Sub GCEvaluate(pGeometricConstructionIn As GeometricConstruction, pGeometricConstructionOut As GeometricConstruction, Optional lErrorNumberIn As Long = -1, Optional lErrorNumberOut As Long = -1)
    On Error Resume Next
    Call pGeometricConstructionIn.Evaluate
    If Err.Number <> 0 Then
        On Error GoTo 0
        Dim pToDoListHelper As IJToDoListHelper
        Set pToDoListHelper = pGeometricConstructionIn
        
        Dim lErrorNumber As Long
        Dim sCodeListTable As String
        Call pToDoListHelper.GetErrorInfo(sCodeListTable, lErrorNumber)
        If lErrorNumberOut <> -1 _
        And (lErrorNumber = lErrorNumberIn Or lErrorNumberIn = -1) Then
            ' transform an expected error
            Call GCProcessError(pGeometricConstructionOut, pGeometricConstructionOut.TypeName + "_Errors", lErrorNumberOut)
        ElseIf lErrorNumberIn = -1 Then
            ' forward any error
            Call GCProcessError(pGeometricConstructionOut, sCodeListTable, lErrorNumber)
        Else
            ' forward unexpected error
            Call GCProcessError(pGeometricConstructionOut, sCodeListTable, lErrorNumber)
        End If
    End If
    On Error GoTo 0
End Sub
Public Sub GCProcessError(pGeometricConstruction As IJGeometricConstruction, Optional sCodeListTable As String = "", Optional lErrorNumber As Long = 0)
    Dim pToDoListHelper As IJToDoListHelper
    Set pToDoListHelper = pGeometricConstruction
    ' MsgBox "GCProcessError: Type= " + pGeometricConstruction.Type + ", CodeListTable= " + sCodeListTable + ", ErrorNumber= " + CStr(lErrorNumber)
    
    ' try to forward the error message to the APS
    Dim vObjectToUpdate As Variant:
    If True Then
        Dim oCollectionOfNames As Collection
        Set oCollectionOfNames = GeometricConstruction_GetNamesOfControlledInputs(pGeometricConstruction)
        If Collection_ContainsName(oCollectionOfNames, "AdvancedPlateSystem") Then
            If pGeometricConstruction.ControlledInputs("AdvancedPlateSystem").Count = 1 Then
                Dim oObjectToUpdate As Object
                Set oObjectToUpdate = pGeometricConstruction.ControlledInputs("AdvancedPlateSystem")(1)
                Set vObjectToUpdate = oObjectToUpdate
            End If
        End If
    End If
    
    If sCodeListTable = "" Then
        Call pToDoListHelper.SetErrorInfo(pGeometricConstruction.TypeName + "_Errors", lErrorNumber, vObjectToUpdate)
    Else
        Call pToDoListHelper.SetErrorInfo(sCodeListTable, lErrorNumber, vObjectToUpdate)
    End If
    Err.Raise E_FAIL
End Sub
Sub ShowError(sModule As String, sComment As String, iNumber As Integer, sDescription As String)
    Debug.Print "Module: " + sModule + ", Comment= " + sComment + ", iNumber: " + CStr(iNumber) + ", Description: " + sDescription
End Sub
Private Function GeometricConstruction_GetNamesOfControlledInputs(pGeometricConstruction As IJGeometricConstruction) As Collection
    ' prepare a collection to be filled
    Dim pCollection As New Collection
    
    ' fill the collection with the names of the controlled inputs
    If True Then
        ' retrieve the definition
        Dim pGeometricConstructionDefinition As IJGeometricConstructionDefinition: Set pGeometricConstructionDefinition = pGeometricConstruction.definition
        
        ' loop on the controlled inputs
        Dim i As Integer
        For i = 1 To pGeometricConstructionDefinition.ControlledInputCount
            ' retrieve the name of the controlled input
            Dim sName As String
            If True Then
                Dim sComputeIIDs As String
                
                ' retrieve the controlled input
                Call pGeometricConstructionDefinition.GetControlledInputInfoByIndex(i, sName, sComputeIIDs)
            End If
            
            ' add the name to the collection
            Call pCollection.Add(sName)
        Next
    End If
    
    ' return result
    Set GeometricConstruction_GetNamesOfControlledInputs = pCollection
End Function
Private Function Collection_ContainsName(pCollection As Collection, sName As String) As Boolean
    'initialize result
    Let Collection_ContainsName = False
    
    ' loop on the collection
    Dim i As Integer
    For i = 1 To pCollection.Count
        ' retrieve name of the current item
        Dim sNameOfItem As String: Let sNameOfItem = pCollection.Item(i)
        
        ' compare item name to searched name
        If sName = sNameOfItem Then
            Let Collection_ContainsName = True
            Exit For
        End If
    Next
End Function



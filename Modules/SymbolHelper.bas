Attribute VB_Name = "SymbolHelper"
Option Explicit

'*******************************************************************
'
'Copyright (C) 2003 Intergraph Corporation. All rights reserved.
'
'File : StructHelper.bas
'
'Author : VG
'
'Description :
'    Symbol machinery helper
'
''********************************************************************


'*************************************************************************
'Function
'CreateNestedSymbol
'
'Abstract
' Create a nested symbol
'
'input
'error object, file name where error occurred, method name, data source, line number
'
'Return
'error code
'
'Exceptions
'
'***************************************************************************
Public Sub CreateNestedSymbol(NestedSymbDefName As String, _
                                ByRef EnumInputArg As IJDEnumArgument, _
                                ByRef oOutputColl As Object, _
                                OutputName As String)
On Error GoTo ErrorHandler
    ' Get the symbol factory
    Dim oSymbolFactory As New IMSSymbolEntities.DSymbolEntitiesFactory
    
    ' Get or create the symbol definition
    Dim oNestedSymDef As IJDSymbolDefinition
    Set oNestedSymDef = GetSymbolDefinition(oSymbolFactory, oOutputColl, NestedSymbDefName)

    ' Place the nested symbol
    Dim oNestedSymbol As IMSSymbolEntities.DSymbol
    Set oNestedSymbol = PlaceSymbol(oSymbolFactory, oNestedSymDef, EnumInputArg, oOutputColl)
    
    ' Update the symbol and connect to outputcollection
    Call UpdateSymbol(oNestedSymbol, oOutputColl, OutputName)
    
    Exit Sub
    
ErrorHandler:
    ' Raise the error
    Err.Raise Err.Number
End Sub

Private Function GetSymbolDefinition(ByRef oSymbolFactory As IMSSymbolEntities.DSymbolEntitiesFactory, _
                                        ByRef oOutputColl As Object, _
                                        NestedSymbDefName As String) As IJDSymbolDefinition
On Error GoTo ErrorHandler

    ' Get the definition collection
    Dim defColl As IJDDefinitionCollection
    Set defColl = oSymbolFactory.DefinitionCollection(oOutputColl.ResourceManager)
    
    ' Get or create the symbol definition
    Dim definitionParams As Variant: definitionParams = ""
    
    Set GetSymbolDefinition = defColl.GetDefinitionByProgId(True, NestedSymbDefName, vbNullString, definitionParams)
    Exit Function

ErrorHandler:
    Err.Raise Err.Number
    
End Function

Private Function PlaceSymbol(ByRef oSymbolFactory As IMSSymbolEntities.DSymbolEntitiesFactory, _
                                        ByRef oNestedSymDef As IJDSymbolDefinition, _
                                        ByRef EnumInputArg As IJDEnumArgument, _
                                        ByRef oOutputColl As Object) As IMSSymbolEntities.DSymbol
On Error GoTo ErrorHandler

    ' Place the nested symbol
    Dim oNestedSymbol As IMSSymbolEntities.DSymbol
    Set oNestedSymbol = oSymbolFactory.PlaceSymbol(oNestedSymDef, oOutputColl.ResourceManager)
    
    ' Add the input
    oNestedSymbol.IJDValuesArg.SetValues EnumInputArg

    ' Return the symbol
    Set PlaceSymbol = oNestedSymbol

    Exit Function

ErrorHandler:
    ' Delete the nested symbol
    If Not oNestedSymbol Is Nothing Then
        Dim oEntity As IMSEntitySupport.IJDObject
        Set oEntity = oNestedSymbol
        oEntity.Remove
    End If
    
    ' Raise the error
    Err.Raise Err.Number
End Function

Private Sub UpdateSymbol(ByRef oSymbol As IMSSymbolEntities.DSymbol, _
                            ByRef oOutputColl As Object, _
                            OutputName As String)
On Error GoTo ErrorHandler

    ' Update
    Dim IJDInputsArg As IMSSymbolEntities.IJDInputsArg
    Set IJDInputsArg = oSymbol
    
    IJDInputsArg.Update
    
    Set IJDInputsArg = Nothing
    
    ' Add the nested symbol to the output collection
    oOutputColl.AddOutput "Support1", oSymbol
    
    Exit Sub

ErrorHandler:
    ' Attach the nested symbol to the output collection
    oOutputColl.AddOutput "Support1", oSymbol
    
    ' Raise the error
    Err.Raise Err.Number
End Sub

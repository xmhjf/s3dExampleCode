VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GussetPlateType1Sym"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'---------------------------------------------------------------------------
'    Copyright (C) 2003 - 2004 Intergraph Corporation. All rights reserved.
'
'
'Abstract
'    USS defining the shape of the parametric plate part
'
'
'Notes
'The plate contour is created at using a local coordinate system with origin at top left of the plate.
' the plate is cut at the bottom right corner
'History
'    Raj Peringoth      05/07/03      Creation
'---------------------------------------------------------------------------------------
Option Explicit
Private Const MODULE = "GussetPlateType1Sym"
Private Const SYMBOLDEFNAME = "SPSPartMacros.GussetPlateType1Sym"


Implements IJDUserSymbolServices


Private Sub IJDUserSymbolServices_InvokeRepresentation(ByVal pSymbolOccurrence As Object, ByVal pRepName As String, ByVal pOutputColl As Object, arrayOfInputs() As Variant)

End Sub

Private Function IJDUserSymbolServices_EditOccurence(pSymbolOccurrence As Object, ByVal pTransactionMgr As Object) As Boolean

End Function

Private Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
Const MT = "IJDUserSymbolServices_GetDefinitionName"
On Error GoTo ErrorHandler
    Dim insertTypeString As String
     Dim insertTypeVal As IMSInsertionType
     Dim serverName As String
    

     IJDUserSymbolServices_GetDefinitionName = SYMBOLDEFNAME
Exit Function
ErrorHandler: HandleError MODULE, MT
End Function

Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(pSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition)
    Const MT = "IJDUserSymbolServices_InitializeSymbolDefinition"
On Error GoTo ErrorHandler
    
    
    'Define the inputs
    
    Dim pIJDInputs As IJDInputs
    Dim oInput As IJDInput

    
    Set pIJDInputs = pSymbolDefinition
    
    Set oInput = New DInput
    
    oInput.Name = "SupportedPort"
    oInput.Properties = igDESCRIPTION_OPTIONAL
    oInput.index = 1
    pIJDInputs.Add oInput
    oInput.Reset
    
    oInput.Name = "Supporting1Port"
    oInput.Properties = igDESCRIPTION_OPTIONAL
    oInput.index = 2
    pIJDInputs.Add oInput
    oInput.Reset
    
    oInput.Name = "Supporting2Port"
    oInput.Properties = igDESCRIPTION_OPTIONAL
    oInput.index = 3
    pIJDInputs.Add oInput
    oInput.Reset
    
    oInput.Name = "CutbackSO"
    oInput.Properties = igDESCRIPTION_OPTIONAL
    oInput.index = 4
    pIJDInputs.Add oInput
    oInput.Reset
    
    
    
    oInput.Name = "Length"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oInput.index = 5
    pIJDInputs.Add oInput
    
    oInput.Reset
    oInput.Name = "Width"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oInput.index = 6
    pIJDInputs.Add oInput
    
    Set oInput = New DInput
    oInput.Name = "CutLength"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oInput.index = 7
    pIJDInputs.Add oInput
    
    oInput.Reset
    oInput.Name = "CutWidth"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oInput.index = 8
    pIJDInputs.Add oInput
    
    
    'Defines the output of this symbol
    Dim TheReps As IJDRepresentations
    Dim oRep As IJDRepresentation
    
    Set TheReps = pSymbolDefinition
    Set oRep = New DRepresentation
    oRep.Name = "ParametricWireBodyShape"
    SetCMOnRepresentation oRep, "StdLib", "CreateParametricWireBodyShape", pSymbolDefinition
    
    Dim TheOutputs As IJDOutputs
    Dim oOutput As IJDOutput
    
    Set TheOutputs = oRep
    
    Set oOutput = New DOutput
    oOutput.Name = "WireBodyShape"
    
    TheOutputs.Add oOutput
    
    TheReps.Add oRep
    'as this symbol def has declared a graphic object as input
    ' GeomOption option will be set to igSYMBOL_GEOM_FIX_TO_ID by the symbol machinerary
    'Because of this the  outputs will be transformed during MDR and the Symbol geometry will
    ' end up in an incorrect location. So resetting the flag - DI226263
    pSymbolDefinition.GeomOption = igSYMBOL_GEOM_FREE
    
    Exit Sub

ErrorHandler: HandleError MODULE, MT
End Sub

Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, ByVal definitionParameters As Variant, ByVal pResourceMgr As Object) As Object
Const MT = "IJDUserSymbolServices_InstanciateDefinition"
On Error GoTo ErrorHandler

    Dim pSymbolDefinition As IJDSymbolDefinition
    Dim oSymbolFactory As New IMSSymbolEntities.DSymbolEntitiesFactory
    
    Set pSymbolDefinition = oSymbolFactory.CreateEntity(definition, pResourceMgr)
    Set oSymbolFactory = Nothing
    
    Dim ProgId As String
    Dim insertType As String
    Dim insertTypeVal As IMSInsertionType
    Dim serverName As String
    
    pSymbolDefinition.ProgId = SYMBOLDEFNAME
    pSymbolDefinition.CodeBase = CodeBase
    pSymbolDefinition.Name = IJDUserSymbolServices_GetDefinitionName(definitionParameters)
    pSymbolDefinition.MetaDataOption = igSYMBOL_DYNAMIC_METADATA
    
'''' support only -> go away when no occurences
    pSymbolDefinition.SupportOnlyOption = igSYMBOL_SUPPORT_ONLY
    
    Dim UserMethod As IJDUserMethods
    Dim LibDesc As New DLibraryDescription
    Dim MethodDesc As New DMethodDescription
    
    Set UserMethod = pSymbolDefinition
    

    LibDesc.Name = "StdLib"
    LibDesc.Type = imsLIBRARY_IS_ACTIVEX
    LibDesc.Properties = imsLIBRARY_AUTO_EXTRACT_METHOD_COOKIES
    LibDesc.Source = SYMBOLDEFNAME
    'Set library in the symbol definition
    UserMethod.SetLibrary LibDesc


    MethodDesc.Name = "CreateParametricWireBodyShape"
    MethodDesc.Properties = 0
    MethodDesc.Cookie = imsCOOKIE_IS_NULL
    MethodDesc.Library = "StdLib"
    UserMethod.SetMethod MethodDesc

    Set MethodDesc = Nothing
    Set LibDesc = Nothing
    Set UserMethod = Nothing
     
    IJDUserSymbolServices_InitializeSymbolDefinition pSymbolDefinition
    
    'returned symbol defintion
    Set IJDUserSymbolServices_InstanciateDefinition = pSymbolDefinition

    Exit Function

ErrorHandler: HandleError MODULE, MT

End Function

Public Sub CreateParametricWireBodyShape(pRepSCM As IJDRepresentationStdCustomMethod)
Const MT = "CreateParametricWireBodyShape"
On Error GoTo ErrorHandler
 Dim oREpDG As IMSSymbolEntities.IJDRepresentationDuringGame
 Dim pSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition
 Dim pOutputCollection As IMSSymbolEntities.IJDOutputCollection
 Dim oOutputs As IJDOutputs
 Dim GraphicOutput As IMSSymbolEntities.IJDOutput
 Dim DefPlayerEx As IMSSymbolEntities.IJDDefinitionPlayerEx
 Dim oSymbolOcc As IJDSymbol
 Dim pEnumJDArgument As IEnumJDArgument
 Dim arg As IJDArgument
 Dim found As Long
 Dim Length As Double, Width As Double, cutLength As Double, cutWidth As Double
 Dim PC As IJDParameterContent



 Set oREpDG = pRepSCM
 Set pSymbolDefinition = oREpDG.definition
 Set pOutputCollection = oREpDG.outputCollection
 Set oOutputs = pRepSCM
 Set GraphicOutput = oOutputs.Item(1)
 
 'Query Def PlayerEx interface
 Set DefPlayerEx = pSymbolDefinition
 'Point on the symbol occurrence
 Set oSymbolOcc = DefPlayerEx.PlayingSymbol
 
 Set pEnumJDArgument = oSymbolOcc.IJDValuesArg.GetValues(3)

  If Not pEnumJDArgument Is Nothing Then
        pEnumJDArgument.Reset
        Do
            pEnumJDArgument.Next 1, arg, found
            If found = 0 Then Exit Do
                Select Case arg.index

                    Case 5
                        Set PC = arg.Entity
                        Length = PC.UomValue
                        Set PC = Nothing

                    Case 6
                        Set PC = arg.Entity
                        Width = PC.UomValue
                        Set PC = Nothing
                    Case 7
                        Set PC = arg.Entity
                        cutLength = PC.UomValue
                        Set PC = Nothing

                    Case 8
                        Set PC = arg.Entity
                        cutWidth = PC.UomValue
                        Set PC = Nothing
                
                End Select
                Set arg = Nothing
        Loop
  End If

  If cutLength > Length Then
    cutLength = Length
  End If
  If cutWidth > Width Then
    cutWidth = Width
  End If
  
  Dim dblPoints() As Double
  Dim oLineString As LineString3d
  Dim oGeomFactory As New GeometryFactory
  Dim oIJGeometryMisc As IJGeometryMisc
  Dim oAggr As Object
  Dim oWireBody As IJWireBody
  Dim intNumPoints As Integer
  
  
  'to avoid point 2 and for point 3 coinciding
  If Abs(Width - cutWidth) < distTol Then
    cutWidth = Width - distTol
  End If
  'to avoid point 5 and for point 4 coinciding
  If Abs(Length - cutLength) < distTol Then
    cutLength = Length - distTol
  End If

  If (cutLength < distTol) Or (cutWidth < distTol) Then
    ReDim dblPoints(1 To 15)
    intNumPoints = 5
'
'  O____________
' 5|1           |2
'  |            |
'  |            |
'  |            |
'  |____________|
'  4            3
    ' clockwise from top left (origin) when looking along plate's z
    'point1
    dblPoints(1) = 0
    dblPoints(2) = 0
    dblPoints(3) = 0
    'point2
    dblPoints(4) = Length
    dblPoints(5) = 0
    dblPoints(6) = 0
    'point3
    dblPoints(7) = Length
    dblPoints(8) = Width
    dblPoints(9) = 0
    
    'point4
    dblPoints(10) = 0
    dblPoints(11) = Width
    dblPoints(12) = 0
    'point5
    dblPoints(13) = 0
    dblPoints(14) = 0
    dblPoints(15) = 0
  Else
    ReDim dblPoints(1 To 18)
    intNumPoints = 6

'
'  O____________
' 6|1           |2
'  |            |
'  |           / 3
'  |          /
'  |_________/
'  5         4
    ' clockwise from top left (origin) when looking along plate's z
    'point1
    dblPoints(1) = 0
    dblPoints(2) = 0
    dblPoints(3) = 0
    'point2
    dblPoints(4) = Length
    dblPoints(5) = 0
    dblPoints(6) = 0
    'point3
    dblPoints(7) = Length
    dblPoints(8) = Width - cutWidth
    dblPoints(9) = 0
    
    'point4
    dblPoints(10) = Length - cutLength
    dblPoints(11) = Width
    dblPoints(12) = 0
    'point5
    dblPoints(13) = 0
    dblPoints(14) = Width
    dblPoints(15) = 0
    'point6 same as point1
    dblPoints(16) = 0
    dblPoints(17) = 0
    dblPoints(18) = 0
      
   End If

  

  Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, intNumPoints, dblPoints)
  
  Set oIJGeometryMisc = New IMSModelGeomOps.DGeomOpsMisc
  oIJGeometryMisc.CreateModelGeometryFromGType pOutputCollection.ResourceManager, oLineString, oAggr, oWireBody

  pOutputCollection.AddOutput "WireBodyShape", oWireBody
  Set oWireBody = Nothing
  Set oIJGeometryMisc = Nothing
  Set oGeomFactory = Nothing
  Set oAggr = Nothing
  
Exit Sub
ErrorHandler: HandleError MODULE, MT
End Sub


Private Sub SetCMOnRepresentation(pRep As IMSSymbolEntities.IJDRepresentation, LibraryName As String, MethodName As String, pSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition)
 Const MT = "SetCMOnRepresentation"
On Error GoTo ErrorHandler

 Dim mthCookie As Long
 Dim UserMethod As IJDUserMethods
 Dim oLib As IMSSymbolEntities.DLibraryDescription
 
 Set UserMethod = pSymbolDefinition
 mthCookie = UserMethod.GetMethodCookie(MethodName, LibraryName)
 Set oLib = UserMethod.GetLibrary(LibraryName)
 pRep.IJDRepresentationStdCustomMethod.SetCMEvaluate oLib.Cookie, mthCookie
 
 Set UserMethod = Nothing
Exit Sub
ErrorHandler:
    HandleError MODULE, MT

End Sub




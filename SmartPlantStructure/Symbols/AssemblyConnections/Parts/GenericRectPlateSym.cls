VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GenericRectPlateSym"
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
'The plate contour is created at using a local coordinate system with origin at the center of the plate.
'
'
'History
'    Raj Peringoth      03/26/03      Creation
'---------------------------------------------------------------------------------------
Option Explicit
Private Const MODULE = "GenericRectPlateSym"
Private Const SYMBOLDEFNAME = "SPSPartMacros.GenericRectPlateSym"


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
    oInput.Name = "Length"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oInput.Index = 1
    pIJDInputs.Add oInput
    
    oInput.Reset
    oInput.Name = "Width"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oInput.Index = 2
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
    
    Exit Sub

ErrorHandler: HandleError MODULE, MT
End Sub

Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, ByVal definitionParameters As Variant, ByVal pResourceMgr As Object) As Object
Const MT = "IJDUserSymbolServices_InstanciateDefinition"
On Error GoTo ErrorHandler

    Dim pSymbolDefinition As IJDSymbolDefinition
    Dim oSymbolFactory As New IMSSymbolEntities.DSymbolEntitiesFactory
    
    Set pSymbolDefinition = oSymbolFactory.CreateEntity(Definition, pResourceMgr)
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
 Dim Depth As Double, Width As Double
 Dim PC As IJDParameterContent



 Set oREpDG = pRepSCM
 Set pSymbolDefinition = oREpDG.Definition
 Set pOutputCollection = oREpDG.OutputCollection
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
                Select Case arg.Index

                    Case 1
                        Set PC = arg.Entity
                        Depth = PC.UomValue
                        Set PC = Nothing

                    Case 2
                        Set PC = arg.Entity
                        Width = PC.UomValue
                        Set PC = Nothing
                End Select
                Set arg = Nothing
        Loop
  End If

  Dim dblPoints(1 To 15) As Double
  Dim oLineString As LineString3d
  Dim oGeomFactory As New GeometryFactory
  Dim oIJGeometryMisc As IJGeometryMisc
  Dim oAggr As Object
  Dim oWireBody As IJWireBody
  Dim oMat As New DT4x4
  Dim x#, y#, z#

  
  ' clockwise from bottom left when looking along plate's z
  'point1
  dblPoints(1) = -Depth / 2#
  dblPoints(2) = -Width / 2#
  dblPoints(3) = 0#
  'point2
  dblPoints(4) = Depth / 2#
  dblPoints(5) = -Width / 2#
  dblPoints(6) = 0#
  'point3
  dblPoints(7) = Depth / 2#
  dblPoints(8) = Width / 2#
  dblPoints(9) = 0#
  'point4
  dblPoints(10) = -Depth / 2#
  dblPoints(11) = Width / 2#
  dblPoints(12) = 0#
  'point5 same as point1
  dblPoints(13) = -Depth / 2#
  dblPoints(14) = -Width / 2#
  dblPoints(15) = 0#
  

  oMat.LoadIdentity
  
  'X Axis
  oMat.IndexValue(0) = 1
  oMat.IndexValue(1) = 0
  oMat.IndexValue(2) = 0
  
  'Y Axis
  oMat.IndexValue(4) = 0
  oMat.IndexValue(5) = 1
  oMat.IndexValue(6) = 0
  
  'Z axis
  oMat.IndexValue(8) = 0
  oMat.IndexValue(9) = 0
  oMat.IndexValue(10) = 1
  
  'origin
  oMat.IndexValue(12) = 0
  oMat.IndexValue(13) = 0
  oMat.IndexValue(14) = 0
  

  Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dblPoints)
  oLineString.Transform oMat
  
  Set oIJGeometryMisc = New IMSModelGeomOps.DGeomOpsMisc
  oIJGeometryMisc.CreateModelGeometryFromGType pOutputCollection.ResourceManager, oLineString, oAggr, oWireBody

  pOutputCollection.AddOutput "WireBodyShape", oWireBody
  Set oMat = Nothing
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



VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PadPlateRectSym"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'---------------------------------------------------------------------------
'    Copyright (C) 2006 Intergraph Corporation. All rights reserved.
'
'
'Abstract
'    USS defining the shape of the parametric rect pad plate
'
'
'Notes
'The plate contour is created as following:
'   local coordinate system with origin at the center of the plate.
'   Corner radius at the corners
'
'History
'    AS      April 20, 2006      Creation
'---------------------------------------------------------------------------------------
Option Explicit
Private Const MODULE = "PadPlateRectSym"
Private Const SYMBOLDEFNAME = "SPSPartMacros.PadPlateRectSym"


Implements IJDUserSymbolServices


Private Sub IJDUserSymbolServices_InvokeRepresentation(ByVal pSymbolOccurrence As Object, ByVal pRepName As String, ByVal pOutputColl As Object, arrayOfInputs() As Variant)

End Sub

Private Function IJDUserSymbolServices_EditOccurence(pSymbolOccurrence As Object, ByVal pTransactionMgr As Object) As Boolean

End Function

Private Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
Const MT = "IJDUserSymbolServices_GetDefinitionName"
On Error GoTo ErrorHandler

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
    oInput.Name = "MemberPart"
    oInput.Properties = igDESCRIPTION_OPTIONAL
    oInput.index = 1
    pIJDInputs.Add oInput

    oInput.Reset
    oInput.Name = "Surface"
    oInput.Properties = igDESCRIPTION_OPTIONAL
    oInput.index = 2
    pIJDInputs.Add oInput
    
    oInput.Reset
    oInput.Name = "Length"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oInput.index = 3
    pIJDInputs.Add oInput
    
    oInput.Reset
    oInput.Name = "Width"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oInput.index = 4
    pIJDInputs.Add oInput
    
    oInput.Reset
    oInput.Name = "CornerRadius"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oInput.index = 5
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
 Dim depth As Double
 Dim Width As Double
 Dim CornerRadius As Double
 
 Dim PC As IJDParameterContent
 
 'Pad contour stuff
 Dim dblPoints(1 To 24) As Double
 Dim oLineString As LineString3d
 Dim oGeomFactory As New GeometryFactory
 Dim oIJGeometryMisc As IJGeometryMisc
 Dim oAggr As Object
 Dim oWireBody As IJWireBody
 Dim oMat As New DT4x4
 Dim x#, y#, z#
 Dim oCurveElems As IJElements
 Dim oLine As IJLine 'IngrGeom3D.Line3d
 Dim oArc As IJArc
 Dim pt1X As Double
 Dim pt1Y As Double
 Dim pt1Z As Double
 Dim pt2X As Double
 Dim pt2Y As Double
 Dim pt2Z As Double
 Dim centerX As Double
 Dim centerY As Double
 Dim centerZ As Double
 Dim oLine3D As ILines3d
 Dim oArc3d As IArcs3d
 Dim oComplexStr3d As IComplexStrings3d
 Dim oComplexString As IngrGeom3D.ComplexString3d 'IJComplexString
  
 Set oREpDG = pRepSCM
 Set pSymbolDefinition = oREpDG.definition
 Set pOutputCollection = oREpDG.outputCollection
 
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

                    Case 3
                        Set PC = arg.Entity
                        depth = PC.UomValue
                        Set PC = Nothing

                    Case 4
                        Set PC = arg.Entity
                        Width = PC.UomValue
                        Set PC = Nothing
                        
                    Case 5
                        Set PC = arg.Entity
                        CornerRadius = PC.UomValue
                        Set PC = Nothing
                End Select
                Set arg = Nothing
        Loop
  End If

    
  'Initialize vars
  Set oCurveElems = New JObjectCollection
  Set oLine3D = oGeomFactory.Lines3d
  Set oComplexStr3d = oGeomFactory.ComplexStrings3d
  
  ' clockwise from bottom left when looking along plate's z
  If CornerRadius > distTol Then
'       6      5
'        ******
'       *      *
'    7 *        * 4
'      *        *
'      *        *
'    8 *        * 3
'       *      *
'        ******
'       1      2
     
    Set oArc3d = oGeomFactory.Arcs3d
    
    'point1
    dblPoints(1) = -depth / 2# + CornerRadius
    dblPoints(2) = -Width / 2#
    dblPoints(3) = 0#
    'point2
    dblPoints(4) = depth / 2# - CornerRadius
    dblPoints(5) = -Width / 2#
    dblPoints(6) = 0#
    
    'point3
    dblPoints(7) = depth / 2#
    dblPoints(8) = -Width / 2# + CornerRadius
    dblPoints(9) = 0#
    
    'point4
    dblPoints(10) = depth / 2#
    dblPoints(11) = Width / 2# - CornerRadius
    dblPoints(12) = 0#
    
    'point5
    dblPoints(13) = depth / 2# - CornerRadius
    dblPoints(14) = Width / 2#
    dblPoints(15) = 0#
    'point6
    dblPoints(16) = -depth / 2# + CornerRadius
    dblPoints(17) = Width / 2#
    dblPoints(18) = 0#
    'point7
    dblPoints(19) = -depth / 2#
    dblPoints(20) = Width / 2# - CornerRadius
    dblPoints(21) = 0#
    'point8 same as point1
    dblPoints(22) = -depth / 2#
    dblPoints(23) = -Width / 2# + CornerRadius
    dblPoints(24) = 0#
    
    Set oLine = oLine3D.CreateBy2Points(Nothing, dblPoints(1), dblPoints(2), dblPoints(3), dblPoints(4), dblPoints(5), dblPoints(6))
    oCurveElems.Add oLine
    Set oComplexString = oComplexStr3d.CreateByCurves(Nothing, oCurveElems)
    Set oCurveElems = Nothing

    centerX = depth / 2# - CornerRadius
    centerY = -Width / 2# + CornerRadius
    centerZ = 0#
    Set oArc = oArc3d.CreateByCenterStartEnd(Nothing, centerX, centerY, centerZ, dblPoints(4), dblPoints(5), dblPoints(6), dblPoints(7), dblPoints(8), dblPoints(9))
    oComplexString.AddCurve oArc, True
    
    Set oLine = oLine3D.CreateBy2Points(Nothing, dblPoints(7), dblPoints(8), dblPoints(9), dblPoints(10), dblPoints(11), dblPoints(12))
    oComplexString.AddCurve oLine, True
  
    centerX = depth / 2# - CornerRadius
    centerY = Width / 2# - CornerRadius
    centerZ = 0#
    Set oArc = oArc3d.CreateByCenterStartEnd(Nothing, centerX, centerY, centerZ, dblPoints(10), dblPoints(11), dblPoints(12), dblPoints(13), dblPoints(14), dblPoints(15))
    oComplexString.AddCurve oArc, True
  
    Set oLine = oLine3D.CreateBy2Points(Nothing, dblPoints(13), dblPoints(14), dblPoints(15), dblPoints(16), dblPoints(17), dblPoints(18))
    oComplexString.AddCurve oLine, True
  
    centerX = -depth / 2# + CornerRadius
    centerY = Width / 2# - CornerRadius
    centerZ = 0#
    Set oArc = oArc3d.CreateByCenterStartEnd(Nothing, centerX, centerY, centerZ, dblPoints(16), dblPoints(17), dblPoints(18), dblPoints(19), dblPoints(20), dblPoints(21))
    oComplexString.AddCurve oArc, True
  
    Set oLine = oLine3D.CreateBy2Points(Nothing, dblPoints(19), dblPoints(20), dblPoints(21), dblPoints(22), dblPoints(23), dblPoints(24))
    oComplexString.AddCurve oLine, True
  
    centerX = -depth / 2# + CornerRadius
    centerY = -Width / 2# + CornerRadius
    centerZ = 0#
    Set oArc = oArc3d.CreateByCenterStartEnd(Nothing, centerX, centerY, centerZ, dblPoints(22), dblPoints(23), dblPoints(24), dblPoints(1), dblPoints(2), dblPoints(3))
    oComplexString.AddCurve oArc, True
  
  Else
'      4        3
'      **********
'      *        *
'      *        *
'      *        *
'      *        *
'      **********
'      1        2
   'if the radius is too small, just create rectangular pad
    dblPoints(1) = -depth / 2#
    dblPoints(2) = -Width / 2#
    dblPoints(3) = 0#
    'point2
    dblPoints(4) = depth / 2#
    dblPoints(5) = -Width / 2#
    dblPoints(6) = 0#
    'point3
    dblPoints(7) = depth / 2#
    dblPoints(8) = Width / 2#
    dblPoints(9) = 0#
    'point4
    dblPoints(10) = -depth / 2#
    dblPoints(11) = Width / 2#
    dblPoints(12) = 0#
    'point5 same as point1
    dblPoints(13) = -depth / 2#
    dblPoints(14) = -Width / 2#
    dblPoints(15) = 0#
    
    Set oLine = oLine3D.CreateBy2Points(Nothing, dblPoints(1), dblPoints(2), dblPoints(3), dblPoints(4), dblPoints(5), dblPoints(6))
    oCurveElems.Add oLine
    Set oComplexString = oGeomFactory.ComplexString3d.CreateByCurves(Nothing, oCurveElems)
    Set oCurveElems = Nothing

    Set oLine = oLine3D.CreateBy2Points(Nothing, dblPoints(4), dblPoints(5), dblPoints(6), dblPoints(7), dblPoints(8), dblPoints(9))
    oComplexString.AddCurve oLine, True
  
    Set oLine = oLine3D.CreateBy2Points(Nothing, dblPoints(7), dblPoints(8), dblPoints(9), dblPoints(10), dblPoints(11), dblPoints(12))
    oComplexString.AddCurve oLine, True
  
    Set oLine = oLine3D.CreateBy2Points(Nothing, dblPoints(10), dblPoints(11), dblPoints(12), dblPoints(13), dblPoints(14), dblPoints(15))
    oComplexString.AddCurve oLine, True
  
  End If

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
  
  oComplexString.Transform oMat
  
  Set oIJGeometryMisc = New IMSModelGeomOps.DGeomOpsMisc
  oIJGeometryMisc.CreateModelGeometryFromGType pOutputCollection.ResourceManager, oComplexString, oAggr, oWireBody
  Set oIJGeometryMisc = Nothing
  
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




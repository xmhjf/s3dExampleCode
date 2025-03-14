VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OctagonFndSym"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'******************************************************************
' Copyright (C) 2003, Intergraph Corporation. All rights reserved.
'
'File
'    OctagonFndSym.cls
'
'Author
'       1-Mar-03        Sudha Srikakolapu
'
'Description
'
'Notes
'
'History:
'
'   Jun 24, 2009        WR      TR-145118 - Updated the Physical method to call
'                               the new CreateSolidAsPlanes method
'                               in EqpFndCommon.
'*******************************************************************

Private Const MODULE = "OctagonFndSym"
Const m_ItemProgId As String = "SPSEqpFndMacros.OctagonFndSym"
Const CheckProgId As String = "SPSValidateArgs.CheckFunctions"

Implements IJDUserSymbolServices

Private Enum InputIndex
    EQPPORT_INDEX = 1
    SUPPPLANE_INDEX
    OCTAGONDIAMETER_INDEX
    OCTAGONHT_INDEX
    OCTAGONANGLE_INDEX
    OCTAGONSIZEBYRULE_INDEX
    OCTAGONEDGECLEARANCE_INDEX
    MATERIAL_INDEX
    GRADE_INDEX
End Enum

Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, ByVal definitionParameters As Variant, ByVal pResourceMgr As Object) As Object
Const METHOD = "IJDUserSymbolServices_InitializeSymbolDefinition"
On Error GoTo ErrorHandler

    'Create a Symbol Definition Object.
    Dim pSymbolFactory As New DSymbolEntitiesFactory
    Dim pSymbolDefinition As IJDSymbolDefinition
    
    Set pSymbolDefinition = pSymbolFactory.CreateEntity(definition, pResourceMgr)
    pSymbolDefinition.ProgId = m_ItemProgId
    pSymbolDefinition.CodeBase = CodeBase
    pSymbolDefinition.name = pSymbolDefinition.ProgId

    IJDUserSymbolServices_InitializeSymbolDefinition pSymbolDefinition
   
    'returned symbol defintion
    Set IJDUserSymbolServices_InstanciateDefinition = pSymbolDefinition
    Set pSymbolFactory = Nothing
    Set pSymbolDefinition = Nothing
    
    Exit Function
 
ErrorHandler:
    HandleError MODULE, METHOD
End Function

Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(pSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition)
Const METHOD = "IJDUserSymbolServices_InitializeSymbolDefinition"
On Error GoTo ErrorHandler
     
    pSymbolDefinition.IJDInputs.RemoveAllInput
    pSymbolDefinition.IJDRepresentations.RemoveAllRepresentation
    pSymbolDefinition.IJDRepresentationEvaluations.RemoveAllRepresentationEvaluations
    
    On Error GoTo ErrorHandler
    
    ' Set the input to the definition
    Dim iInputs As IMSSymbolEntities.IJDInputs
    Set iInputs = pSymbolDefinition
    
    Dim iUM As IMSSymbolEntities.IJDUserMethods
    Set iUM = pSymbolDefinition
    
    Dim libDesc As New DLibraryDescription
    Dim mCookie As Long
    Dim libCookie As Long
    Dim methodCookie As Long

    libDesc.name = "mySelfAsLib"
    libDesc.Type = imsLIBRARY_IS_ACTIVEX
    libDesc.Properties = imsLIBRARY_AUTO_EXTRACT_METHOD_COOKIES
    libDesc.Source = m_ItemProgId
    
    pSymbolDefinition.IJDUserMethods.SetLibrary libDesc
    
    Dim ChecklibDesc As New DLibraryDescription
    Dim ChecklibCookie As Long
    Dim GTZeroCheck As Long
    Dim DegCheck As Long
    
    ChecklibDesc.name = "CMCheckLib"
    ChecklibDesc.Type = imsLIBRARY_IS_ACTIVEX
    ChecklibDesc.Properties = imsLIBRARY_AUTO_EXTRACT_METHOD_COOKIES
    ChecklibDesc.Source = CheckProgId
    pSymbolDefinition.IJDUserMethods.SetLibrary ChecklibDesc
    ChecklibCookie = ChecklibDesc.Cookie

    GTZeroCheck = pSymbolDefinition.IJDUserMethods.GetMethodCookie("GTZero", ChecklibCookie)
    DegCheck = pSymbolDefinition.IJDUserMethods.GetMethodCookie("IN360Deg", ChecklibCookie)
    
'    Get the lib/method cookie
    libCookie = libDesc.Cookie
     
    ' set to variable number of inputs for suppored equipments and supporting plane
    Dim inputsProp As IMSDescriptionProperties
    inputsProp = pSymbolDefinition.IJDInputs.Property
    pSymbolDefinition.IJDInputs.Property = inputsProp Or igCOLLECTION_VARIABLE
        
    Dim pIJDInput As IMSSymbolEntities.IJDInput
    Set pIJDInput = New IMSSymbolEntities.DInput
    
    Dim oInput As IMSSymbolEntities.IJDInput
    Set oInput = New IMSSymbolEntities.DInput
    Dim PC As IMSSymbolEntities.IJDParameterContent
    Set PC = New IMSSymbolEntities.DParameterContent
    PC.Type = igValue
    
    oInput.name = "EquipmentPorts"
    oInput.Description = "Supporting Equipment Port(s)"
    iInputs.SetInput oInput, EQPPORT_INDEX
    oInput.Reset

    oInput.name = "SupportedPlane"
    oInput.Description = "Supporting Surface/Plane"
    oInput.Properties = igDESCRIPTION_OPTIONAL
    iInputs.SetInput oInput, SUPPPLANE_INDEX
    oInput.Reset

    oInput.name = "Diameter"
    oInput.Description = "diameter of the equipment containing bolt-holes position"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, GTZeroCheck
    PC.UomValue = 0.55
    oInput.DefaultParameterValue = PC
    iInputs.SetInput oInput, OCTAGONDIAMETER_INDEX
    oInput.Reset
    PC.Reset
    
    oInput.name = "OctagonHeight"
    oInput.Description = "Height of the Octagon foundation"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, GTZeroCheck
    PC.UomValue = 0.5
    oInput.DefaultParameterValue = PC
    iInputs.SetInput oInput, OCTAGONHT_INDEX
    oInput.Reset
    PC.Reset
    
    oInput.name = "RotationAngle"
    oInput.Description = "Rotation ANgle the Octagon foundation"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, DegCheck
    PC.UomValue = 0
    oInput.DefaultParameterValue = PC
    iInputs.SetInput oInput, OCTAGONANGLE_INDEX
    oInput.Reset
    PC.Reset
    
    oInput.name = "IsOctagonSizeDrivenByRule"
    oInput.Description = "Octagon Size Driven By Rule (bolt hole locations)"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 1
    oInput.DefaultParameterValue = PC
    iInputs.SetInput oInput, OCTAGONSIZEBYRULE_INDEX
    oInput.Reset
    PC.Reset
    
    oInput.name = "OctagonEdgeClearance"
    oInput.Description = "OctagonEdgeClearance"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, GTZeroCheck
    PC.UomValue = 0.0001
    oInput.DefaultParameterValue = PC
    iInputs.SetInput oInput, OCTAGONEDGECLEARANCE_INDEX
    oInput.Reset
    PC.Reset
    
    oInput.name = "SPSMaterial"
    oInput.Description = "SPSMaterial"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    PC.Type = igString
    PC.String = "Concrete"
    oInput.DefaultParameterValue = PC
    iInputs.SetInput oInput, MATERIAL_INDEX
    oInput.Reset
    PC.Reset
    
    oInput.name = "SPSGrade"
    oInput.Description = "SPSGrade"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    PC.Type = igString
    PC.String = "Fc 4000"
    oInput.DefaultParameterValue = PC
    iInputs.SetInput oInput, GRADE_INDEX
    oInput.Reset
    PC.Reset
    
    Dim pIReps As IMSSymbolEntities.IJDRepresentations
    Set pIReps = pSymbolDefinition
    Dim pIRep As IMSSymbolEntities.IJDRepresentation
    Set pIRep = New IMSSymbolEntities.DRepresentation
    
    pIRep.name = "Physical"
    pIRep.Description = "Physical representation"
    pIRep.RepresentationId = SimplePhysical
    mCookie = iUM.GetMethodCookie("Physical", libCookie)
    pIRep.IJDRepresentationStdCustomMethod.SetCMEvaluate libCookie, mCookie
    
    Dim pOutputs As IMSSymbolEntities.IJDOutputs
    Set pOutputs = pIRep
    pOutputs.Property = igCOLLECTION_VARIABLE ' declare that the number of outputs is variable
    
    Dim output As IMSSymbolEntities.IJDOutput
    Set output = New IMSSymbolEntities.DOutput
    
    output.name = "Octagon"
    output.Description = "Octagon"
    pOutputs.SetOutput output
    output.Reset
    
    pIReps.SetRepresentation pIRep 'Add representation to definition

    pIRep.name = "DetailPhysical"
    pIRep.Description = "DetailPhysical representation"
    pIRep.RepresentationId = DetailPhysical
    mCookie = iUM.GetMethodCookie("Physical", libCookie)
    pIRep.IJDRepresentationStdCustomMethod.SetCMEvaluate libCookie, mCookie
    Set pOutputs = pIRep
    pIReps.SetRepresentation pIRep 'Add representation to definition

    ' Set definition cache properties
    pSymbolDefinition.CacheOption = igSYMBOL_CACHE_OPTION_NOT_SHARED
    'as this symbol def has declared a graphic object as input
    ' GeomOption option will be set to igSYMBOL_GEOM_FIX_TO_ID by the symbol machinerary
    'Because of this the  outputs will be transformed during MDR and the Symbol geometry will
    ' end up in an incorrect location. So resetting the flag - DI226263
    pSymbolDefinition.GeomOption = igSYMBOL_GEOM_FREE

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Private Function IJDUserSymbolServices_EditOccurence(pSymbolOccurrence As Object, ByVal pTransactionMgr As Object) As Boolean

End Function

Private Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
     IJDUserSymbolServices_GetDefinitionName = m_ItemProgId
End Function

Private Sub IJDUserSymbolServices_InvokeRepresentation(ByVal pSymbolOccurrence As Object, ByVal pRepName As String, ByVal pOutputColl As Object, arrayOfInputs() As Variant)

End Sub

Public Sub Physical(pIRepSCM As IJDRepresentationStdCustomMethod)
Const METHOD = "Physical"
On Error GoTo ErrorHandler

    Dim OctagonR As Double, octagonDiameter As Double
    Dim OctagonHeight As Double
    Dim rotationAngle As Double
    
    Dim FoundationType As Double
    Dim elems As IJElements
    Dim pObj As Object
 
    Dim pRepDG As IJDRepresentationDuringGame
    Dim pOC As IJDOutputCollection
    
    Set pRepDG = pIRepSCM
    Set pOC = pRepDG.outputCollection
    
    'remove all outputs
    Dim oRep As IJDRepresentation
    Dim oOutputs As IJDOutputs
    Set oRep = pOC.definition.IJDRepresentations.GetRepresentationByName("Physical")
    Set oOutputs = oRep
    oOutputs.RemoveAllOutput
    
    Dim pInputs As IJDInputs
    Set pInputs = pRepDG.definition.IJDInputs
   
    Dim tmpLineString As IngrGeom3D.LineString3d
    Dim oProjection As IngrGeom3D.Projection3d
    
    'assign to meaningful variables from the input array
    octagonDiameter = pInputs.GetInputByIndex(OCTAGONDIAMETER_INDEX).IJDInputDuringGame.Result.UomValue
    OctagonR = octagonDiameter / 2
    
    OctagonHeight = pInputs.GetInputByIndex(OCTAGONHT_INDEX).IJDInputDuringGame.Result.UomValue
     
    rotationAngle = pInputs.GetInputByIndex(OCTAGONANGLE_INDEX).IJDInputDuringGame.Result.UomValue

'    get the playing symbol occ matrix
    Dim pSuppPlane As IJPlane
    Dim pRC2 As IJDReferencesCollection
    
    Dim pDefPlayerEx As IJDDefinitionPlayerEx
    Set pDefPlayerEx = pRepDG.definition.IJDDefinitionPlayerEx
      
    Dim Arguments As IJDArguments
    Set Arguments = pDefPlayerEx.SymbolInputs
        
    Dim pArg1 As IJDArgument
    If TypeOf Arguments.Item(2) Is IJDArgument Then Set pArg1 = Arguments.Item(2)
    If Not pArg1 Is Nothing And Not pArg1.Entity Is Nothing Then
        If TypeOf pArg1.Entity Is IJDReferencesCollection Then Set pRC2 = pArg1.Entity
    End If
    
    Dim bSuppPlaneDefined As Boolean
    
    If pRC2 Is Nothing Then
        bSuppPlaneDefined = False
    Else
        Set pSuppPlane = pRC2.IJDEditJDArgument.GetEntityByIndex(1)
    End If
    
    If pSuppPlane Is Nothing Then bSuppPlaneDefined = False Else bSuppPlaneDefined = True
    
    Set Arguments = Nothing
    Set pArg1 = Nothing
    
    '=================
    'BUILD Octagon
    '=================
        
    Dim Dirvec              As IJDVector
    Dim Octagon(0 To 26)    As Double
    
    Set Dirvec = New DVector
    Dirvec.Set 0, 0, -OctagonHeight
    Dirvec.Length = 1

    InitOctCurvePoints Octagon, OctagonR * 2, 0 ' TR#65124- rectfied octagon vertices coordinates
    
    Dim OutStr As String
    OutStr = "Octagon"
    
    Dim ii As Integer
    ii = 0
    
    '================================================================================
    Dim oGeomFactory As New IngrGeom3D.GeometryFactory
    Set tmpLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 9, Octagon)

    If bSuppPlaneDefined = True Then    ' need to clip against support plane for sloped surfaces
        'increased projection length by 20% as intersect was not giving proper result.
        Set oProjection = oGeomFactory.Projections3d.CreateByCurve(Nothing, _
                                tmpLineString, Dirvec.x, Dirvec.y, Dirvec.z, OctagonHeight + OctagonR, True)
                                
        Dim pSymbol As IJDSymbol
        Set pSymbol = pDefPlayerEx.PlayingSymbol
        Dim pOcc As IJDOccurrence
        Set pOcc = pSymbol
             
        Dim oMatrix As IJDT4x4
        Set oMatrix = New DT4x4
        Set oMatrix = pOcc.Matrix
        
        Dim z1 As Double, z2 As Double, z3 As Double
        
        z1 = oMatrix.IndexValue(8)
        z2 = oMatrix.IndexValue(9)
        z3 = oMatrix.IndexValue(10)
        
        Dim oClipSurfWrapper As SP3DSPSClipGeomService.SPSClipSurfWrapper
        Set oClipSurfWrapper = New SPSClipSurfWrapper
        
        Dim dTol As Double
        dTol = 0.000001
        
        Dim oinverseMatrix As IJDT4x4
        Set oinverseMatrix = New DT4x4
        
        Set oinverseMatrix = oMatrix.Clone
        oinverseMatrix.Invert
        
        Dim oDummySupportPlane As IJPlane
        Dim dox As Double, doy As Double, doz As Double
        Dim nox As Double, noy As Double, noz As Double
        pSuppPlane.GetRootPoint dox, doy, doz
        pSuppPlane.GetNormal nox, noy, noz
        
        ' make sure plane normal is in the same direction as the symbol
        If z1 * nox + z2 * noy + z3 * noz > 0 Then
        Else
            nox = -nox
            noy = -noy
            noz = -noz
        End If

        Set oDummySupportPlane = oGeomFactory.Planes3d.CreateByPointNormal(Nothing, _
                                                                            dox, doy, doz, _
                                                                            nox, noy, noz)   ' tr# 46074
                                                                       
        oDummySupportPlane.Transform oinverseMatrix
        
        Set oinverseMatrix = Nothing
        
        Dim pClipSurfs As IJElements
        Dim oTopPlane As New Plane3d
        Dim oBottomPlane As New Plane3d
        Dim tmpSurf As IJSurface
        Set tmpSurf = oDummySupportPlane
        Dim Geomconst As Geom3dIntersectConstants
        Dim pIntersectSurfs As IJElements
        
        'Top plane
        Set oTopPlane = oGeomFactory.Planes3d.CreateByOuterBdry(pOC.ResourceManager, tmpLineString)
        Call InitNewOutput(pOC, "Octagon Top Plane")
        pOC.AddOutput "Octagon Top Plane", oTopPlane
        
        'Bottom Plane
        Call tmpSurf.Intersect(oProjection, pIntersectSurfs, Geomconst)
        If pIntersectSurfs.count = 1 Then
            If pIntersectSurfs.Item(1) Is Nothing Then
            Else
                ' tr 53900
                Dim oIntersectCurve As IJCurve
                Set oIntersectCurve = pIntersectSurfs.Item(1)
                If oIntersectCurve.Form = CURVE_FORM_OPEN Then
                Else
                    Set oBottomPlane = oGeomFactory.Planes3d.CreateByOuterBdry(pOC.ResourceManager, pIntersectSurfs.Item(1))
                    If Not oBottomPlane Is Nothing Then
                        Call InitNewOutput(pOC, "Octagon Bottom Plane")
                        pOC.AddOutput "Octagon Bottom Plane", oBottomPlane
                    End If
                End If
            End If
        End If
        
        'TR-CP�49367  Equipment Foundations should offer correct surface normals for relationships
        oProjection.ReverseNormal = True 'Reverse the normal so that the planes we get back has normals outward
        Call oClipSurfWrapper.GetClipSurfaceWithPlanes(pOC.ResourceManager, _
                                                        oProjection, _
                                                        oDummySupportPlane, _
                                                        dTol, _
                                                        pClipSurfs)
                                                            
        Dim iClipSurfCount As Integer
        iClipSurfCount = pClipSurfs.count
        Dim pSurface As iJDObject
        If iClipSurfCount >= 1 Then
            
            For ii = 1 To iClipSurfCount
                Set pSurface = pClipSurfs(ii)
                pOC.AddOutput OutStr, pSurface
                OutStr = OutStr & Trim$(Str$(ii))
                Call InitNewOutput(pOC, OutStr)
                Set pSurface = Nothing
            Next ii
            
            If pIntersectSurfs Is Nothing Then
            Else
                pIntersectSurfs.Clear
                Set pIntersectSurfs = Nothing
            End If
            Set tmpSurf = Nothing
            Set oTopPlane = Nothing
            Set oBottomPlane = Nothing
                 
        Else
            Set elems = CreateSolidAsPlanes(pOC.ResourceManager, tmpLineString, Dirvec, OctagonHeight)
            For Each pObj In elems
                pOC.AddOutput OutStr, pObj
                ii = ii + 1
                OutStr = OutStr & Trim$(Str$(ii))
                Call InitNewOutput(pOC, OutStr)
            Next
            Set elems = Nothing
        End If
    Else
        Set elems = CreateSolidAsPlanes(pOC.ResourceManager, tmpLineString, Dirvec, OctagonHeight)
        For Each pObj In elems
            pOC.AddOutput OutStr, pObj
            ii = ii + 1
            OutStr = OutStr & Trim$(Str$(ii))
            Call InitNewOutput(pOC, OutStr)
        Next
        Set elems = Nothing
    End If
    
    
    '================================================================================
    
    Set oProjection = Nothing
    Set tmpLineString = Nothing
    Set pOC = Nothing
    Set Dirvec = Nothing
    Set oGeomFactory = Nothing
    Exit Sub
    
ErrorHandler:
    HandleError MODULE, METHOD
End Sub




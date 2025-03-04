VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BUCan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'******************************************************************
' Copyright (C) 2006, Intergraph Corporation. All rights reserved.
'
'File
'    BUCan.cls
'
'Author
'       4 March 2008  ras
'
'Description
'       BuiltUp Can definition
'
'Notes
'
'History:
'
'   22-Sept-2009 GG TR#167167 - DesignedMember does not set PG of plates to its own PG
'   06-Aug-2010 GG DI-169828  SCHEMA: Interface to indicate member symmetry and open/closed cross sections
'*******************************************************************
'                   <Radius>
'
'                  (Z)
'                   |     ----
'                   |   /      \
'                   | /     -    \
'                   ||    /   \   |
'                   ||   |  +  |  |
'                   ||    \   /   |
'                   | \     -    /
'                   |   \      /
'                   |     ----
' (Y) --------------+-------------------- . (-Y)
'                 / |
'               /   |
'             /     |
'           /       |
'         /         |
'       /           |
'     (-X)          |
'                   |
'                  (-Z)
'
'  (X) is into the paper, direction of the extrusion
'  (0,0) is Cardinal Point = 1


Option Explicit

Private Const MODULE = "BUCan"
Private Const strSourceFile = "BUCan.def"
Private Const CONST_ItemProgId As String = "SM3DBUCan.BUCan"

Private m_strErrorDescr As String
Private m_oLocalizer As IJLocalizer
Private m_oDesignMemberHelper As BUHelperUtils
Private m_oCalcXProps As BUCalcSectionProperties

Private Enum BUTubeMembers
  Tube = 3
End Enum

Implements IJDUserSymbolServices
Implements IJUserAttributeMgmt
Implements ISPSDesignedMemberHelper
Implements ICustomSectionShapeService

Private Function IJDUserSymbolServices_EditOccurence(pSymbolOccurrence As Object, ByVal pTransactionMgr As Object) As Boolean
     IJDUserSymbolServices_EditOccurence = False
End Function

Private Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
     IJDUserSymbolServices_GetDefinitionName = CONST_ItemProgId
End Function

Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(pDefinition As IMSSymbolEntities.IJDSymbolDefinition)
Const METHOD = "IJDUserSymbolServices_InitializeSymbolDefinition"
On Error GoTo ErrorHandler
     pDefinition.SupportOnlyOption = igSYMBOL_NOT_SUPPORT_ONLY
     pDefinition.MetaDataOption = igSYMBOL_DYNAMIC_METADATA

     ' Define the inputs -
     Dim pIH As IJDInputsHelper
     Set pIH = New InputHelper
     pIH.definition = pDefinition
     pIH.SetInput "DefiningCurve"
     
     ' Aggregator Type
     Dim pAD As IJDAggregatorDescription
     Set pAD = pDefinition
     pAD.AggregatorClsid = "{F4CDE773-A760-4561-A43A-D44A9C8340A7}" 'CSPSDesignedMember
     pAD.UserTypeClsid = "{66AF28B3-72A6-407c-A9EA-FF139FC6F3E5}"   '**BUCan (Generated here for each new DM)
     pAD.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructAsm"
     pAD.SetCMConstruct imsCOOKIE_ID_USS_LIB, "CMConstructAsm"
     pAD.SetCMSetInputs -1, -1
     pAD.SetCMRemoveInputs -1, -1
     Set pAD = Nothing
      
     ' Aggregator property
     Dim pAPDs As IJDPropertyDescriptions
     Set pAPDs = pDefinition
     With pAPDs
        .RemoveAll ' Remove all the previous property descriptions
        .AddProperty "IStructCrossSection", 1, IStructCrossSection, , imsCOOKIE_ID_USS_LIB
        .AddProperty "IStructCrossSectionDimensions", 2, IStructCrossSectionDimensions, "CMEvaluateCAO", imsCOOKIE_ID_USS_LIB
        .AddProperty "IStructCrossSectionDesignProperties", 3, IStructCrossSectionDesignProperties, "CMEvaluateCAO", imsCOOKIE_ID_USS_LIB
        .AddProperty "ISPSDesignedMemberDesignNotifyInput", 4, "ISPSDesignedMemberDesignNotifyInput", "CMEvaluateCAO1", imsCOOKIE_ID_USS_LIB
        .AddProperty "IUABuiltUpLengthExt", 5, IID_IUABuiltUpLengthExt, , imsCOOKIE_ID_USS_LIB
        .AddProperty "IUABuiltUpTube", 6, IID_IUABuiltUpTube, , imsCOOKIE_ID_USS_LIB
        .AddProperty "IUABuiltUpCan", 7, IID_IUABuiltUpCan, , imsCOOKIE_ID_USS_LIB
        .AddProperty "IUABuiltUpCompute", 8, IID_IUABuiltUpCompute, , imsCOOKIE_ID_USS_LIB
'        .AddProperty "IUABuiltUpCone1", 9, IID_IUABuiltUpCone1, , imsCOOKIE_ID_USS_LIB
'        .AddProperty "IUABuiltUpCone2", 10, IID_IUABuiltUpCone2, , imsCOOKIE_ID_USS_LIB
         
         
     End With
     Set pAPDs = Nothing
               
     ' Define the members
     Dim pMemberDescriptions As IJDMemberDescriptions
     Dim pMemberDescription As IJDMemberDescription
     Dim pPropertyDescriptions As IJDPropertyDescriptions
     Set pMemberDescriptions = pDefinition
     
     ' Remove all the previous member descriptions
     pMemberDescriptions.RemoveAll
     
'''''
'' BuiltUpCan Curve To Revolve
'''''
      
     Set pMemberDescription = pMemberDescriptions.AddMember("BuiltUpCanCurveToRevolve", 1, "CMConstructCurveToRevolve", imsCOOKIE_ID_USS_LIB)
     pMemberDescription.RelationshipClsid = AssemblyMembers1RelationshipCLSID
     Set pPropertyDescriptions = pMemberDescription
     pPropertyDescriptions.AddProperty "BuiltUpCanCurveToRevolveProperties", 1, IJCurve, "CMComputeCurveToRevolve", imsCOOKIE_ID_USS_LIB

'''''
'' BuiltUpCan's Axis To Revolve Around
'''''

     Set pMemberDescription = pMemberDescriptions.AddMember("BuiltUpCanAxisOfRevolution", 2, "CMConstructAxisOfRevolution", imsCOOKIE_ID_USS_LIB)
     pMemberDescription.RelationshipClsid = AssemblyMembers1RelationshipCLSID
     Set pPropertyDescriptions = pMemberDescription
     pPropertyDescriptions.AddProperty "BuiltUpCanAxisOfRevolutionProperties", 1, IJCurve, "CMComputeAxisOfRevolution", imsCOOKIE_ID_USS_LIB

'''''
'' BuiltUpCan Plate
'''''
     Set pMemberDescription = pMemberDescriptions.AddMember("Can", 3, "CMConstructBuiltUpCanPlateSystem", imsCOOKIE_ID_USS_LIB)
     pMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructBuiltUpCanPlateSystem"
     pMemberDescription.RelationshipClsid = AssemblyMembers1RelationshipCLSID
     pMemberDescription.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CMMigrateBuiltUpCanPlateSystem"
          
     Set pPropertyDescriptions = pMemberDescription
     pPropertyDescriptions.AddProperty "BuiltUpCanProperties", 1, IJCurve, "CMComputeBuiltUpCanPlateSystem", imsCOOKIE_ID_USS_LIB
             
     Set pMemberDescriptions = Nothing
     Set pMemberDescription = Nothing
     Set pPropertyDescriptions = Nothing
     
Exit Sub
ErrorHandler:  HandleError MODULE, METHOD
End Sub

Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, ByVal defParams As Variant, ByVal pResourceMgr As Object) As Object
' This method is in charge of the creation of the symbol definition object
Const METHOD = "IJDUserSymbolServices_InstanciateDefinition"
On Error GoTo ErrorHandler
     Dim pDefinition As IJDSymbolDefinition
     Dim pFact As IJCAFactory
     Set pFact = New CAFactory
     Set pDefinition = pFact.CreateCAD(pResourceMgr)
     
     ' Set definition progId and codebase
     pDefinition.ProgId = CONST_ItemProgId
     pDefinition.CodeBase = CodeBase
     
     ' Initialize the definition
     IJDUserSymbolServices_InitializeSymbolDefinition pDefinition
     pDefinition.Name = IJDUserSymbolServices_GetDefinitionName(defParams)
     
     ' Persistence behavior
     pDefinition.SupportOnlyOption = igSYMBOL_NOT_SUPPORT_ONLY
     pDefinition.MetaDataOption = igSYMBOL_DYNAMIC_METADATA
     
     'returned symbol definition
     Set IJDUserSymbolServices_InstanciateDefinition = pDefinition
  
Exit Function
ErrorHandler:  HandleError MODULE, METHOD
End Function

Private Sub IJDUserSymbolServices_InvokeRepresentation(ByVal pSymbolOccurrence As Object, ByVal pRepName As String, ByVal pOutputColl As Object, arrayOfInputs() As Variant)

End Sub

'-------------------------------------------------------
'                       BuiltUpCan
'-------------------------------------------------------

Public Sub CMConstructCurveToRevolve(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObj As Object)
Const METHOD = "CMConstructCurveToRevolve "
On Error GoTo ErrorHandler

    ' Just Create a initial dummy Curve with points
    Dim pCurve As IJCurve
    m_oDesignMemberHelper.CreatePositionLineStringCurveBy2PointsWithNPoints pResourceManager, -0.25, -1, 0, _
                                                                                              1.25, -1, 0, _
                                                                                              6, pCurve
    CopyPermissionGroup pCurve, pMemberDescription.CAO
    ' the called function interpolates the points from start to end so there is no need
    ' to set individual points
'    Dim oLineString3d As LineString3d
'    Set oLineString3d = pCurve
'    oLineString3d.SetPoint 1, -0.25, -1, 0
'    oLineString3d.SetPoint 2, 0, -1, 0
'    oLineString3d.SetPoint 3, 0.25, -2, 0
'    oLineString3d.SetPoint 4, 0.75, -2, 0
'    oLineString3d.SetPoint 5, 1, -1, 0
'    oLineString3d.SetPoint 6, 1.25, -1, 0
    
    Set pObj = pCurve
    
    Dim oControlFlags As IJControlFlags
    Set oControlFlags = pObj
    oControlFlags.ControlFlags(&H4) = &H4
    
Exit Sub
ErrorHandler:  HandleError MODULE, METHOD
End Sub

Public Sub CMComputeCurveToRevolve(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMComputeCurveToRevolve"
    On Error GoTo ErrorHandler

    Dim ii As Long, jj As Long, nPoints As Long
    Dim dDiameterTube As Double, dDiameterStart As Double, dDiameterEnd As Double
    Dim dRadiusTube As Double, dRadiusStart As Double, dRadiusEnd As Double
    Dim dStartConeLength As Double, dEndConeLength As Double
    Dim dLength As Double, dAxisLength As Double
    Dim eCanType As Long            ' 1 = InLine or no CanRule, 2 = stubEnd
    Dim oSmartOcc As IJSmartOccurrence
    
    Dim oPts(1 To 6) As IJDPosition
    Dim oDir As IJDVector

    Dim dTubeThickness As Double, dStartConeThickness As Double, dEndConeThickness As Double
    Dim bHasStartCone As Boolean, bHasEndCone As Boolean
    
    Dim bDebugPoints As Boolean ' set to true for debugging
    bDebugPoints = False
    
    Dim distTol As Double, distTweak As Double
    
    Set oSmartOcc = pPropertyDescriptions.CAO
    distTol = 0.001         ' tolerance
    distTweak = 0.005       ' distance to adjust outer knuckle points
    
    If Not AreOccurrencePropertiesValid(oSmartOcc) Then
        GoTo ErrorHandler
    End If
    
    GetBuiltUpCanExtrusionParameters oSmartOcc, _
                                     dDiameterStart, _
                                     dDiameterEnd, _
                                     dStartConeLength, _
                                     dEndConeLength, _
                                     dDiameterTube, _
                                     dLength, _
                                     dAxisLength, _
                                     eCanType

    ' need thicknesses to shift curveToRevolve to inner diameter
    GetThicknesses oSmartOcc, dTubeThickness, dStartConeThickness, dEndConeThickness

    ' determine whether cones will be created.
    If dStartConeLength < distTol Then      'if no cone length, then cone dia = tube dia
        dDiameterStart = dDiameterTube
    End If
    If dEndConeLength < distTol Then
        dDiameterEnd = dDiameterTube
    End If

    If (Abs(dDiameterTube - dDiameterStart)) > distTol Then
        bHasStartCone = True
    Else
        bHasStartCone = False
    End If
    If (Abs(dDiameterTube - dDiameterEnd)) > distTol Then
        bHasEndCone = True
    Else
        bHasEndCone = False
    End If

    If bHasStartCone = False Then
        dStartConeThickness = dTubeThickness
    End If
    If bHasEndCone = False Then
        dEndConeThickness = dTubeThickness
    End If

    ' we are constructing the curve to revolve as the inner curve, and we subtract the thicknesses.
    ' to keep the cone angle same, the thickness of the cone is adjusted by the slope of the cone
    ' The adjusted thickness is the cone thickness measured perp to the BUCan axis.
    ' Note that the adjusted thickness needs to use the ID to keep the slope true.  See TR# 167837.

    Dim dApproxAdjustedThickness As Double
    If bHasStartCone Then
        ' compute an adjusted cone thickness using OD's
        dApproxAdjustedThickness = AdjustThicknessForAngle(dStartConeLength, dDiameterTube, dDiameterStart, dStartConeThickness)
        ' compute adjusted cone thickness now using ID's, including the adjusted thickness on the end-of-cone side
        dApproxAdjustedThickness = AdjustThicknessForAngle(dStartConeLength, dDiameterTube - 2# * dTubeThickness, dDiameterStart - 2 * dApproxAdjustedThickness, dStartConeThickness)
        ' compute final adjusted cone thickness using ID's, including the re-adjusted thickness on the end-of-cone side.  diff is about 1e-8
        dStartConeThickness = AdjustThicknessForAngle(dStartConeLength, dDiameterTube - 2# * dTubeThickness, dDiameterStart - 2 * dApproxAdjustedThickness, dStartConeThickness)
    End If
    If bHasEndCone Then
        dApproxAdjustedThickness = AdjustThicknessForAngle(dEndConeLength, dDiameterTube, dDiameterEnd, dEndConeThickness)
        dApproxAdjustedThickness = AdjustThicknessForAngle(dEndConeLength, dDiameterTube - 2# * dTubeThickness, dDiameterEnd - 2 * dApproxAdjustedThickness, dEndConeThickness)
        dEndConeThickness = AdjustThicknessForAngle(dEndConeLength, dDiameterTube - 2# * dTubeThickness, dDiameterEnd - 2 * dApproxAdjustedThickness, dEndConeThickness)
    End If
    
    'TODO:  check for:
    ' coneLengths whose sum is larger than axisLength
    ' dDiameterStart < dStartConeThickness
    ' dDiameterEnd < dEndConeThickness
    ' handle case of no cylindrical portion

    ' alter diameters so curve to revolve becomes the inner curve, not outer.
    dRadiusTube = 0.5 * dDiameterTube - dTubeThickness
    dRadiusStart = 0.5 * dDiameterStart - dStartConeThickness
    dRadiusEnd = 0.5 * dDiameterEnd - dEndConeThickness
    
    If dRadiusTube < distTol Then
        Err.Raise E_FAIL
    End If
    If dRadiusStart < distTol Then
        Err.Raise E_FAIL
    End If
    If dRadiusEnd < distTol Then
        Err.Raise E_FAIL
    End If

    Dim dLengthExt As Double
    m_oDesignMemberHelper.GetLengthExtension oSmartOcc, dLengthExt
    If dLengthExt <= 0.0001 Then
        dLengthExt = 0.1
    End If

    For ii = 1 To 6
        Set oPts(ii) = New DPosition
    Next ii

    ' overLength makes the tube portion extend beyond the system axis so it is bounded by a plane created by the CanRule.
    ' only exists on EndCan.
    Dim dOverLength As Double
    dOverLength = dLength - dAxisLength        ' overLength is only applied on ends without a cone.  will be zero for non-rule-based can
    
    If bHasStartCone = True And bHasEndCone = True Then     ' need six points

        If dStartConeLength + dEndConeLength + distTol < dAxisLength Then       ' will have a cylindrical part
            oPts(1).Set -dLengthExt, dRadiusStart, 0#                ' at extension length
            oPts(2).Set 0, dRadiusStart, 0#                          ' inside corner of cone1
            oPts(3).Set dStartConeLength, dRadiusTube, 0#            ' outside corner of cone1
            oPts(4).Set dAxisLength - dEndConeLength, dRadiusTube, 0#    ' outside corner of cone2
            oPts(5).Set dAxisLength, dRadiusEnd, 0#                      ' inside corner of cone2
            oPts(6).Set dAxisLength + dLengthExt, dRadiusEnd, 0#         ' at extension length
            
            ' tweak the inside corners along the cone direction.  move corresponding extension points also.
            Set oDir = oPts(2).Subtract(oPts(3))      ' vector along cone, toward start
            oDir.Length = distTweak
            Set oPts(1) = oPts(1).Offset(oDir)
            Set oPts(2) = oPts(2).Offset(oDir)
            Set oDir = oPts(5).Subtract(oPts(4))      ' vector along cone, toward end
            oDir.Length = distTweak
            Set oPts(5) = oPts(5).Offset(oDir)
            Set oPts(6) = oPts(6).Offset(oDir)
            nPoints = 6
        
        ElseIf dStartConeLength + dEndConeLength - distTol < dAxisLength Then       ' will not have a cylindrical part
            oPts(1).Set -dLengthExt, dRadiusStart, 0#                ' at extension length
            oPts(2).Set 0, dRadiusStart, 0#                          ' inside corner of cone1
            oPts(3).Set dStartConeLength, dRadiusTube, 0#            ' outside corner of cone1
            oPts(4).Set dAxisLength, dRadiusEnd, 0#                      ' inside corner of cone2
            oPts(5).Set dAxisLength + dLengthExt, dRadiusEnd, 0#         ' at extension length
            
            ' tweak the inside corners along the cone direction.  move corresponding extension points also.
            Set oDir = oPts(2).Subtract(oPts(3))      ' vector along cone, toward start
            oDir.Length = distTweak
            Set oPts(1) = oPts(1).Offset(oDir)
            Set oPts(2) = oPts(2).Offset(oDir)
            Set oDir = oPts(4).Subtract(oPts(3))      ' vector along cone, toward end
            oDir.Length = distTweak
            Set oPts(4) = oPts(4).Offset(oDir)
            Set oPts(5) = oPts(5).Offset(oDir)
            nPoints = 5

        Else                                    ' sorry.  coneLengths exceed the axis length.
            Err.Raise E_FAIL
        End If
    
    ElseIf bHasStartCone Then

        oPts(1).Set -dLengthExt, dRadiusStart, 0#                ' at extension
        oPts(2).Set 0, dRadiusStart, 0#                          ' inside corner of cone1
        oPts(3).Set dStartConeLength, dRadiusTube, 0#            ' outside corner of cone1
        If dOverLength > distTol Then                               ' end of tube
            oPts(4).Set dLength + distTweak, dRadiusTube, 0#
        Else
            oPts(4).Set dAxisLength + dLengthExt, dRadiusTube, 0#
        End If
        
        Set oDir = oPts(2).Subtract(oPts(3))      ' vector along cone, toward start
        oDir.Length = distTweak
        Set oPts(1) = oPts(1).Offset(oDir)
        Set oPts(2) = oPts(2).Offset(oDir)
        nPoints = 4

    ElseIf bHasEndCone Then

        If dOverLength > distTol Then
            oPts(1).Set -(dOverLength + distTweak), dRadiusTube, 0#
        Else
            oPts(1).Set -dLengthExt, dRadiusTube, 0#
        End If
        oPts(2).Set dAxisLength - dEndConeLength, dRadiusTube, 0#
        oPts(3).Set dAxisLength, dRadiusEnd, 0#
        oPts(4).Set dAxisLength + dLengthExt, dRadiusEnd, 0#

        Set oDir = oPts(3).Subtract(oPts(2))      ' vector along cone, toward end
        oDir.Length = distTweak
        Set oPts(3) = oPts(3).Offset(oDir)
        Set oPts(4) = oPts(4).Offset(oDir)
        nPoints = 4

    Else
        ' with no cone, to which end to I apply the dOverLength ?
        If dOverLength > distTol Then
            Dim oMP As ISPSMemberPartCommon
            Dim oSPSPort As ISPSAxisEndPort
            Dim oPoint As IJPoint
            
            Set oMP = oSmartOcc
            Set oSPSPort = oMP.AxisPort(SPSMemberAxisStart)
            Set oPoint = oSPSPort.ILC
            
            ' FC at start means apply the overlength at the start.
            If TypeOf oPoint Is ISPSFrameConnection Then
                oPts(1).Set -(dOverLength + distTweak), dRadiusTube, 0#
                oPts(2).Set dAxisLength + dLengthExt, dRadiusTube, 0#
            Else
                oPts(1).Set -dLengthExt, dRadiusTube, 0#
                oPts(2).Set dLength + distTweak, dRadiusTube, 0#
            End If

        Else
            oPts(1).Set -dLengthExt, dRadiusTube, 0#
            oPts(2).Set dAxisLength + dLengthExt, dRadiusTube, 0#
        End If
        Set oDir = oPts(1).Subtract(oPts(2))      ' dummy subtract, just to create the DVector
        nPoints = 2
    
    End If

    ' Now transform the points locally, and send to the linestring.
    ' Then transform the linestring to global coordinates.
    Dim oMatrix As IJDT4x4
    Dim pts(1 To 18) As Double                  ' load a points array to send the linestring object
    Dim dDiameterMax As Double
    Dim oLineString3d As LineString3d
    
    dDiameterMax = GetMaxDiameter(dDiameterTube, dDiameterStart, dDiameterEnd)
    
    m_oDesignMemberHelper.CreateTransform oSmartOcc, dDiameterMax, dDiameterMax, oMatrix
    
    ' offset the points in local coordindates to make it relative to the center of the coordinate space
    ' move it by zero in X, -dDiameterMax / 2 in Y, dDiameterMax / 2 in Z

    oDir.Set 0#, -dDiameterMax / 2, dDiameterMax / 2        ' offset toward center of local coordinate space
    
    'offset the points locally, and load the pts array to send to the lineString
    jj = 0
    For ii = 1 To nPoints
        Set oPts(ii) = oPts(ii).Offset(oDir)
        jj = jj + 1
        pts(jj) = oPts(ii).x
        jj = jj + 1
        pts(jj) = oPts(ii).y
        jj = jj + 1
        pts(jj) = oPts(ii).z
    Next ii

    Set oLineString3d = pObject
    oLineString3d.SetPoints nPoints, pts
    oLineString3d.Transform oMatrix

    If True = bDebugPoints Then
        Dim oControlFlags As IJControlFlags
        Set oControlFlags = pObject
        oControlFlags.ControlFlags(&H4) = 0
        Dim oGFact As GeometryFactory
        Set oGFact = New GeometryFactory
        Dim oCurveElem As IJElements
        Set oCurveElem = New imscorecollections.JObjectCollection
        oCurveElem.Add oLineString3d
        oGFact.GeometryServices.DebugPrint oCurveElem, Environ("TEMP") & "\BUCanRevCurve.txt"
    End If
    
    For ii = 1 To 6
        Set oPts(ii) = Nothing
    Next ii
    
    Exit Sub

ErrorHandler:
    HandleError MODULE, METHOD
    SPSToDoErrorNotify "StructBUToDoMessages", 14, oSmartOcc, Nothing
    Err.Raise E_FAIL
End Sub

Public Sub CMConstructAxisOfRevolution(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObj As Object)
Const METHOD = "CMConstructAxisOfRevolution"
On Error GoTo ErrorHandler
    
    ' Just Create a initial dummy Curve
    Dim pCurve As IJCurve
    m_oDesignMemberHelper.CreateCurveBy2Points pResourceManager, 0, 0, 0, 1, 0, 0, pCurve
    CopyPermissionGroup pCurve, pMemberDescription.CAO
    
    Set pObj = pCurve
    
    Dim oControlFlags As IJControlFlags
    Set oControlFlags = pObj
    oControlFlags.ControlFlags(&H4) = &H4
    
Exit Sub
ErrorHandler:  HandleError MODULE, METHOD
End Sub

Public Sub CMComputeAxisOfRevolution(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
Const METHOD = "CMComputeAxisOfRevolution"
On Error GoTo ErrorHandler

    Dim oMatrix As IJDT4x4
    Dim dDiameterStart As Double
    Dim dDiameterEnd As Double
    Dim dStartConeLength As Double
    Dim dEndConeLength As Double
    Dim dTubeDiameter As Double
    Dim dLength As Double
    Dim dAxisLength As Double
    Dim eCanType As Long
    Dim oSmartOcc As IJSmartOccurrence
    Set oSmartOcc = pPropertyDescriptions.CAO
    
    'validate attributes
    If Not AreOccurrencePropertiesValid(oSmartOcc) Then
        GoTo ErrorHandler
    End If
    
    
    GetBuiltUpCanExtrusionParameters oSmartOcc, _
                                    dDiameterStart, _
                                    dDiameterEnd, _
                                    dStartConeLength, _
                                    dEndConeLength, _
                                    dTubeDiameter, _
                                    dLength, _
                                    dAxisLength, _
                                    eCanType

    Dim dDiameterMax As Double
    Dim pIJLine As IJLine
    
    dDiameterMax = GetMaxDiameter(dTubeDiameter, dDiameterStart, dDiameterEnd)
        
    m_oDesignMemberHelper.CreateTransform oSmartOcc, dDiameterMax, dDiameterMax, oMatrix

    Set pIJLine = pObject
    
    pIJLine.DefineBy2Points 0, -dDiameterMax / 2, dDiameterMax / 2, dLength, -dDiameterMax / 2, dDiameterMax / 2
    pIJLine.Transform oMatrix
    
    Exit Sub

ErrorHandler:
    HandleError MODULE, METHOD
    SPSToDoErrorNotify "StructBUToDoMessages", 14, oSmartOcc, Nothing
    Err.Raise E_FAIL
End Sub

Public Sub CMConstructBuiltUpCanPlateSystem(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObj As Object)
Const METHOD = "CMConstructBuiltUpCanPlateSystem"
On Error GoTo ErrorHandler
    
    Dim oSmartOcc As IJSmartOccurrence
    Dim oSmartItem As IJSmartItem
    Dim dThickness As Double
    Dim strMaterial As String
    Dim strGrade As String
   
    Dim oAttrCol As IJDAttributesCol
    Dim oAttr As IJDAttributes
   
    Set oSmartOcc = pMemberDescription.CAO
    Set oSmartItem = oSmartOcc.ItemObject
    
    Set oAttr = oSmartItem
    Set oAttrCol = oAttr.CollectionOfAttributes(IID_IUABuiltUpTube)
    If Not oAttrCol Is Nothing Then
        dThickness = oAttrCol.Item("TubeThickness").Value
        strMaterial = oAttrCol.Item("TubeMaterial").Value
        strGrade = oAttrCol.Item("TubeGrade").Value
    Else
        GoTo ErrorHandler
    End If
    
    'initialize the tube-thickness attribute if it was not declared as occ-attr
    Set oAttr = oSmartOcc
    Set oAttrCol = oAttr.CollectionOfAttributes(IID_IUABuiltUpTube)
    If Not oAttrCol Is Nothing Then
        If oAttrCol.Item("TubeThickness").Value < 0.000001 Then ' not yet set.
            oAttrCol.Item("TubeThickness").Value = dThickness
        End If
    Else
        GoTo ErrorHandler
    End If
    
    ' Create the curves to extrude
    Dim pCurveToRevolve As IJCurve
    Dim pAxisCurve As IJCurve
    Dim pIJDMemberObjects As IJDMemberObjects
    
    Set pIJDMemberObjects = oSmartOcc
    
    Set pCurveToRevolve = pIJDMemberObjects.ItemByDispid(1)
    Set pAxisCurve = pIJDMemberObjects.ItemByDispid(2)
    
    Set pObj = m_oDesignMemberHelper.CreateARevolvedBuiltUpPlate(pResourceManager, pCurveToRevolve, _
                                                                pAxisCurve, dThickness, _
                                                                strMaterial, strGrade, oSmartOcc, TubePlate, _
                                                                , , , OutDir, MoldedFormPlateNameCat.NameCatCan)
    m_oDesignMemberHelper.SetPlateBoundaries oSmartOcc, pObj
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
    Err.Raise E_FAIL
End Sub

Public Sub CMFinalConstructBuiltUpCanPlateSystem(pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructBuiltUpCanPlateSystem"
    On Error GoTo ErrorHandler
     
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
    Err.Raise E_FAIL
End Sub
Public Sub CMComputeBuiltUpCanPlateSystem(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
Const METHOD = "CMComputeBuiltUpCanPlateSystem"
On Error GoTo ErrorHandler
    ' get the  plate
    Dim oSmartOcc As IJSmartOccurrence
    Dim oSmartItem As IJSmartItem
    Dim pIJDMemberObjects As IJDMemberObjects

    Dim oTubePlate As IJPlate
    Dim oAttrCol As IJDAttributesCol
    Dim oAttr As IJDAttributes
   
    Dim dThickness As Double
    Dim strMaterial As String
    Dim strGrade As String
   
    Set oSmartOcc = pPropertyDescriptions.CAO
    Set oSmartItem = oSmartOcc.ItemObject
    
    Set pIJDMemberObjects = oSmartOcc
    Set oTubePlate = pIJDMemberObjects.ItemByDispid(Tube)
    
    Set oAttr = oSmartOcc
    Set oAttrCol = oAttr.CollectionOfAttributes(IID_IUABuiltUpTube)
    If Not oAttrCol Is Nothing Then
        dThickness = oAttrCol.Item("TubeThickness").Value
        strMaterial = oAttrCol.Item("TubeMaterial").Value
        strGrade = oAttrCol.Item("TubeGrade").Value
    Else
        GoTo ErrorHandler
    End If
    
    ' apply the new parameters
    With m_oDesignMemberHelper
        .SetMaterialAndGrade oTubePlate, strMaterial, strGrade
        .SetPlateDimensions oTubePlate, dThickness
    End With
    
    Exit Sub

ErrorHandler:
    HandleError MODULE, METHOD
    Err.Raise E_FAIL
End Sub
Public Sub CMMigrateBuiltUpCanPlateSystem(pMemberDesc As IJDMemberDescription, pMigrateHelper As IJMigrateHelper)
Const METHOD = "CMMigrateBuiltUpCanPlateSystem "
On Error GoTo ErrorHandler
    
    Dim oSmartOcc As IJSmartOccurrence
    Dim oSmartItem As IJSmartItem
    Dim oAttrCol As IJDAttributesCol
    Dim oAttr As IJDAttributes
    Dim dTubeThickness As Double
    Dim strTubeMaterial As String
    Dim strTubeGrade As String
    
    Dim dCone1Thickness As Double
    Dim strCone1Material As String
    Dim strCone1Grade As String
    Dim dCone2Thickness As Double
    Dim strCone2Material As String
    Dim strCone2Grade As String

    Set oSmartOcc = pMemberDesc.CAO
    Set oSmartItem = oSmartOcc.ItemObject
    
    Set oAttr = oSmartOcc
    Set oAttrCol = oAttr.CollectionOfAttributes(IID_IUABuiltUpTube)
    If Not oAttrCol Is Nothing Then
        dTubeThickness = oAttrCol.Item("TubeThickness").Value
        strTubeMaterial = oAttrCol.Item("TubeMaterial").Value
        strTubeGrade = oAttrCol.Item("TubeGrade").Value
    Else
        GoTo ErrorHandler
    End If

    If dTubeThickness < 0.000001 Then           ' thickness not set as OA, use the catalog attribute
        Set oAttr = oSmartItem
        Set oAttrCol = oAttr.CollectionOfAttributes(IID_IUABuiltUpTube)
        If Not oAttrCol Is Nothing Then
            dTubeThickness = oAttrCol.Item("TubeThickness").Value
        Else
            GoTo ErrorHandler
        End If
    End If

    Set oAttr = oSmartOcc
    Set oAttrCol = oAttr.CollectionOfAttributes(IID_IUABuiltUpCone1)
    If Not oAttrCol Is Nothing Then
        dCone1Thickness = oAttrCol.Item("Cone1Thickness").Value
        strCone1Material = oAttrCol.Item("Cone1Material").Value
        strCone1Grade = oAttrCol.Item("Cone1Grade").Value
    Else
        ' cones do not have a distinct attribute set for this
        Exit Sub
    End If
'    MsgBox "Thickness: " & CStr(dCone1Thickness) & vbCrLf _
'         & "Material:  " & strCone1Material & vbCrLf _
'         & "Grade:     " & strCone1Grade, vbInformation, "Start Cone Properties"
    Set oAttrCol = oAttr.CollectionOfAttributes(IID_IUABuiltUpCone2)
    If Not oAttrCol Is Nothing Then
        dCone2Thickness = oAttrCol.Item("Cone2Thickness").Value
        strCone2Material = oAttrCol.Item("Cone2Material").Value
        strCone2Grade = oAttrCol.Item("Cone2Grade").Value
    Else
        ' cones do not have a distinct attribute set for this
        Exit Sub
    End If
    
'    MsgBox "Thickness: " & CStr(dCone2Thickness) & vbCrLf _
'     & "Material:  " & strCone2Material & vbCrLf _
'     & "Grade:     " & strCone2Grade, vbInformation, "End Cone Properties"

   
       
    Dim pIJDMemberObjects As IJDMemberObjects
    Set pIJDMemberObjects = oSmartOcc
    
    Dim oCanAxis As IJLine
    Dim oAxisDirection As IJDVector
    
    Set oCanAxis = pIJDMemberObjects.ItemByDispid(2)
    Set oAxisDirection = GetAxisDirection(oCanAxis)
    
    Dim oPlate As IJPlate
    Set oPlate = pIJDMemberObjects.ItemByDispid(Tube)
    
    Dim oPlateUtils As IJPlateAttributes
    Set oPlateUtils = New PlateUtils
    Dim oSurf As IJSurfaceBody
    
    Dim oLeafPlateCollection As Collection
    Set oLeafPlateCollection = oPlateUtils.GetSplitResults(oPlate)
    
    Dim oTubePlates As Collection
    Dim oConePlates As Collection
    Set oTubePlates = New Collection
    Set oConePlates = New Collection
'    MsgBox oLeafPlateCollection.Count & " Leaf Plates"
    If oLeafPlateCollection.Count > 0 Then
        Dim oLeafPlate As IJPlate
        Dim dNormAxisDot
        Dim oSurfaceNormal As IJDVector
        ' sort the plates into tubes vs. cones
        For Each oLeafPlate In oLeafPlateCollection
            Set oSurf = oLeafPlate
            Set oSurfaceNormal = GetSurfaceNormal(oSurf)
            dNormAxisDot = oSurfaceNormal.Dot(oAxisDirection)
            Set oSurfaceNormal = Nothing
      ' let each leaf plate decide if it needs to
      ' set the material and grade
            If Abs(dNormAxisDot) < 0.000001 Then
                ' this is a tube
                oTubePlates.Add oLeafPlate
            Else
                oConePlates.Add oLeafPlate
            End If
            
        Next
        ' set the tube properties
        For Each oLeafPlate In oTubePlates
            m_oDesignMemberHelper.SetMaterialAndGrade oLeafPlate, strTubeMaterial, strTubeGrade
            m_oDesignMemberHelper.SetPlateDimensions oLeafPlate, dTubeThickness
        Next
        Set oTubePlates = Nothing
        Dim bConePropertiesIdentical As Boolean
        bConePropertiesIdentical = (Abs(dCone1Thickness - dCone2Thickness) <= 0.000001) And _
                                  (strCone1Material = strCone2Material) And _
                                   (strCone1Grade = strCone2Grade)
        If bConePropertiesIdentical Then
            For Each oLeafPlate In oConePlates
                m_oDesignMemberHelper.SetMaterialAndGrade oLeafPlate, strCone1Material, strCone1Grade
                m_oDesignMemberHelper.SetPlateDimensions oLeafPlate, dCone1Thickness
            Next
        Else
            ' have to distinguish start cones from end cones
            ' get the length of the start cone (may be zero)
            Dim dStartConeLength  As Double
            Set oAttr = oSmartOcc
            Set oAttrCol = oAttr.CollectionOfAttributes(IID_IUABuiltUpCan)
            If Not oAttrCol Is Nothing Then
                dStartConeLength = oAttrCol.Item("LengthStartCone").Value
            End If
            ' get the start point
            Dim oStartPos As IJDPosition
            Dim sX As Double
            Dim sY As Double
            Dim sZ As Double
            oCanAxis.GetStartPoint sX, sY, sZ
            Set oStartPos = New DPosition
            oStartPos.Set sX, sY, sZ
            Dim dDirLen As Double
            Debug.Print "Start Position: (" & CStr(sX) & ", " & CStr(sY) & ", " & CStr(sZ) & ")"
            Dim oDirTowardStart As IJDVector
            Dim oLCS As IJLocalCoordinateSystem
            For Each oLeafPlate In oConePlates
                Set oLCS = oLeafPlate
                Set oDirTowardStart = oLCS.Position.Subtract(oStartPos)
                dDirLen = oDirTowardStart.Dot(oAxisDirection)
                If dDirLen <= dStartConeLength Then
                    ' this is a start cone
                    m_oDesignMemberHelper.SetMaterialAndGrade oLeafPlate, strCone1Material, strCone1Grade
                    m_oDesignMemberHelper.SetPlateDimensions oLeafPlate, dCone1Thickness
                Else
                    m_oDesignMemberHelper.SetMaterialAndGrade oLeafPlate, strCone2Material, strCone2Grade
                    m_oDesignMemberHelper.SetPlateDimensions oLeafPlate, dCone2Thickness
                End If
               Set oDirTowardStart = Nothing
            Next
        End If
        
        Set oConePlates = Nothing
        
    End If
    
    Exit Sub

ErrorHandler:
    HandleError MODULE, METHOD
    Err.Raise E_FAIL
End Sub

Public Sub CMFinalConstructAsm(pAggregatorDescription As IJDAggregatorDescription)
Const METHOD = "CMFinalConstructAsm"
On Error GoTo ErrorHandler
    
  
Exit Sub
ErrorHandler: HandleError MODULE, METHOD
End Sub

Public Sub CMConstructAsm(pAggregatorDescription As IJDAggregatorDescription)
Const METHOD = "CMConstructAsm"
On Error GoTo ErrorHandler

  
Exit Sub
ErrorHandler: HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateCAO(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
Const METHOD = "CMEvaluateCAO"
On Error GoTo ErrHandler
    
     CalcuteOutputParameters pPropertyDescriptions.CAO

Exit Sub
ErrHandler: HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateCAO1(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
Const METHOD = "CMEvaluateCAO1"
On Error GoTo ErrHandler
    
    Dim oSmartOcc As IJSmartOccurrence
    Dim oDesMem As ISPSDesignedMember
    
    Set oSmartOcc = pPropertyDescriptions.CAO
    Set oDesMem = oSmartOcc
    
    oDesMem.NotifyDesignChange
    
Exit Sub
ErrHandler: HandleError MODULE, METHOD
End Sub

Public Sub ISPSDesignedMemberHelper_GetNominalSectionSize(ByVal pIDesignedMember As ISPSDesignedMember, _
                                                          ByVal pPosAlong As IJDPosition, _
                                                          ByRef pdWidth As Double, _
                                                          ByRef pdDepth As Double)

Const METHOD = "ISPSDesignedMemberHelper_GetNominalSectionSize"
On Error GoTo ErrHandler
        
   
    Dim dDiameterStart As Double
    Dim dDiameterEnd As Double
    Dim dStartConeLength As Double
    Dim dEndConeLength As Double
    Dim dTubeDiameter As Double
    Dim dLength As Double
    Dim dAxisLength As Double
    Dim eCanType As Long
    Dim oSmartOcc As IJSmartOccurrence
    Set oSmartOcc = pIDesignedMember
    
    GetBuiltUpCanExtrusionParameters oSmartOcc, _
                                     dDiameterStart, _
                                     dDiameterEnd, _
                                     dStartConeLength, _
                                     dEndConeLength, _
                                     dTubeDiameter, _
                                     dLength, _
                                     dAxisLength, _
                                     eCanType

    
    If Not pPosAlong Is Nothing Then
        Dim pISPSMemberPartCommon As ISPSMemberPartCommon
        Set pISPSMemberPartCommon = pIDesignedMember
        Dim axisCurve As IJCurve
        Set axisCurve = pISPSMemberPartCommon.Axis

        ' get the parameter of the input position
        Dim posParam As Double
        axisCurve.Parameter pPosAlong.x, pPosAlong.y, pPosAlong.z, posParam
        ' get the parameter range
        Dim startParam As Double
        Dim endParam As Double
        
        axisCurve.ParamRange startParam, endParam
        
        If posParam < startParam Then
            posParam = startParam
        ElseIf posParam > endParam Then
            posParam = endParam
        End If
'
'         ____________________________
'        /                            \
'       /                              \
'      /                                \
'     /                                  \
'
'    A-a-B----------------------------C-b-D
'
'     \                                  /
'      \                                /
'       \                              /
'        \____________________________/
        
' A->B = Start Cone Length
' B->C = Tube Length
' C->D = End Cone Length
' A->D = Overall Length
' the ratio for point 'a' (on the start cone) is A->a/A->B
' the ratio for point 'b' (on the end cone) is D->b/C->D
        
        Dim dPositionRatio As Double
        dPositionRatio = posParam / (endParam - startParam)
        
        Dim bOnStartCone As Boolean

        Dim bOnEndCone As Boolean
        
        Dim lengthToPosition As Double
        lengthToPosition = dPositionRatio * dAxisLength
        bOnStartCone = lengthToPosition < dStartConeLength
        
        bOnEndCone = lengthToPosition > (dAxisLength - dEndConeLength)
        
        Dim diameterDiff As Double
        'if bOnStartCone is true,dDiameterStart should never be zero (or close to it)
        ' if it is for some reason, it means that we think the point lies on the cone
        ' but there is no cone, so the point must be on the tube
        If bOnStartCone And dStartConeLength > 0.001 Then
            diameterDiff = dTubeDiameter - dDiameterStart
            dPositionRatio = lengthToPosition / dStartConeLength
            pdWidth = dDiameterStart + (diameterDiff * dPositionRatio)

        ElseIf bOnEndCone And dEndConeLength > 0.001 Then
            diameterDiff = dTubeDiameter - dDiameterEnd
            dPositionRatio = (dAxisLength - lengthToPosition) / dEndConeLength
            pdWidth = dDiameterEnd + (diameterDiff * dPositionRatio)

        Else
            pdWidth = dTubeDiameter

        End If
    Else
        pdWidth = GetMaxDiameter(dTubeDiameter, dDiameterStart, dDiameterEnd)
    End If
    
    pdDepth = pdWidth
   
Exit Sub
ErrHandler: HandleError MODULE, METHOD
End Sub

Private Function IJUserAttributeMgmt_OnAttributeChange(ByVal pIJDAttrs As IJDAttributes, ByVal CollAllDisplayedValues As Object, ByVal pAttrToChange As IJAttributeDescriptor, ByVal varNewAttrValue As Variant) As String
Const METHOD = "IJUserAttributeMgmt_OnAttributeChange"
On Error GoTo ErrHandler

    ' Validate the attribute new value first before any further processing
    Dim ErrStr As String

    Dim oAttrCol As IJDInfosCol
    Set oAttrCol = Nothing
    Set oAttrCol = m_oDesignMemberHelper.GetInfosCollection(pIJDAttrs, pAttrToChange.InterfaceName, pAttrToChange.AttrName)

    If oAttrCol Is Nothing Then
        GoTo ErrHandler
    End If
    
    Dim oAttrObj As IJDAttributeInfo
    Dim oInterfaceInfo As IJDInterfaceInfo
    Dim AttrCount As Long
    Dim AttrType As Long
    
    ' loop on the attributes on the interface to match the supplied attribute type
    For AttrCount = 1 To oAttrCol.Count
        Set oAttrObj = oAttrCol.Item(AttrCount)
        If oAttrObj.Name = pAttrToChange.AttrName Then
            Select Case oAttrObj.Type
                Case m_oDesignMemberHelper.DoubleValue
                    ErrStr = BuiltUpDefValidate(pAttrToChange.AttrName, varNewAttrValue, 0#)
                    If Len(ErrStr) > 0 Then
                        IJUserAttributeMgmt_OnAttributeChange = ErrStr
                        Exit Function
                    End If
                    
                    Dim oDesMem As ISPSDesignedMember
                    Set oDesMem = pIJDAttrs
                    If Not oDesMem Is Nothing And IsDesignParameter(pAttrToChange) Then
                        oDesMem.NotifyDesignChange
                    End If
                End Select
        End If
    Next
    
    IJUserAttributeMgmt_OnAttributeChange = ""
   
Exit Function
ErrHandler:
    IJUserAttributeMgmt_OnAttributeChange = m_oLocalizer.GetString(IDS_BUILTUP_ERROR, "ERROR")
    HandleError MODULE, METHOD
End Function

Private Function IJUserAttributeMgmt_OnPreCommit(ByVal pIJDAttrs As IJDAttributes, ByVal CollAllDisplayedValues As Object) As String
Const METHOD = "IJUserAttributeMgmt_OnPreCommit"
On Error GoTo ErrHandler

    IJUserAttributeMgmt_OnPreCommit = ""
    
Exit Function
ErrHandler: HandleError MODULE, METHOD
End Function

Private Function IJUserAttributeMgmt_OnPreLoad(ByVal pIJDAttrs As IJDAttributes, ByVal CollAllDisplayedValues As Object) As String
Const METHOD = "IJUserAttributeMgmt_OnPreLoad"
On Error GoTo ErrHandler
    
    Dim oAttrCol As IJDAttributesCol
    Dim bIsModifiable As Boolean
    Dim i As Integer
    Dim pAttrColl As Collection
    Dim pAttrDescr As IJAttributeDescriptor
    
    ' Setup the Smart Item
    Dim oSmartOcc As IJSmartOccurrence
    Dim oSmartItem As IJSmartItem
    Set oSmartOcc = pIJDAttrs
    Set oSmartItem = oSmartOcc.ItemObject
    
    Dim oAttr As IJDAttributes
    Set oAttr = oSmartItem
    
    bIsModifiable = m_oDesignMemberHelper.IsAttributeModifiable(oAttr)
    
    If Not bIsModifiable Then
        Set pAttrColl = CollAllDisplayedValues
        For i = 1 To pAttrColl.Count
            Set pAttrDescr = pAttrColl.Item(i)
            If pAttrDescr.InterfaceName = IID_IUABuiltUpCan Then
                If pAttrDescr.AttrName = "DiameterStart" Or pAttrDescr.AttrName = "DiameterEnd" Then
                    pAttrDescr.AttrState = AttributeDescriptor_ReadOnly
                End If
            ElseIf pAttrDescr.InterfaceName = IID_IUABuiltUpTube Then
                If pAttrDescr.AttrName = "TubeDiameter" Then
                    pAttrDescr.AttrState = AttributeDescriptor_ReadOnly
                End If
            End If
        Next
    End If
    
    'set the common read only attributes
    m_oDesignMemberHelper.SetCrossSectionReadOnlyAttributesSpecCase pIJDAttrs, CollAllDisplayedValues
    
    Dim bSectionProperties As Boolean
    bSectionProperties = m_oDesignMemberHelper.AreSectionPropertiesModifiable(oSmartOcc)
    
    If bSectionProperties = False Then
        Set pAttrColl = CollAllDisplayedValues
        For i = 1 To pAttrColl.Count
            Set pAttrDescr = pAttrColl.Item(i)
            If pAttrDescr.InterfaceName = "IStructCrossSectionDesignProperties" _
                Or pAttrDescr.InterfaceName = "IStructCrossSectionUnitWeight" Then
                pAttrDescr.AttrState = AttributeDescriptor_ReadOnly
            End If
        Next
    End If
    
    IJUserAttributeMgmt_OnPreLoad = ""
    
Exit Function
ErrHandler: HandleError MODULE, METHOD
End Function
Private Sub Class_Initialize()
Set m_oLocalizer = New IMSLocalizer.Localizer
    m_oLocalizer.Initialize App.Path & "\" & "SPSDesignedMemberDefs"
    
    Set m_oCalcXProps = New BUCalcSectionProperties
    Set m_oDesignMemberHelper = New BUHelperUtils
End Sub

Private Sub Class_Terminate()
    Set m_oLocalizer = Nothing
    Set m_oCalcXProps = Nothing
    Set m_oDesignMemberHelper = Nothing
End Sub

Private Sub GetBuiltUpCanExtrusionParameters(ByVal oSmartOcc As IJSmartOccurrence, _
                                             ByRef dDiameterStart As Double, _
                                             ByRef dDiameterEnd As Double, _
                                             ByRef dStartConeLength As Double, _
                                             ByRef dEndConeLength As Double, _
                                             ByRef dTubeDiameter As Double, _
                                             ByRef dLength As Double, _
                                             ByRef dAxisLength As Double, _
                                             ByRef eCanType As Long)

Const METHOD = "GetBuiltUpCanExtrusionParameters"

On Error Resume Next
    
    Dim oAttrCol As IJDAttributesCol
    Dim oAttr As IJDAttributes

    Set oAttr = oSmartOcc

    ' Get Parameters of Tube
    Set oAttrCol = oAttr.CollectionOfAttributes(IID_IUABuiltUpTube)
    If Not oAttrCol Is Nothing Then
        dTubeDiameter = oAttrCol.Item("TubeDiameter").Value
    Else
        GoTo ErrorHandler
    End If
    
    ' Get Parameters of BU
    Set oAttrCol = oAttr.CollectionOfAttributes(IID_IUABuiltUpCan)
    If Not oAttrCol Is Nothing Then
        dDiameterEnd = oAttrCol.Item("DiameterEnd").Value
        dDiameterStart = oAttrCol.Item("DiameterStart").Value
        dStartConeLength = oAttrCol.Item("LengthStartCone").Value
        dEndConeLength = oAttrCol.Item("LengthEndCone").Value
    Else
        GoTo ErrorHandler
    End If
   
    ' the length may be set by the rule
    Set oAttrCol = oAttr.CollectionOfAttributes(IID_IJUASMCanRuleResult)
    If Not oAttrCol Is Nothing Then
        dLength = oAttrCol.Item("UncutLength").Value
        eCanType = oAttrCol.Item(attrCanType).Value
    End If
    Debug.Print "Can Length: " & CStr(dLength)
    If 0 = eCanType Then
        eCanType = CanType_InLine
    End If

    ' Get Length of input curve
    Dim pISPSMemberPartCommon As ISPSMemberPartCommon
    Set pISPSMemberPartCommon = oSmartOcc

    dAxisLength = pISPSMemberPartCommon.Axis.Length
    If dLength < 0.001 Then             ' presumably was not set by a CanRule
        dLength = dAxisLength
    End If

    Dim pCurveToRevolve As IJCurve
    Dim pAxisCurve As IJCurve
    Dim pIJDMemberObjects As IJDMemberObjects
    Set pIJDMemberObjects = oSmartOcc
    Set pCurveToRevolve = pIJDMemberObjects.ItemByDispid(1)
    Set pAxisCurve = pIJDMemberObjects.ItemByDispid(2)

    If pCurveToRevolve Is Nothing Or pAxisCurve Is Nothing Then              'not computed yet.
        Exit Sub
    End If

    ' do we need to initialize LengthStartCone, LengthEndCone parameters ?
    ' the answer is yes if either is < 0, or if newly created and sum of Start + End > AxisLength

    Dim iFlags As Long
    Const RELATION_INSERTED = &H100000
    Dim structAssocCompute As IJStructAssocCompute
    Dim bResetLengths As Boolean
    bResetLengths = False

    If dStartConeLength < 0 Then                            ' initial placement
        bResetLengths = True
    End If
    If dEndConeLength < 0 Then
        bResetLengths = True
    End If
    
    If Not bResetLengths Then
        
        Set structAssocCompute = New StructAssocTools
        structAssocCompute.GetAssocFlags oSmartOcc, iFlags
        
        If (iFlags And RELATION_INSERTED) = RELATION_INSERTED Then          ' inserted this compute
            
            If dStartConeLength + dEndConeLength > dAxisLength Then         ' axis smaller than sum of distances
                bResetLengths = True
            End If
        End If
    End If

    If bResetLengths Then

        Set oAttrCol = oAttr.CollectionOfAttributes(IID_IUABuiltUpCan)

        dStartConeLength = 0.333 * dAxisLength
        oAttrCol.Item("LengthStartCone").Value = dStartConeLength
        
        dEndConeLength = 0.333 * dAxisLength
        oAttrCol.Item("LengthEndCone").Value = dEndConeLength

    End If

    Exit Sub

ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Private Function GetMaxDiameter(ByVal dTubeDiameter As Double, _
                                ByVal dDiameterStart As Double, _
                                ByVal dDiameterEnd As Double) As Double
Const METHOD = "GetMaxDiameter"
On Error GoTo ErrorHandler

    Dim dDiameterMax As Double
    dDiameterMax = 0#
    
    If dTubeDiameter >= dDiameterStart And dTubeDiameter >= dDiameterEnd Then
        dDiameterMax = dTubeDiameter
    ElseIf dDiameterStart >= dTubeDiameter And dDiameterStart >= dDiameterEnd Then
        dDiameterMax = dDiameterStart
    Else
        dDiameterMax = dDiameterEnd
    End If
    
    GetMaxDiameter = dDiameterMax
        
    Exit Function
ErrorHandler:
    HandleError MODULE, METHOD
    Err.Raise E_FAIL
End Function
Public Function AreOccurrencePropertiesValid(ByVal oSmartOcc As IJSmartOccurrence) As Boolean
Const METHOD = "AreOccurrencePropertiesValid"

    On Error GoTo ErrorHandler
    AreOccurrencePropertiesValid = False
    
    Dim dDiameterStart As Double
    Dim dDiameterEnd As Double
    Dim dStartConeLength As Double
    Dim dEndConeLength As Double
    Dim dTubeDiameter As Double
    Dim dLength As Double
    Dim sError As String
    Dim eCanType As Long
    Dim dAxisLength As Double
    
    GetBuiltUpCanExtrusionParameters oSmartOcc, dDiameterStart, dDiameterEnd, _
                                     dStartConeLength, dEndConeLength, _
                                     dTubeDiameter, dLength, dAxisLength, eCanType
        
    sError = BuiltUpDefValidate("DiameterStart", dDiameterStart, 0#) ', dTubeDiameter)
    
    If Len(sError) > 0 Then
        AreOccurrencePropertiesValid = False
        Exit Function
    End If
    
    sError = BuiltUpDefValidate("DiameterEnd", dDiameterEnd, 0#) ', dTubeDiameter)
    
    If Len(sError) > 0 Then
        AreOccurrencePropertiesValid = False
        Exit Function
    End If
    
    sError = BuiltUpDefValidate("LengthStartCone", dStartConeLength, 0#)
    
    If Len(sError) > 0 Then
        AreOccurrencePropertiesValid = False
        Exit Function
    End If
    
    sError = BuiltUpDefValidate("LengthEndCone", dEndConeLength, 0#)
    
    If Len(sError) > 0 Then
        AreOccurrencePropertiesValid = False
        Exit Function
    End If
    
    sError = BuiltUpDefValidate("TubeDiameter", dTubeDiameter, 0#) 'm_oCalcXProps.Min(dDiameterStart, dDiameterEnd))
    
    If Len(sError) > 0 Then
        AreOccurrencePropertiesValid = False
        Exit Function
    End If
    
    sError = BuiltUpDefValidate("Length", dLength, 0#)
    
    If Len(sError) > 0 Then
        AreOccurrencePropertiesValid = False
        Exit Function
    End If
        
    Dim dLengthExt As Double
    m_oDesignMemberHelper.GetLengthExtension oSmartOcc, dLengthExt
    
    sError = BuiltUpDefValidate("LengthExt", dLengthExt, 0#)
    
    If Len(sError) > 0 Then
        AreOccurrencePropertiesValid = False
        Exit Function
    End If
  
    AreOccurrencePropertiesValid = True

    Exit Function
ErrorHandler:
    HandleError MODULE, METHOD
    Err.Raise E_FAIL
End Function

Private Sub CalcuteOutputParameters(ByVal oSmartOcc As IJSmartOccurrence)
Const METHOD = "CalcuteOutputParameters"
On Error GoTo ErrorHandler
   
    Dim oAttrCol As IJDAttributesCol
    Dim oAttr As IJDAttributes
    Dim oSmartItem As IJSmartItem
     
    Dim dDiameterStart As Double
    Dim dDiameterEnd As Double
    Dim dStartConeLength As Double
    Dim dEndConeLength As Double
    Dim dTubeDiameter As Double
    Dim dLength As Double
    Dim dTubeThickness As Double
    
    Dim dArea As Double
    Dim dPerimeter As Double
    Dim dCentroidX As Double
    Dim dCentroidY As Double
    Dim dXp As Double
    Dim dYp As Double
    Dim dIxx As Double
    Dim dIyy As Double
    Dim dRo As Double
    Dim dRxx As Double
    Dim dRxy As Double
    Dim dRyy As Double
    Dim dSw As Double
    Dim dSxx As Double
    Dim dSyy As Double
    Dim dZxx As Double
    Dim dZyy As Double
    Dim dJ As Double
    Dim dCw As Double
    Dim dH As Double
    Dim dUnitWt As Double
    Dim dAxisLength As Double
    Dim eCanType As Long
    'Initialize all parameters to Zero (0)
    dArea = dPerimeter = dCentroidX = dCentroidY = dXp = dYp = dIxx = dIyy = _
    dRo = dRxx = dRxy = dRyy = dSw = dSxx = dSyy = dZxx = dZyy = dJ = dCw = dH = dUnitWt = 0#
    
    GetBuiltUpCanExtrusionParameters oSmartOcc, dDiameterStart, dDiameterEnd, _
                                     dStartConeLength, dEndConeLength, _
                                     dTubeDiameter, dLength, dAxisLength, eCanType

    Dim dDiameterMax As Double
    dDiameterMax = GetMaxDiameter(dTubeDiameter, dDiameterStart, dDiameterEnd)
           
    Set oSmartItem = oSmartOcc.ItemObject
    Set oAttr = oSmartItem
    Set oAttrCol = oAttr.CollectionOfAttributes(IID_IUABuiltUpTube)
    If Not oAttrCol Is Nothing Then
        dTubeThickness = oAttrCol.Item("TubeThickness").Value
    Else
        GoTo ErrorHandler
    End If
    
    Dim ErrorStatus As ErrorSectionStatus
    ErrorStatus = GetSectionPropertiesFromCalculator(dArea, dPerimeter, dCentroidX, dCentroidY, _
                                                     dXp, dYp, dIxx, dIyy, dRo, dRxx, dRxy, dRyy, _
                                                     dSw, dSxx, dSyy, dZxx, dZyy, dJ, dCw, dH, dUnitWt, _
                                                     dDiameterMax, dTubeThickness)
                                                     
    If ErrorStatus <> SectionProperties_OK Then
        GoTo ErrorHandler
    End If

    Set oAttr = Nothing
    Set oAttrCol = Nothing
    Set oAttr = oSmartOcc
    Set oAttrCol = oAttr.CollectionOfAttributes(IStructCrossSectionDimensions)
    If Not oAttrCol Is Nothing Then
        oAttrCol.Item("Width").Value = dDiameterMax
        oAttrCol.Item("Depth").Value = dDiameterMax
        oAttrCol.Item("Area").Value = dArea
        oAttrCol.Item("Perimeter").Value = dPerimeter
    Else
        GoTo ErrorHandler
    End If
    
    Set oAttrCol = oAttr.CollectionOfAttributes(IStructCrossSectionDesignProperties)
    If Not oAttrCol Is Nothing Then
        On Error Resume Next
        oAttrCol.Item("IsHollow").Value = True
        oAttrCol.Item("IsSymmetricAboutX").Value = True
        oAttrCol.Item("IsSymmetricAboutY").Value = True
        On Error GoTo ErrorHandler
    End If
    
    Dim bIsSetOk As Boolean
    bIsSetOk = True
    
    ' set the cross section design properties if they are not user defined
    If (m_oDesignMemberHelper.AreSectionPropertiesModifiable(oAttr) = False) Then
        bIsSetOk = m_oDesignMemberHelper.SetXSectionDesignProperties(oSmartOcc, dCentroidX, dCentroidY, _
                                                                     dXp, dYp, dIxx, dIyy, dRo, dRxx, dRxy, dRyy, _
                                                                     dSw, dSxx, dSyy, dZxx, dZyy, dJ, dCw, dH, dUnitWt)
    End If
                                           
    ' if there was an error while setting the design properties go the the error handler
    If bIsSetOk = False Then
        GoTo ErrorHandler
    End If
      
Exit Sub
ErrorHandler: HandleError MODULE, METHOD
End Sub

Private Function BuiltUpDefValidate(sAttributeName As String, _
                                    varAttributeValue As Variant, _
                                    dLowRange As Double, _
                                    Optional dUpRange As Double = -1908) As String
                                    '-1908 is an arbritraty value any negative number would do
                                    '1908 is the last time the Chicago Cubs won a world series
                                    
Const METHOD = "BuiltUpDefValidate"
On Error GoTo ErrorHandler

    If dUpRange > 0 Then
         If (varAttributeValue < dLowRange) Or ((varAttributeValue - dUpRange) > m_oDesignMemberHelper.distTol) Then
            BuiltUpDefValidate = sAttributeName & ": " & m_oLocalizer.GetString(IDS_BUILTUP_VALUE_MUSTBE_IN_RANGE, _
                                 "Value must be within range of " & dLowRange & " and " & dUpRange) _
                                 & " [" & dLowRange & ", " & dUpRange & "]"
            Exit Function
        End If
    End If
    
    Select Case sAttributeName
        Case "DiameterStart", "DiameterEnd", "TubeDiameter", "Length"
            If varAttributeValue <= 0# Then
                BuiltUpDefValidate = sAttributeName & ": " & m_oLocalizer.GetString(IDS_BUILTUP_VALUE_MUSTBE_POSITIVE, "Value must be > 0")
                Exit Function
            End If
        Case Else
            If (varAttributeValue < dLowRange) Then
                BuiltUpDefValidate = sAttributeName & ": " & m_oLocalizer.GetString(IDS_BUILTUP_VALUE_MUSTBE_GREATERTHAN_OR_EQUAL_TO_ZERO, "Value must be >= 0")
                Exit Function
            End If
    End Select

Exit Function
ErrorHandler: HandleError MODULE, METHOD
End Function

Private Function GetSectionPropertiesFromCalculator(ByRef dArea As Double, _
                                                    ByRef dPerimeter As Double, _
                                                    ByRef dCentroidX As Double, _
                                                    ByRef dCentroidY As Double, _
                                                    ByRef dXp As Double, ByRef dYp As Double, _
                                                    ByRef dIxx As Double, ByRef dIyy As Double, _
                                                    ByRef dRo As Double, ByRef dRxx As Double, _
                                                    ByRef dRxy As Double, ByRef dRyy As Double, _
                                                    ByRef dSw As Double, ByRef dSxx As Double, _
                                                    ByRef dSyy As Double, ByRef dZxx As Double, _
                                                    ByRef dZyy As Double, ByRef dJ As Double, _
                                                    ByRef dCw As Double, dH As Double, _
                                                    ByRef dUnitWt As Double, _
                                                    ByVal dDiameterMax As Double, _
                                                    ByVal dTubeThickness As Double) As ErrorSectionStatus
                                               
    
Const METHOD = "GetSectionPropertiesFromCalculator"
On Error GoTo ErrorHandler
    
    Dim ErrorStatus As ErrorSectionStatus
    Dim bIsSetOk As Boolean
      
    m_oCalcXProps.SectionTypeAlias = EnumSectionTypeAlias.Section_Circular
    m_oCalcXProps.depth = dDiameterMax
    m_oCalcXProps.WebThickness = dTubeThickness

    ErrorStatus = m_oCalcXProps.SectionProperties(dArea, dPerimeter, dCentroidX, dCentroidY, _
                                                  dXp, dYp, dIxx, dIyy, dRo, dRxx, dRxy, dRyy, _
                                                  dSw, dSxx, dSyy, dZxx, dZyy, dJ, dCw, dH, dUnitWt)
                                                                         
Exit Function
ErrorHandler: HandleError MODULE, METHOD
ErrorStatus = SectionProperties_UnExpectedError
End Function

Public Sub ISPSDesignedMemberHelper_LoadEmulatedFacePorts(ByVal oDesignedMember As SPSMembers.ISPSDesignedMember)
Const METHOD = "ISPSDesignedMemberHelper_LoadEmulatedFacePorts"
On Error GoTo ErrorHandler

    Dim oAssyMembers     As IJDMemberObjects
    Dim oPlateSystem     As IJPlateSystem
    Dim eSectionAlias As EnumSectionTypeAlias
    
    eSectionAlias = Section_Circular
    'get assy members from SO
    Set oAssyMembers = oDesignedMember
    
   'handle the web
    Set oPlateSystem = oAssyMembers.Item(BUTubeMembers.Tube)
    If Not oPlateSystem Is Nothing Then
        m_oDesignMemberHelper.EvaluateWebPlateSystemXIDs oPlateSystem, StructBUPlateType.Tube, eSectionAlias, oDesignedMember
    End If
    
Exit Sub
ErrorHandler: HandleError MODULE, METHOD
End Sub

Public Sub ISPSDesignedMemberHelper_ResolveAmbiguity(ByVal pDesignedMember As SPSMembers.ISPSDesignedMember)
Const METHOD = "ISPSDesignedMemberHelper_ResolveAmbiguity"
On Error GoTo ErrorHandler
    'Nothing required for this particular Built-Up Definition
Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

' check lets the caller decide to take a certain action for changes in deaign parameters, like notification
' to the listeners
Private Function IsDesignParameter(ByVal oAttrDesc As IJAttributeDescriptor) As Boolean
Const METHOD = "IsDesignParameter"
On Error GoTo ErrorHandler

    Dim sAttrName As String
    Dim sIntfName As String
    
    If Not oAttrDesc Is Nothing Then
    
        IsDesignParameter = False
        sAttrName = oAttrDesc.AttrName
        sIntfName = oAttrDesc.InterfaceName
        
        Select Case sIntfName
            Case "IUABuiltUpTube"
                If sAttrName = "TubeDiameter" Then
                    IsDesignParameter = True
                End If
            Case "IUABuiltUpCan"
            'actualy don't neeed this logic below. All attributes on this I/F are design attrs but
            ' its an extra check in case user decides to add a non-design attribute on the same I/F
                Select Case sAttrName
                    Case "DiameterStart"
                        IsDesignParameter = True
                    Case "DiameterEnd"
                        IsDesignParameter = True
                    Case "LengthStartCone"
                        IsDesignParameter = True
                    Case "LengthEndCone"
                        IsDesignParameter = True
                    Case Else
                End Select ' attr name
            
            Case Else
        End Select ' I/F name
    End If


Exit Function
ErrorHandler:
    HandleError MODULE, METHOD
End Function
    
Private Function GetSurfaceNormal(ByVal oSurf As IJSurfaceBody) As IJDVector
    Const METHOD = "GetSurfaceNormal"
    On Error GoTo ErrorHandler
    
    Dim oPosOnSurf As IJDPosition
    Dim oCS As IJLocalCoordinateSystem
    Dim oSurfNorm As IJDVector

    oSurf.GetCenterOfGravity oPosOnSurf
    Set oCS = oSurf
    Set oPosOnSurf = oCS.Position

    oSurf.GetNormalFromPosition oPosOnSurf, oSurfNorm
    
'    Dim nX As Double
'    Dim nY As Double
'    Dim nZ As Double
'    oSurfNorm.Get nX, nY, nZ
'    MsgBox "X: " & nX & vbCrLf _
'          & "Y: " & nY & vbCrLf _
'          & "Z: " & nZ, vbInformation, "Surface Normal"
    Set GetSurfaceNormal = oSurfNorm

    Exit Function
ErrorHandler:
    HandleError MODULE, METHOD
End Function


Private Function GetAxisDirection(ByVal oAxis As IJLine) As IJDVector
Const METHOD = "GetAxisDirection"
On Error GoTo ErrorHandler
    Dim dStartParam As Double
    Dim dEndParam As Double
   
    Dim nX As Double
    Dim nY As Double
    Dim nZ As Double
    Dim oAxisDir As IJDVector
    Set oAxisDir = New DVector
    oAxis.GetDirection nX, nY, nZ
    oAxisDir.Set nX, nY, nZ
'    MsgBox "X: " & nX & vbCrLf _
'          & "Y: " & nY & vbCrLf _
'          & "Z: " & nZ, vbInformation, "Un normalized Axis Direction"
    oAxisDir.Length = 1
    Set GetAxisDirection = oAxisDir

    Exit Function

ErrorHandler:
    HandleError MODULE, METHOD
End Function

Private Sub GetThicknesses(ByRef oSmartOcc As IJSmartOccurrence, _
                           ByRef dTubeThickness As Double, _
                           ByRef dStartConeThickness As Double, _
                           ByRef dEndConeThickness As Double)
    Const METHOD = "GetThicknesses"
    On Error GoTo ErrorHandler
                           
    Dim oAttrCol As IJDAttributesCol
    Dim oAttr As IJDAttributes
   
    Set oAttr = oSmartOcc
    Set oAttrCol = oAttr.CollectionOfAttributes(IID_IUABuiltUpTube)
    If Not oAttrCol Is Nothing Then
        dTubeThickness = oAttrCol.Item("TubeThickness").Value
        If dTubeThickness < 0.000001 Then           ' occurrence attribute may not have been set yet
            Dim oSmartItem As IJSmartItem
            Set oSmartItem = oSmartOcc.ItemObject
            Set oAttr = oSmartItem
            Set oAttrCol = oAttr.CollectionOfAttributes(IID_IUABuiltUpTube)
            If Not oAttrCol Is Nothing Then
                dTubeThickness = oAttrCol.Item("TubeThickness").Value
            End If
            Set oAttr = oSmartOcc           ' reset to the smartoccurrence for remainder of processing
        End If
    End If
    
    Set oAttrCol = oAttr.CollectionOfAttributes(IID_IUABuiltUpCone1)
    If Not oAttrCol Is Nothing Then
        dStartConeThickness = oAttrCol.Item("Cone1Thickness").Value
     Else
        dStartConeThickness = dTubeThickness
    End If

    Set oAttrCol = oAttr.CollectionOfAttributes(IID_IUABuiltUpCone2)
    If Not oAttrCol Is Nothing Then
        dEndConeThickness = oAttrCol.Item("Cone2Thickness").Value
     Else
        dEndConeThickness = dTubeThickness
    End If
    Exit Sub

ErrorHandler:
    HandleError MODULE, METHOD
    Err.Raise E_FAIL
End Sub
                           

Private Function AdjustThicknessForAngle(dConeLength As Double, dDiameterTube As Double, dDiameterCone As Double, dConeThickness As Double) As Double

    Dim diffRadius As Double, dAdjustedThickness As Double
    
    dAdjustedThickness = dConeThickness
    If dConeLength > 0 Then
        diffRadius = 0.5 * Abs(dDiameterTube - dDiameterCone)
        dAdjustedThickness = dConeThickness / Cos(Atn(diffRadius / dConeLength))
    End If
    
    AdjustThicknessForAngle = dAdjustedThickness
    
    Exit Function

End Function

    


Private Sub ICustomSectionShapeService_GetCrossSectionData(ByVal pProfileObject As Object, ByVal distFromStart As Double, ByVal eRepresentationType As SP3DStructInterfaces.structShapeRepresentationEnum, ByVal bBreakIntoComponents As Boolean, sectionPrev As SP3DStructInterfaces.IStructSectionShape, sectionNext As SP3DStructInterfaces.IStructSectionShape)
    ' we're only creating a circular tube type now.  There is no use case for creating a compound shape
    ' or a graphic shape yet and no way for the user to ask for something different.
    
    ' in the future, utility methods will be implemented to convert to a graphic shape and to decompose the shape
    ' if there is a need to do this

   
    ' The parameters that vary along the length of this member is its depth(diameter)
    ' Also thickness of the start code, tube and endcone may be different
    
    Dim tubeMaterial As String
    Dim tubeGrade As String
    Dim tubeThickness As Double, startConeThickness As Double, endConeThickness As Double
    Dim dDiameterStart As Double, dDiameterEnd As Double, dStartConeLength As Double, dEndConeLength As Double
    Dim dTubeDiameter As Double, dLength As Double, dAxisLength As Double
    Dim eCanType As Long
    
    GetComponentMaterial pProfileObject, IID_IUABuiltUpTube, tubeMaterial, tubeGrade, tubeThickness
    
    GetThicknesses pProfileObject, tubeThickness, startConeThickness, endConeThickness

    GetBuiltUpCanExtrusionParameters pProfileObject, dDiameterStart, dDiameterEnd, dStartConeLength, dEndConeLength, dTubeDiameter, dLength, dAxisLength, eCanType
    
    
    
    ' get an IJDPosition at the requested distance along the profile's curve
    Dim pISPSMemberPartCommon As ISPSMemberPartCommon
    Set pISPSMemberPartCommon = pProfileObject
    
    Dim myAxisCurve As IJCurve
    
    Set myAxisCurve = pISPSMemberPartCommon.Axis
    
    Dim pX As Double
    Dim pY As Double
    Dim pZ As Double
    Dim axisLength As Double
    axisLength = myAxisCurve.Length
    
    myAxisCurve.PositionFRatio distFromStart / axisLength, pX, pY, pZ
    
    Dim pointOnCurve As IJDPosition
    Set pointOnCurve = New DPosition
    pointOnCurve.Set pX, pY, pZ
    
    Dim meAsDM As ISPSDesignedMember
    Set meAsDM = pProfileObject
    
    Dim depth As Double
    Dim width As Double
    
    ISPSDesignedMemberHelper_GetNominalSectionSize meAsDM, pointOnCurve, width, depth
    
    If dStartConeLength < distTol Then
        startConeThickness = tubeThickness
    End If
    
    If dEndConeLength < distTol Then
        endConeThickness = tubeThickness
    End If
    
    ' always create a parametric shape.  It may be used to create the other shapes as needed
    Dim parametricShapePrev As IStructSectionRoundHollowShape
    Dim parametricShapeNext As IStructSectionRoundHollowShape
    
    If distFromStart < distTol Then ' at start end
        Set parametricShapeNext = New StructSectionRoundHollowShape
        parametricShapeNext.SetProperties depth, startConeThickness
    ElseIf dStartConeLength - distFromStart > distTol Then ' along start cone
        Set parametricShapeNext = New StructSectionRoundHollowShape
        parametricShapeNext.SetProperties depth, startConeThickness
            
        Set parametricShapePrev = parametricShapeNext
    
    ElseIf (Abs(distFromStart - dStartConeLength)) < distTol And (dStartConeLength > distTol) Then   ' on the startcone tube boundary
        
        Set parametricShapePrev = New StructSectionRoundHollowShape
        Set parametricShapeNext = New StructSectionRoundHollowShape
        
        parametricShapePrev.SetProperties depth, startConeThickness
        parametricShapeNext.SetProperties depth, tubeThickness
  
    ElseIf (axisLength - dEndConeLength - distFromStart) > distTol Then ' along the tube
        Set parametricShapePrev = New StructSectionRoundHollowShape
        parametricShapePrev.SetProperties depth, tubeThickness

        Set parametricShapeNext = parametricShapePrev
    ElseIf (Abs(distFromStart - (axisLength - dEndConeLength)) < distTol) And (dEndConeLength > distTol) Then  'tube endcone boundary
        Set parametricShapePrev = New StructSectionRoundHollowShape
        Set parametricShapeNext = New StructSectionRoundHollowShape
        
        parametricShapePrev.SetProperties depth, tubeThickness
        parametricShapeNext.SetProperties depth, endConeThickness
    
    ElseIf (axisLength - distFromStart) > distTol Then ' along the endcone
        Set parametricShapePrev = New StructSectionRoundHollowShape
        parametricShapePrev.SetProperties depth, endConeThickness
        Set parametricShapeNext = parametricShapePrev
    Else ' at the end
        Set parametricShapePrev = New StructSectionRoundHollowShape
        parametricShapePrev.SetProperties depth, endConeThickness
    End If
    
    
    Dim sectionMaterial As IStructSectionMaterial
    
    If Not parametricShapePrev Is Nothing Then
        Set sectionMaterial = parametricShapePrev
    Else
        Set sectionMaterial = parametricShapeNext
    End If
    
    sectionMaterial.SetMaterial tubeMaterial, tubeGrade
    
    Set sectionPrev = parametricShapePrev
    Set sectionNext = parametricShapeNext
         
    Exit Sub

End Sub

Private Sub ICustomSectionShapeService_GetTransitionLocations(ByVal pProfileObject As Object, distFromStart() As Double)
    
    Dim dDiameterStart As Double
    Dim dDiameterEnd As Double
    Dim dStartConeLength As Double
    Dim dEndConeLength As Double
    Dim dTubeDiameter As Double
    Dim dLength As Double
    Dim dAxisLength As Double
    Dim eCanType As Long
    Dim numLoc As Long
    ReDim distFromStart(0 To 3)
    
    GetBuiltUpCanExtrusionParameters pProfileObject, _
                                     dDiameterStart, _
                                     dDiameterEnd, _
                                     dStartConeLength, _
                                     dEndConeLength, _
                                     dTubeDiameter, _
                                     dLength, _
                                     dAxisLength, _
                                     eCanType
    
    
    
    numLoc = 0
    distFromStart(numLoc) = 0#
    
    If dStartConeLength > distTol Then
        numLoc = numLoc + 1
        distFromStart(numLoc) = dStartConeLength
    End If
    If dEndConeLength > distTol Then
        numLoc = numLoc + 1
        distFromStart(numLoc) = dAxisLength - dEndConeLength
    End If
    
    If (dStartConeLength > distTol) Or (dEndConeLength > distTol) Then
        numLoc = numLoc + 1
        distFromStart(numLoc) = dAxisLength
    End If
    
    ReDim Preserve distFromStart(0 To numLoc)
    
End Sub




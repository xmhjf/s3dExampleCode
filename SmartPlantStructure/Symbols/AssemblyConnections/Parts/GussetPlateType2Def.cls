VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GussetPlateType2Def"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'---------------------------------------------------------------------------
'    Copyright (C) 2003 - 2004 Intergraph Corporation. All rights reserved.
'
'
'
'   Custom assembly definition for the plate part
'   --------------------------------------------
'
'
'History
'    Raj Peringoth      05/07/03      Creation
'---------------------------------------------------------------------------------------

Option Explicit

Private Const MODULE = "GussetPlateType2Def"

' TODO : - Replace <defname> by your selected name
Const m_ItemProgId As String = "SPSPartMacros.GussetPlateType2Def"
Const m_ItemName As String = "SPSPartMacros.GussetPlateType2Def"
Private Const strSourceFile As String = "GussetPlateType2Def.cls"

' Declaration of the User Symbol Services interface
Private Const DOUBLE_VALUE = 8
Private Const BOOL = -7
Private Const CHAR = 1
Dim bOnPreLoad As Boolean
Private m_oLocalizer As IJLocalizer

Implements IJDUserSymbolServices
Implements IJUserAttributeMgmt
Implements IJStructCustomFoulCheck

Public Sub DefinitionInputs(pIH As IJDInputsHelper)
  On Error GoTo ErrorHandler
  
  pIH.SetInput "SupportedPort"
  pIH.SetInput "SupportingPort"
  pIH.SetInput "CutbackPlane"
  
  Exit Sub
ErrorHandler:
  pIH.ReportError
End Sub


Public Sub IJDUserSymbolServices_InitializeSymbolDefinition(ByRef pDefinition As IJDSymbolDefinition)
  Const MT = "IJDUserSymbolServices_InitializeSymbolDefinition"
  On Error GoTo ErrorHandler


  pDefinition.SupportOnlyOption = igSYMBOL_NOT_SUPPORT_ONLY
  pDefinition.MetaDataOption = igSYMBOL_DYNAMIC_METADATA
    
  ' Define the inputs -
  Dim pIH As IJDInputsHelper
  Set pIH = New InputHelper
  pIH.Definition = pDefinition
'  pIH.InitAs m_FamilyProgid
  DefinitionInputs pIH
  
  ' Aggregator Type
  Dim pAD As IJDAggregatorDescription
  Set pAD = pDefinition
  pAD.AggregatorClsid = "{A46498E6-9116-42B1-8A18-031415C07428}"     'CStructCustomPlatePart
  pAD.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputAggregator"
  pAD.SetCMRemoveInputs imsCOOKIE_ID_USS_LIB, "CMRemoveInputAggregator"
  pAD.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructAggregator"
  Set pAD = Nothing
  
  ' Aggregator property
  Dim pAPDs As IJDPropertyDescriptions
  Set pAPDs = pDefinition
  pAPDs.RemoveAll ' Remove all the previous property descriptions
  'Watches IJStructPlate to get notified about thickness changes. The plates position is affected by
  'the thickness
  pAPDs.AddProperty "IJStructPlate", 1, "{274EE192-A0A5-44E7-B536-8D44A08FA64F}", "CMEvaluateThickness", imsCOOKIE_ID_USS_LIB
  
  'computes occ attributes
  pAPDs.AddProperty "IJDAttributes", 2, "{B25FD387-CFEB-11D1-850B-080036DE8E03}", "CMEvaluateOccAttributes", imsCOOKIE_ID_USS_LIB
  'computes transformation matrix
  'computes IJDOccurrence
  pAPDs.AddProperty "PlatePosition", 3, "{274317DB-0F9D-11D2-94AD-080036CD8E03}", "CMComputePlatePosition", imsCOOKIE_ID_USS_LIB

  
  
  Set pAPDs = Nothing
   
  ' Define the members
  Dim pMemberDescriptions As IJDMemberDescriptions
  Dim pMemberDescription As IJDMemberDescription
  Dim pPropertyDescriptions As IJDPropertyDescriptions
  
  Set pMemberDescriptions = pDefinition
  ' Remove all the previous member descriptions
  pMemberDescriptions.RemoveAll
  Set pMemberDescriptions = Nothing
  'No assembly members for this parent

  Exit Sub
ErrorHandler:  HandleError MODULE, MT
End Sub
'Aggregator custom methods.......................................................
Public Sub CMFinalConstructAggregator(pAggregatorDescription As IJDAggregatorDescription)
Const METHOD = "CMFinalConstructAggregator"
On Error GoTo ErrorHandler
  'Get the smart plate object
  Dim oCustomPlate As IJStructCustomPlatePart
  Set oCustomPlate = pAggregatorDescription.CAO
  
  'Set the generation pattern of the smart plate geometry
    SetCustomPlatePartGenerationAE oCustomPlate
    
Exit Sub
ErrorHandler:      HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputAggregator(pAggregatorDescription As IJDAggregatorDescription)
Const METHOD = "CMSetInputAggregator"
On Error GoTo ErrorHandler
    Dim pIJSymbolOfCA0 As IMSSymbolEntities.IJDSymbol
 
    Dim i As Long, lcnt As Long
    Dim oRCSymbol As Object
    Set pIJSymbolOfCA0 = pAggregatorDescription.CAO
    Set oRCSymbol = pIJSymbolOfCA0.IJDReferencesArg.GetReferencesCollection

    Dim pIRCAsm As IJDReferencesCollection
    Dim oRefColl1 As IMSSymbolEntities.IJDReferencesCollection
    Set pIRCAsm = GetRefCollFromSmartOccurrence(pAggregatorDescription.CAO)

    If pIRCAsm Is Nothing Then
    Exit Sub
    End If

    If Not oRCSymbol Is Nothing Then
        pIJSymbolOfCA0.IJDReferencesArg.SetReferencesCollection pIRCAsm
    Exit Sub 'already have one get out unless we want to change it's contents
    Else 'create a new or initial one

    ' Create the new symbol input reference collection and copy the inputs reference

        pIJSymbolOfCA0.IJDReferencesArg.SetReferencesCollection pIRCAsm

    End If

  
Exit Sub
ErrorHandler:     HandleError MODULE, METHOD
End Sub
Public Sub CMRemoveInputAggregator(pAggregatorDescription As IJDAggregatorDescription)
Const METHOD = "CMRemoveInputAggregator"
On Error GoTo ErrorHandler

  
Exit Sub
ErrorHandler:     HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateOccAttributes(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
 Const MT = "CMEvaluateOccAttributes"
 
    
Exit Sub
End Sub
Public Sub CMEvaluateThickness(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
 Const MT = "CMEvaluateThickness"
 
    
Exit Sub
End Sub

Public Sub CMComputePlatePosition(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
 Const MT = "CMComputePlatePosition"
 On Error GoTo ErrorHandler
    Dim oRefCollCAO As IJDReferencesCollection
    Dim oSuppedPort As ISPSSplitAxisPort, oSuppingPort As ISPSSplitAxisPort
    Dim oSuppedPart As ISPSMemberPartPrismatic, oSuppingPart As ISPSMemberPartPrismatic
    Dim oXVecSupped As New DVector, oXVecSupping As New DVector
    Dim oYVecSupped As New DVector, oZVecSupped As New DVector, oYVecSupping As New DVector, oZVecSupping As New DVector
    Dim oPlateXvec As IJDVector, oPlateYvec As IJDVector, oPlateZvec As IJDVector, oSuppingPlaneNorm As IJDVector
    Dim oPlateThicknessDir As New DVector
    Dim oSuppingPlaneRootPos As IJDPosition
    Dim oCPLine As IJLine, oLine As IJLine
    Dim iQuadrant As Integer
    Dim numOverlaps As Long
    Dim numIntersection As Long
    Dim intersectionPoint As IJPoint
    Dim code As Geom3dIntersectConstants
    Dim oPlateCornerPos As New DPosition
    Dim oOcc As IJDOccurrence
    Dim oMat As New DT4x4
    Dim plateDepth#, plateWidth#, overlapLength#, cutLength#, cutWidth#
    Dim pIJAttribs   As IJDAttributes
    Dim oPlatePos1 As New DPosition, oPlatePos2 As New DPosition
    Dim oSurf As IJSurface
    Dim colIntrsectnElms As IJElements
    Dim x#, y#, z#, nX#, nY#, nZ#, x1#, y1#, z1#
    Dim oGeomFactory As New GeometryFactory
    Dim oSuppedDir As New DVector
    Dim iPort As SPSMemberAxisPortIndex
    Dim colCutbacks As IJElements
    Dim oPlane As IJPlane
    Dim bYSuppedYSuppingParallel As Boolean, bYSuppedZSuppingParallel As Boolean
    Dim bZSuppedYSuppingParallel As Boolean, bZSuppedZSuppingParallel As Boolean
    Dim oMembTrnsfm As IJDT4x4
    Dim oPos As IJDPosition, oPos1 As IJDPosition
    Dim strError As String
    Dim dblOffsetDist1#, dblOffsetDist2#
    Dim oStructPlate As IJStructPlate
    Dim Thickness As Double
    Dim oPlateCornerPos1 As New DPosition
    Dim oMatGtoL As New DT4x4
    Dim strCSType As String
    Dim bCenterPlate As Boolean ' controls how the plate is positioned length wise
    Dim bWarningError As Boolean
    
    Set oStructPlate = pObject
    Set pIJAttribs = pObject
    plateDepth = pIJAttribs.CollectionOfAttributes("IJUASPSPlatePartDim").Item("Length").Value
    plateWidth = pIJAttribs.CollectionOfAttributes("IJUASPSPlatePartDim").Item("Width").Value
    overlapLength = pIJAttribs.CollectionOfAttributes("IJUASPSGussetPlatePartType2").Item("OverlapLength").Value
    cutLength = pIJAttribs.CollectionOfAttributes("IJUASPSPlatePartCutDim").Item("cutLength").Value
    cutWidth = pIJAttribs.CollectionOfAttributes("IJUASPSPlatePartCutDim").Item("cutWidth").Value
    Thickness = oStructPlate.Thickness
    
    Set oOcc = pObject
    Set oRefCollCAO = GetRefCollFromSmartOccurrence(pObject)
    'get input ports

    Set oSuppedPort = oRefCollCAO.IJDEditJDArgument.GetEntityByIndex(1)
    Set oSuppingPort = oRefCollCAO.IJDEditJDArgument.GetEntityByIndex(2)
    Set oSuppedPart = oSuppedPort.Part
    Set oSuppingPart = oSuppingPort.Part
    iPort = oSuppedPort.portIndex

    'get X, Y and Z Vectors of supported
    Set oXVecSupped = GetXVector(oSuppedPart)
    Set oYVecSupped = GetYVector(oSuppedPart)
    Set oZVecSupped = GetZVector(oSuppedPart)
    'get X,Y and Z Vectors of supporting
    Set oXVecSupping = GetXVector(oSuppingPart)
    Set oYVecSupping = GetYVector(oSuppingPart)
    Set oZVecSupping = GetZVector(oSuppingPart)
    
    If Abs(oXVecSupped.Dot(oXVecSupping)) < distTol Then ' supped and supping are at 90 degrees
        'if angle between supped and supping is 90 degree, center the plate
        'other wise it is positioned based on the over lap length
        bCenterPlate = True
    End If
    
    If iPort = SPSMemberAxisStart Then
        oSuppedPart.Axis.EndPoints x, y, z, x1, y1, z1
    Else
        oSuppedPart.Axis.EndPoints x1, y1, z1, x, y, z
    End If
    'oSuppedDir is used to decide which way the plate the plate is positioned along the X direction of the plate
    oSuppedDir.Set x1 - x, y1 - y, z1 - z
    oSuppedDir.Length = 1
    bYSuppedYSuppingParallel = AreVectorsParallel(oYVecSupped, oYVecSupping)
    bYSuppedZSuppingParallel = AreVectorsParallel(oYVecSupped, oZVecSupping)
    bZSuppedYSuppingParallel = AreVectorsParallel(oZVecSupped, oYVecSupping)
    bZSuppedZSuppingParallel = AreVectorsParallel(oZVecSupped, oZVecSupping)
    strCSType = oSuppedPart.CrossSection.Definition.Type
    If bYSuppedYSuppingParallel Or bYSuppedZSuppingParallel Then
        ' weak axis of supported and weak or strong axis supporting are parallel
        
        If strCSType = "WT" Or strCSType = "MT" Or strCSType = "ST" Then
            'Th strong axis of the suppported is in the plane containing
            'supported and supporting axes.  We need to create a warning todo list item
            'as the plate cannot be positioned on the supported member for this orientation
            bWarningError = True
            'don't raise an error before the plates's position is computed as this would cause
            'plate to be placed at global origin. Instead we will create a todo list item at the end
            
'            SPSToDoErrorNotify PlatePartToDoMsgCodelist, TDL_PARTMACROS_PLATE_CANNOT_POSITIONED, oStructPlate, Nothing
'            Err.Raise SPS_MACRO_WARNING
            Set oCPLine = GetCPLine(oSuppedPart, 5)
            dblOffsetDist1 = Thickness / 2 'take care of the situation where the plate's middle plane on the CP line
            dblOffsetDist2 = 0 'take care of the situation where the plate's Z and thickness dir are in opposite directions
        
        ElseIf strCSType = "2L" Then
            'for 2L ,place the plate between the flanges
            Set oCPLine = GetCPLine(oSuppedPart, 5)
            dblOffsetDist1 = Thickness / 2 'take care of the situation where the plate's middle plane on the CP line
            dblOffsetDist2 = 0 'take care of the situation where the plate's Z and thickness dir are in opposite directions
        Else
        'get  CP4 line for all other shapes. This gives  correct location for "L", "C"
            Set oCPLine = GetCPLine(oSuppedPart, 4)
            dblOffsetDist1 = 0 'take care of the situation where the plate's middle plane on the CP line
            dblOffsetDist2 = Thickness 'take care of the situation where the plate's Z and thickness dir are in opposite directions
        End If
        'get the normal vector for the plate. Plate's thickness is in this direction
        Set oPlateThicknessDir = oYVecSupped
        If oSuppedPart.Rotation.Mirror = True Then
            'mirrored so  reverse the direction of the normal.
            oPlateThicknessDir.x = -oPlateThicknessDir.x
            oPlateThicknessDir.y = -oPlateThicknessDir.y
            oPlateThicknessDir.z = -oPlateThicknessDir.z
        End If
    ElseIf bZSuppedYSuppingParallel Or bZSuppedZSuppingParallel Then
        ' strong axis of supported and weak or strong axis supporting are parallel
        'get the normal vector for the plate. Plate's thickness is in this direction
        Set oPlateThicknessDir = oZVecSupped
        dblOffsetDist1 = 0 'take care of the situation where the plate's middle plane on the CP line
        dblOffsetDist2 = Thickness 'take care of the situation where the plate's Z and thickness dir are in opposite directions
        Select Case oSuppedPart.CrossSection.Definition.Type
        Case "WT", "MT", "ST"
            'for "T"
            'get CP8 line
            Set oCPLine = GetCPLine(oSuppedPart, 8)
        Case Else
            'for "L", "C" and other shapes
            'get CP2 line
            Set oCPLine = GetCPLine(oSuppedPart, 2)
            'reverse the direction of the normal so that it is away from the face of supported. Plate's thickness is in this direction
            oPlateThicknessDir.x = -oPlateThicknessDir.x
            oPlateThicknessDir.y = -oPlateThicknessDir.y
            oPlateThicknessDir.z = -oPlateThicknessDir.z
        End Select
    End If

    'get supporting face where the plate will be placed
    iQuadrant = GetIncidentMemberQuadrant(oSuppingPart, oSuppedPart, iPort)
    GetMembSidePlane oSuppingPart, iQuadrant, oSuppingPlaneRootPos, oSuppingPlaneNorm
    
    'make the line infinite
    If Not oCPLine Is Nothing Then
        If TypeOf oCPLine Is IJLine Then
            Set oLine = oCPLine
            oLine.Infinite = True
        End If
    End If
    
    'get cutback planes for this end of the supported member
    Set colCutbacks = oSuppedPart.Cutbacks(iPort)
    If colCutbacks.Count > 0 Then
        'get the first cutback plane. The end is not expected to have more than one cutback plane for this connection type
        Set oSurf = colCutbacks.Item(1) 'get the cutback plane
        Set oPlane = oSurf
        'get intersection of projected CP lines.
        oSurf.Intersect oCPLine, colIntrsectnElms, code
        If colIntrsectnElms.Count > 0 Then
            Set intersectionPoint = colIntrsectnElms.Item(1)
            intersectionPoint.GetPoint x, y, z
            oSuppedDir.Get nX, nY, nZ
            x = x + overlapLength * nX
            y = y + overlapLength * nY
            z = z + overlapLength * nZ
            oPlatePos1.Set x, y, z
            Set oPlatePos2 = ProjectPosToPlane(oPlatePos1, oSuppingPlaneRootPos, oSuppingPlaneNorm)
            plateWidth = oPlatePos2.DistPt(oPlatePos1)
            pIJAttribs.CollectionOfAttributes("IJUASPSPlatePartDim").Item("Width").Value = plateWidth
            Set oPlateCornerPos = oPlatePos2

            
            'get the plate X and Y vector from Z vector and oSuppingPlaneNorm
            Set oPlateYvec = oSuppingPlaneNorm
            Set oPlateXvec = oPlateYvec.Cross(oPlateThicknessDir)
            oPlateXvec.Length = 1
            
            If oPlateXvec.Dot(oSuppedDir) > 0 Then
                ' we need to change direction of x so that the plate get lengthened in the right direction
                oPlateXvec.x = -oPlateXvec.x
                oPlateXvec.y = -oPlateXvec.y
                oPlateXvec.z = -oPlateXvec.z
            End If
            Set oPlateZvec = oPlateXvec.Cross(oPlateYvec)
            oPlateZvec.Length = 1
         
            'set up the transformation matrix
            oMat.LoadIdentity
            'X-axis of plate
            oMat.IndexValue(0) = oPlateXvec.x
            oMat.IndexValue(1) = oPlateXvec.y
            oMat.IndexValue(2) = oPlateXvec.z
            'Y-axis of plate
            oMat.IndexValue(4) = oPlateYvec.x
            oMat.IndexValue(5) = oPlateYvec.y
            oMat.IndexValue(6) = oPlateYvec.z
            'Z-axis of plate
            oMat.IndexValue(8) = oPlateZvec.x
            oMat.IndexValue(9) = oPlateZvec.y
            oMat.IndexValue(10) = oPlateZvec.z
            'take care of the situation where the plate's middle plane on the CP line
            'case of 2L with the plate between webs
            
            oMatGtoL.LoadMatrix oMat 'create a global to local matrix
            oMatGtoL.Invert
            
            oPlatePos1.Set 0, 0, -dblOffsetDist1
            If oPlateZvec.Dot(oPlateThicknessDir) < 0 Then
                oPlatePos1.z = oPlatePos1.z - dblOffsetDist2
            End If
            If bCenterPlate = True Then
                oPlatePos1.x = -plateDepth / 2
            End If
            
            Set oPlatePos1 = oMat.TransformPosition(oPlatePos1)
            oPlateCornerPos.x = oPlateCornerPos.x + oPlatePos1.x
            oPlateCornerPos.y = oPlateCornerPos.y + oPlatePos1.y
            oPlateCornerPos.z = oPlateCornerPos.z + oPlatePos1.z
            

            'oPlateCornerPos1 is equal to oPlateCornerPos offset by the thickness of the plate
            Set oPlateCornerPos1 = oMatGtoL.TransformPosition(oPlateCornerPos) 'to local
            oPlateCornerPos1.z = oPlateCornerPos1.z + Thickness 'add thickness
            Set oPlateCornerPos1 = oMat.TransformPosition(oPlateCornerPos1) ' to global
           

            
            'check if the plate is placed within the depth or width of the supporting member
            'the plate is within the depth or width if oPlateCornerPos1 and oPlateCornerPos are within the depth
            'to check this we transform oPlateCornerPos1 and oPlateCornerPos to member coordinates
            oSuppingPart.Rotation.GetTransform oMembTrnsfm
            oMembTrnsfm.Invert 'global to local
            'Origin of plate
            oMat.IndexValue(12) = oPlateCornerPos.x
            oMat.IndexValue(13) = oPlateCornerPos.y
            oMat.IndexValue(14) = oPlateCornerPos.z
            'set the matrix on the Baseplate occurrence
            oOcc.Matrix = oMat
            If bWarningError Then
                Err.Raise SPS_MACRO_WARNING
            End If
            
            Set oPos = oMembTrnsfm.TransformPosition(oPlateCornerPos)
            Set oPos1 = oMembTrnsfm.TransformPosition(oPlateCornerPos1)
            If bYSuppedYSuppingParallel Or bZSuppedYSuppingParallel Then
                Set oCPLine = GetCPLine(oSuppingPart, 1)
                oCPLine.Transform oMembTrnsfm
                oCPLine.GetStartPoint x, y, z
                Set oCPLine = GetCPLine(oSuppingPart, 3)
                oCPLine.Transform oMembTrnsfm 'transform to member coordinates
                oCPLine.GetStartPoint x1, y1, z1
                If y > y1 Then
                    If oPos.y > y Or oPos.y < y1 Or oPos1.y > y Or oPos1.y < y1 Then
                        Err.Raise SPS_MACRO_WARNING
                        'plate is not within the width of the supporting member
                    End If
                Else
                    If oPos.y > y1 Or oPos.y < y Or oPos1.y > y1 Or oPos1.y < y Then
                        Err.Raise SPS_MACRO_WARNING   'plate is not within the width of the supporting member
                    End If
                End If
            ElseIf bYSuppedZSuppingParallel Or bZSuppedZSuppingParallel Then
                Set oCPLine = GetCPLine(oSuppingPart, 1)
                oCPLine.Transform oMembTrnsfm
                oCPLine.GetStartPoint x, y, z
                Set oCPLine = GetCPLine(oSuppingPart, 7)
                oCPLine.Transform oMembTrnsfm 'transform to member coordinates
                oCPLine.GetStartPoint x1, y1, z1
                If z > z1 Then
                    If oPos.z > z Or oPos.z < z1 Or oPos1.z > z Or oPos1.z < z1 Then
                        Err.Raise SPS_MACRO_WARNING 'plate is not within the depth of the supporting member
                    End If
                Else
                    If oPos.z > z1 Or oPos.z < z Or oPos1.z > z1 Or oPos1.z < z Then
                        Err.Raise SPS_MACRO_WARNING 'plate is not within the depth of the supporting member
                    End If
                End If
            End If
        End If
    End If
Exit Sub
ErrorHandler:
    ' For errors logged with E_FAIL, a todo list error will be generated so we should not
    '   be logging anything to the error log
    If Err.Number = SPS_MACRO_WARNING Then
        SPSToDoErrorNotify PlatePartToDoMsgCodelist, TDL_PARTMACROS_PLATE_CANNOT_POSITIONED, oStructPlate, Nothing
        Err.Raise SPS_MACRO_WARNING
    Else
        Err.Raise ReportError(Err, strSourceFile, MT).Number
    End If
End Sub


'
' The following methods are generic for all the Custom assembly
'
'
Public Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, ByVal defParams As Variant, ByVal ActiveConnection As Object) As Object
  ' This method is in charge of the creation of the symbol definition object
  ' You can keep the current design unchanged for basic VB symbol definition.
   Const MT = "IJDUserSymbolServices_InstanciateDefinition"
  On Error GoTo ErrorHandler
 
  
  Dim pDefinition As IJDSymbolDefinition
  Dim pFact As IJCAFactory
  Set pFact = New CAFactory
  Set pDefinition = pFact.CreateCAD(ActiveConnection)
  
  ' Set definition progId and codebase
  pDefinition.ProgId = m_ItemProgId
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
ErrorHandler:  HandleError MODULE, MT
End Function
Public Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
  ' Name should be unique
  IJDUserSymbolServices_GetDefinitionName = m_ItemName
End Function
Public Sub IJDUserSymbolServices_InvokeRepresentation(ByVal sblOcc As Object, ByVal repName As String, ByVal outputcoll As Object, ByRef arrayOfInputs())
 ' Obsolete method.
End Sub

Public Function IJDUserSymbolServices_EditOccurence(ByRef pSymbolOccurence As Object, ByVal transactionMgr As Object) As Boolean
 ' Obsolete method. Instead you can record your custom command within the definition (see IJDCommandDescription interface)
 IJDUserSymbolServices_EditOccurence = False
End Function




Private Sub IJStructCustomFoulCheck_GetConnectedParts(ByVal pPartObject As Object, ByVal pIJMonUnks As SP3DStructGeneric.IJElements)
    Call GetRelatedParts(pPartObject, pIJMonUnks)
End Sub

Private Sub IJStructCustomFoulCheck_GetFoulInterfaceType(pFoulInterfaceType As SP3DStructGeneric.FoulInterfaceType)
    pFoulInterfaceType = StandardGraphicEntity
End Sub

Private Function IJUserAttributeMgmt_OnAttributeChange(ByVal pIJDAttrs As SPSMembers.IJDAttributes, ByVal CollAllDisplayedValues As Object, ByVal pAttrToChange As SPSMembers.IJAttributeDescriptor, ByVal varNewAttrValue As Variant) As String
Const METHOD = "IJUserAttributeMgmt_OnAttributeChange"
On Error GoTo ErrorHandler
    IJUserAttributeMgmt_OnAttributeChange = m_oLocalizer.GetString(IDS_PARTMACROS_ERROR, "ERROR")
    
    ' Validate the attribute new value first before any further processing
    Dim ErrStr As String
    Dim i As Integer
    Dim pColl As Collection
    Dim pAttrDescr As IJAttributeDescriptor
    Dim pCutAttrDescr As IJAttributeDescriptor
    Dim NonStateRO As Long
    
    If bOnPreLoad = False Then
        ErrStr = UserAttributeMgmt_Validate(pIJDAttrs, pAttrToChange.InterfaceName, pAttrToChange.attrName, varNewAttrValue)
        If Len(ErrStr) > 0 Then
'            IJUserAttributeMgmt_OnAttributeChange = "ERROR::Bad Value"
            IJUserAttributeMgmt_OnAttributeChange = ErrStr
            Exit Function
        End If
    End If
    
    'the cut dimensions (Length /Width ) need to be all zero or all > 0
    Set pColl = CollAllDisplayedValues
    
    If pAttrToChange.attrName = "CutLength" Then
        For i = 1 To pColl.Count
            Set pAttrDescr = pColl.Item(i)
            If pAttrDescr.attrName = "CutWidth" Then
                Set pCutAttrDescr = pAttrDescr
                Exit For
            End If
        Next
    ElseIf pAttrToChange.attrName = "CutWidth" Then
        For i = 1 To pColl.Count
            Set pAttrDescr = pColl.Item(i)
            If pAttrDescr.attrName = "CutLength" Then
                Set pCutAttrDescr = pAttrDescr
                Exit For
            End If
        Next
    End If
    If Not pCutAttrDescr Is Nothing Then
        If varNewAttrValue = 0 Then
            pCutAttrDescr.AttrValue = 0
        ElseIf varNewAttrValue > 0 Then
            If pCutAttrDescr.AttrValue = 0 Then
                pCutAttrDescr.AttrValue = varNewAttrValue
            End If
        End If
    End If
    IJUserAttributeMgmt_OnAttributeChange = ""
   
Exit Function
ErrorHandler:  HandleError MODULE, METHOD
End Function

Private Function IJUserAttributeMgmt_OnPreCommit(ByVal pIJDAttrs As SPSMembers.IJDAttributes, ByVal CollAllDisplayedValues As Object) As String

End Function

Private Function IJUserAttributeMgmt_OnPreLoad(ByVal pIJDAttrs As SPSMembers.IJDAttributes, ByVal CollAllDisplayedValues As Object) As String
Const METHOD = "IJUserAttributeMgmt_OnPreLoad"
On Error GoTo ErrorHandler
    IJUserAttributeMgmt_OnPreLoad = m_oLocalizer.GetString(IDS_PARTMACROS_ERROR, "ERROR")
    bOnPreLoad = True ' optimization to avoid value validation in OnAttrChange
    
    Dim i As Integer
    Dim pAttrColl As Collection
    Dim pAttrDescr As IJAttributeDescriptor
    Dim attrName As String
    Dim ErrStr As String

    Set pAttrColl = CollAllDisplayedValues
    
    For i = 1 To pAttrColl.Count
        Set pAttrDescr = pAttrColl.Item(i)
        If pAttrDescr.attrName = "Width" Then
            pAttrDescr.AttrState = pAttrDescr.AttrState Or AttributeDescriptor_ReadOnly
        End If
    Next
    
    For i = 1 To pAttrColl.Count
        Set pAttrDescr = pAttrColl.Item(i)
            ErrStr = IJUserAttributeMgmt_OnAttributeChange(pIJDAttrs, CollAllDisplayedValues, pAttrDescr, pAttrDescr.AttrValue)
            If Len(ErrStr) > 0 Then
                bOnPreLoad = False
                Exit Function
            End If
    Next
    
    bOnPreLoad = False

    IJUserAttributeMgmt_OnPreLoad = ""
Exit Function
ErrorHandler:  HandleError MODULE, METHOD
End Function
Private Function UserAttributeMgmt_Validate(ByVal pIJDAttrs As SPSMembers.IJDAttributes, sInterfaceName As String, sAttributeName As String, ByVal varAttributeValue As Variant) As String
Const METHOD = "UserAttributeMgmt_Validate"
On Error GoTo ErrorHandler
    UserAttributeMgmt_Validate = m_oLocalizer.GetString(IDS_PARTMACROS_ERROR, "ERROR")

    Dim dInputs As IJDInputs
    Dim CurrentInput As IJDInput
    Dim oAttribute As IJDAttribute
    Dim PC As DParameterContent
    Dim bvalid As Boolean
    Dim oSymbolOcc As IJDSymbol
    Set oSymbolOcc = pIJDAttrs
    Dim oSymbolDef As IJDSymbolDefinition
    Dim ErrMessage As String
    Set oSymbolDef = oSymbolOcc.IJDSymbolDefinition(2)
    Set dInputs = oSymbolDef.IJDInputs
    Set PC = New DParameterContent
    
    Set oAttribute = pIJDAttrs.CollectionOfAttributes(sInterfaceName).Item(sAttributeName)
    If oAttribute.Value <> "" Then
        If oAttribute.AttributeInfo.Type = igString Then    ' check for string type here
        Else
            PC.UomValue = oAttribute.Value
            Set CurrentInput = Nothing
            bvalid = True
            On Error Resume Next
            Set CurrentInput = dInputs.GetInputByName(oAttribute.AttributeInfo.Name)
            If Not CurrentInput Is Nothing Then
                CurrentInput.IJDInputDuringGame.Definition = oSymbolDef
                CurrentInput.IJDInputStdCustomMethod.InvokeCMCheck PC, bvalid, ErrMessage
                CurrentInput.IJDInputDuringGame.Definition = Nothing
                Set oSymbolOcc = Nothing
                Set oSymbolDef = Nothing
                If bvalid = False Then
'                    UserAttributeMgmt_Validate = "Symbol CMCheck Failed"
                    UserAttributeMgmt_Validate = ErrMessage
                    Exit Function
                Else
                End If
            End If
            On Error GoTo ErrorHandler
        End If
    End If
' get the list of interfaces implemented by the schema from IJDAttributes
' make sure that you are not looking into a system interface
' from the input interfaceName and propertyName, get the property type from catalog info
' select case on the property types, and in there, mention the valid attribute values for each propertyName
    Dim InterfaceID As Variant
'    Dim oAttrObj As IJDAttribute
    Dim oAttrObj As IJDAttributeInfo
    Dim oInterfaceInfo As IJDInterfaceInfo
    Dim oAttributeMetaData As IJDAttributeMetaData
'    Dim oAttrCol As IJDAttributesCol
    Dim oAttrCol As IJDInfosCol
    Dim IsInterfaceFound As Boolean
    Dim AttrCount As Long
    Dim AttrType As Long
    
    Set oAttributeMetaData = pIJDAttrs
    IsInterfaceFound = False
    For Each InterfaceID In pIJDAttrs
        Set oInterfaceInfo = Nothing
        Set oInterfaceInfo = oAttributeMetaData.InterfaceInfo(InterfaceID)
        If (oInterfaceInfo.IsHardCoded = False) Then
            If (oInterfaceInfo.Name = sInterfaceName) Then
                IsInterfaceFound = True
                Exit For
            End If
        End If
    Next
    
'    Set oAttributeMetaData = Nothing
    Set oInterfaceInfo = Nothing
    
    If IsInterfaceFound = False Then
        UserAttributeMgmt_Validate = m_oLocalizer.GetString(IDS_PARTMACROS_SCHEMAERROR, "SchemaERROR")
        GoTo ErrorHandler
    End If
'    Set oAttrCol = pIJDAttrs.CollectionOfAttributes(InterfaceID)
    Set oAttrCol = oAttributeMetaData.InterfaceAttributes(InterfaceID)
    ' loop on the attributes on the interface to match the supplied attribute type
    For AttrCount = 1 To oAttrCol.Count
        Set oAttrObj = oAttrCol.Item(AttrCount)
        If oAttrObj.Name = sAttributeName Then
            
            Select Case oAttrObj.Type
                Case DOUBLE_VALUE
                    If (sAttributeName = "OverlapLength") Or (sAttributeName = "Length") Then 'overlap length and length should be >0
                        If (varAttributeValue <= 0#) Then
                            UserAttributeMgmt_Validate = sAttributeName
                            Set oAttributeMetaData = Nothing
                            Exit Function
                        End If
                    ElseIf (sAttributeName = "CutLength") Or (sAttributeName = "CutWidth") Then 'cut length/width  should not be < 0
                        If (varAttributeValue < 0#) Then
'                            UserAttributeMgmt_Validate = sAttributeName
                            UserAttributeMgmt_Validate = m_oLocalizer.GetString(IDS_PARTMACROS_VALUE_NEGATIVE, "Negative Attribute Value")
                            Set oAttributeMetaData = Nothing
                            Exit Function
                        End If
                    End If
            
            End Select

        End If
    Next
    
    UserAttributeMgmt_Validate = ""
    Set oAttributeMetaData = Nothing

Exit Function
ErrorHandler:  HandleError MODULE, METHOD
End Function


Private Sub Class_Initialize()
Set m_oLocalizer = New IMSLocalizer.Localizer
m_oLocalizer.Initialize App.Path & "\" & App.EXEName
End Sub

Private Sub Class_Terminate()
Set m_oLocalizer = Nothing
End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PadPlateRectDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'---------------------------------------------------------------------------
'    Copyright (C) 2006 Intergraph Corporation. All rights reserved.
'
'
'
'   Pad definition:
'   ---------
'   defines the ractangular pad
'
'
'History
' 03/26/03  AS    Creation
' 07/26/06  AS    TR#102438 Gen Surface AC does not work correctly against IS Plates. Changed plate positioning logic.
' 10/09/06  AS    TR#102696  Surface Trim feature can not be placed correctly against IS Plates
'---------------------------------------------------------------------------------------


Option Explicit

Const m_ProgId As String = "SPSPartMacros.PadPlateRectDef"
Private Const MODULE = "PadPlateRectDef"
Private Const m_strSourceFile As String = "PadPlateRectDef.cls"

' Error management
Private sMethod As String
Private sERROR As String

'Attr Management
Private bOnPreLoad As Boolean
Private Const DOUBLE_VALUE = 8

' Declaration of the User Symbol Services and Attr Management interfaces
Implements IJDUserSymbolServices
Implements IJUserAttributeMgmt


Public Sub IJDUserSymbolServices_InitializeSymbolDefinition(ByRef pDefinition As IJDSymbolDefinition)
  On Error GoTo ErrorHandler
  sMethod = "IJDUserSymbolServices_InitializeSymbolDefinition"
  sERROR = ""
  
  'MsgBox sMethod
    
  pDefinition.SupportOnlyOption = igSYMBOL_NOT_SUPPORT_ONLY
  pDefinition.MetaDataOption = igSYMBOL_DYNAMIC_METADATA

  ' Get user method collection
  Dim pUMs As IJDUserMethods
  Set pUMs = pDefinition
  Dim mthCookie As Long
    
  ' Define the inputs -
  SetInputs4Class_Pad pDefinition
  
  ' Aggregator
  Dim pAD As IJDAggregatorDescription
  Set pAD = pDefinition
  'pAD.AggregatorClsid = CONST_CustomPlatePart     'CStructCustomPlatePart 'Commented as it is not being used.
  pAD.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputAggregator"
  pAD.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructAggregator"
  pAD.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CMMigrateAggregator"
  Set pAD = Nothing
  
  ' Aggregator property
  Dim pAPDs As IJDPropertyDescriptions
  Set pAPDs = pDefinition
  pAPDs.RemoveAll ' Remove all the previous property descriptions
  'listens to IJUASPSPlatePartDim (changes in plate size) and also outputs it for auto sizing rule
  pAPDs.AddProperty "PlateSize", 1, "IJUASPSPlatePartDim"
  'listens to IJUASPSPlateThickness, corner radius and sizing rule
  pAPDs.AddProperty "PlateThickness", 2, "IJUASPSPlateThickness"
  pAPDs.AddProperty "CornerRadius", 3, "IJUACornerRadius"
  pAPDs.AddProperty "PlateSizeParams", 4, "IJUAPlateSizeParams"
  
  'computes transformation matrix through IJDOccurrence
  pAPDs.AddProperty "PlatePosition", 5, CONST_IJDOccurrence, "CMComputePlatePosition", imsCOOKIE_ID_USS_LIB ', igPROCESS_PD_BEFORE_SYMBOL_UPDATE
  
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
ErrorHandler:  HandleError MODULE, sMethod

  
End Sub


Public Sub CMFinalConstructAggregator(pAggregatorDescription As IJDAggregatorDescription)
Const METHOD = "CMFinalConstructAggregator"
On Error GoTo ErrorHandler
  'Get the smart plate object
    Dim oCustomPlate As IJStructCustomPlatePart
    Set oCustomPlate = pAggregatorDescription.CAO
  
    'Set the generation pattern of the smart plate geometry
    SetCustomPlatePartGenerationAE oCustomPlate

'MsgBox MODULE + METHOD
Exit Sub
ErrorHandler:      HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputAggregator(pAggregatorDescription As IJDAggregatorDescription)
Const METHOD = "CMSetInputAggregator"
On Error GoTo ErrorHandler
    'TODO: Have a second look at this method. May not be needed
    'MsgBox MODULE + METHOD
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

' Compute the position of the plate. The logic is following:
' Get the size of plate, location of member's CP 10 and its axis.
'   Get member's CP10 line
'   Intersect CP10 line with Dynmic surface
'   Get the intersection point, P0.
'   Get the surface normal at P0,i.e., N0
'   Get all corner points C1,.. Cn of the pad (if pad is circular, divide the circumference in to 8 vertices)
'   Pass infinite lines through these points parallel to member axis
'   Get intersection between surface and these lines,P1,……Pn
'   Get subset of Pn, where the points are in +ve direction (front) of the surface, Ppn
'   Get largest distance between the Ppn and Cn, dm. This is dist by which the pad is to be moved
'   Get Members xform matrix and convert the point P0 to LCS of the member, P0l
'   Increment or decrement (Start or End of member) ther x-component of P0l by dm.
'   Convert P0l to global coord sys, P0g
'   So P0g is the root point of plates's bottom plane and N0 it's normal
Public Sub CMComputePlatePosition(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
Const METHOD = "CMComputePlatePosition"
On Error GoTo ErrorHandler
   ' MsgBox MODULE + METHOD
    
    'CA and Pad inputs
    Dim oRefCollCAO As IJDReferencesCollection
    Dim oAttribs As IJDAttributes
    Dim oEditJDArg As IJDEditJDArgument
    Dim oSuppedPort As ISPSSplitAxisPort
    Dim iSuppedPortIdx As SPSMemberAxisPortIndex
    Dim oSuppedPart As ISPSMemberPartPrismatic
    Dim oSupgObj As Object
    Dim oSupgSurf As IJSurface
    Dim oSupgPlane As IJPlane
    Dim oStructPlate As IJStructPlate
    Dim oCollProxy As CollectionProxy
    Dim oOcc As IJDOccurrence
    Dim oMat As New DT4x4
    
    'Member Part
    Dim lCP As Long
    Dim oCPLine As IJLine
    Dim oAxisStart As IJPoint
    Dim oAxisEnd As IJPoint
    Dim oAxisDir As New DVector
    Dim dAx As Double
    Dim dAy As Double
    Dim dAz As Double
    Dim oLocToGloXform As IJDT4x4   ' local to global matrix for member xform
    Dim oGloToLocXform As IJDT4x4   ' global to local
    
    'Surface intersection stuff
    Dim code As Geom3dIntersectConstants
    Dim oIntersectionElems As IJElements
    Dim oP0 As IJPoint
    Dim dP0x As Double
    Dim dP0y As Double
    Dim dP0z As Double
    Dim dP0u As Double
    Dim dP0v As Double
    Dim dN0x As Double
    Dim dN0y As Double
    Dim dN0z As Double
    Dim oIntersectPt As IJPoint
               
    
    'Pad geometry
    Dim oLine As IJLine
    Dim oPadPlane As IJPlane
    Dim oPadCornerPos(1 To 4) As IJPoint
    Dim oPadCurves(1 To 4) As IJLine
    Dim oPadBoundary As IJComplexString
    Dim oPos As IJDPosition, oPos1 As IJDPosition
    Dim oPadXVec As IJDVector
    Dim oPadYVec As IJDVector
    Dim oPadZVec As IJDVector
    Dim oGeom3DFactory As New GeometryFactory
    Dim oPoint3D As IPoints3d
    Dim oLine3D As ILines3d
    Dim oCornerVectorDir As New DVector
    Dim dDotProd As Double
    
    'Pad attributes
    Dim dPlateDepth As Double
    Dim dPlateWidth As Double
    Dim dClearance As Double
    
    Dim idx As Integer
    
    'Error handling and utility stuff
    Dim strError As String
    Dim oP0l As IJPoint
    Dim tempX As Double
    Dim tempY As Double
    Dim tempZ As Double
    Dim temp1X As Double
    Dim temp1Y As Double
    Dim temp1Z As Double
    
    'Dim oLocalizer As IJLocalizer
    
    'Set oLocalizer = New IMSLocalizer.Localizer
    'oLocalizer.intialize App.Path & "\" & App.EXEName
    
    'get the pad plate and its attributes
    Set oStructPlate = pObject
    Set oAttribs = pObject
    Set oCollProxy = oAttribs.CollectionOfAttributes("IJUASPSPlatePartDim")
    dPlateDepth = oCollProxy.Item("Length").Value
    dPlateWidth = oCollProxy.Item("Width").Value
    Set oCollProxy = Nothing
    
    Set oCollProxy = oAttribs.CollectionOfAttributes("IJUAPlateSizeParams")
    dClearance = oCollProxy.Item("Offset").Value
    Set oCollProxy = Nothing
    
    'get the smartoccurrence from the Pad Plate
    Set oOcc = pObject
    'Get the ref coll
    Set oRefCollCAO = GetRefCollFromSmartOccurrence(pObject)
    
    'get the inputs from the ref coll
    Set oEditJDArg = oRefCollCAO.IJDEditJDArgument
    'Check if the oEditJDArg has 2 entities
    If oEditJDArg.GetCount < 2 Then
        'Send error if P0 can not be found. Sorry, we can't position Pad on such a surface.
        'strError = oLocalizer.GetString(IDS_PARTMACROS_PLATE_CANNOT_POSITIONED, "Plate cannot be positioned properly on the supporting member.")
        SPSToDoErrorNotify PlatePartToDoMsgCodelist, TDL_PARTMACROS_PLATE_INVALID_INPUTS, oStructPlate, Nothing
        Err.Raise SPS_MACRO_WARNING 'number of inputs on the plate pad is incorrect
        GoTo ErrorHandler
    End If
    
    Set oSuppedPort = oEditJDArg.GetEntityByIndex(1)
    iSuppedPortIdx = oSuppedPort.portIndex
    Set oSuppedPart = oSuppedPort.Part
    Set oSupgObj = oRefCollCAO.IJDEditJDArgument.GetEntityByIndex(2) 'oSuppingPort.Geometry
   
    'get members CP and use that line instead.
    'get the Member Part CP line.
    lCP = 10  'oSuppedPart.CrossSection.CardinalPoint 'cant use member CP as it misallign the pad with member
    Dim bStart As Boolean
    bStart = True
    If iSuppedPortIdx = SPSMemberAxisEnd Then
        bStart = False
    End If
    
    Set oCPLine = GetTangentLineAtCP(oSuppedPart, lCP, bStart)
    ' Now set the line to infinite to make sur we get correct intersection
    oCPLine.Infinite = True
    'also get the axis of the member pointing into the member.
    Set oAxisStart = oSuppedPart.PointAtEnd(SPSMemberAxisStart)
    Set oAxisEnd = oSuppedPart.PointAtEnd(SPSMemberAxisEnd)
    
    If TypeOf oSupgObj Is IJPlane Then
        'If plane, the create a surafce for later use
        Set oSupgPlane = oSupgObj
        oSupgPlane.GetRootPoint tempX, tempY, tempZ
        oSupgPlane.GetNormal temp1X, temp1Y, temp1Z
        Set oSupgSurf = oGeom3DFactory.Planes3d.CreateByPointNormal(Nothing, tempX, tempY, tempZ, temp1X, temp1Y, temp1Z)
        Set oSupgPlane = Nothing
        
    ElseIf TypeOf oSupgObj Is IJSurface Then
        Set oSupgSurf = oSupgObj
    End If
    
    'Intersect the CP line with surface
     oSupgSurf.Intersect oCPLine, oIntersectionElems, code
    
    If code > 0 Then
        'Send error if P0 can not be found. Sorry, we can't position Pad on such a surface.
        'strError = oLocalizer.GetString(IDS_PARTMACROS_PLATE_CANNOT_POS_CPLINE_NO_INTERSECT, "Plate cannot be positioned properly as the Member's CP line does not intersect with the supporting surface.")
        SPSToDoErrorNotify PlatePartToDoMsgCodelist, TDL_PARTMACROS_PLATE_CANNOT_POS_CPLINE_NO_INTERSECT, oStructPlate, Nothing
        Err.Raise SPS_MACRO_WARNING 'plate is not within the depth of the supporting member
    End If
    
    'We have valid intersection point. Get surface normal at this point
    Set oP0 = oIntersectionElems.Item(1)
    oP0.GetPoint dP0x, dP0y, dP0z
    oSupgSurf.Parameter dP0x, dP0y, dP0z, dP0u, dP0v
    oSupgSurf.Normal dP0u, dP0v, dN0x, dN0y, dN0z
        
    'calculate member axis going into the member
    oAxisStart.GetPoint tempX, tempY, tempZ
    oAxisEnd.GetPoint temp1X, temp1Y, temp1Z
    If iSuppedPortIdx = SPSMemberAxisEnd Then
        dAx = tempX - temp1X
        dAy = tempY - temp1Y
        dAz = tempZ - temp1Z
    Else 'if axisStart is the end we are dealing with
        dAx = -(tempX - temp1X)
        dAy = -(tempY - temp1Y)
        dAz = -(tempZ - temp1Z)
    End If
    oAxisDir.Set dAx, dAy, dAz
    Set oAxisStart = Nothing
    Set oAxisEnd = Nothing
    
    'get the member xform and inverse - for Global to Local and local to Global CS.
    oSuppedPart.Rotation.GetTransform oLocToGloXform
    Set oGloToLocXform = oLocToGloXform.Clone
    oGloToLocXform.Invert
    
    'Now xform the point P0 to LCS
    'Initialize the geom factory automation object for point3d
    Set oPoint3D = oGeom3DFactory.Points3d
    
    Set oP0l = oPoint3D.CreateByPoint(Nothing, dP0x, dP0y, dP0z)
    oP0l.Transform oGloToLocXform
    oP0l.GetPoint tempX, tempY, tempZ
    Set oP0l = Nothing
    Set oGloToLocXform = Nothing
    
   'Get all corner points of the Pad Plate assuming P0 is the center point of the plate.
    Set oPadCornerPos(1) = oPoint3D.CreateByPoint(Nothing, tempX, tempY + dPlateWidth / 2#, tempZ - dPlateDepth / 2#) 'Lower left corner
    Set oPadCornerPos(2) = oPoint3D.CreateByPoint(Nothing, tempX, tempY - dPlateWidth / 2#, tempZ - dPlateDepth / 2#)  'Lower right corner
    Set oPadCornerPos(3) = oPoint3D.CreateByPoint(Nothing, tempX, tempY - dPlateWidth / 2#, tempZ + dPlateDepth / 2#) 'Upper right corner
    Set oPadCornerPos(4) = oPoint3D.CreateByPoint(Nothing, tempX, tempY + dPlateWidth / 2#, tempZ + dPlateDepth / 2#) 'Upper left corner
    
    Set oPoint3D = Nothing

    'Now we can xform the corner points back to global CS
    For idx = 1 To 4
      oPadCornerPos(idx).Transform oLocToGloXform
    Next idx
    
    'Now we have the correct pad size and orientation, but position needs to be adjusted to clear the surface
    
    'Get intersection point with the surface for each such corner point which is oppsite side of surface
    Dim dProjDis As Double
    Dim dDist As Double
    Dim dMemberLength As Double
    dMemberLength = oSuppedPart.cutLength
    dProjDis = 0#
    
    Set oLine3D = oGeom3DFactory.Lines3d
    For idx = 1 To 4
        oPadCornerPos(idx).GetPoint tempX, tempY, tempZ
        Set oLine = oLine3D.CreateByPtVectLength(Nothing, tempX, tempY, tempZ, _
                oAxisDir.x, oAxisDir.y, oAxisDir.z, dMemberLength)
                
        oSupgSurf.Intersect oLine, oIntersectionElems, code
        Set oLine = Nothing
        
        If code = 0 Then
            If oIntersectionElems.Count > 0 Then
                Set oIntersectPt = oIntersectionElems.Item(1)
                
                oIntersectPt.GetPoint tempX, tempY, tempZ
                oCornerVectorDir.Set tempX - dP0x, tempY - dP0y, tempZ - dP0z
                oCornerVectorDir.Length = 1#
                
                dDotProd = oCornerVectorDir.Dot(oAxisDir)
                
                If dDotProd >= 0 Then
                    'if the 2 vectors are in same direction,this point is considered for projecting pad
                    dDist = oIntersectPt.DistFromPt(oPadCornerPos(idx))
                    If dDist > dProjDis Then
                        dProjDis = dDist
                    End If
                End If
                
                Set oPadCornerPos(idx) = Nothing
                Set oIntersectPt = Nothing
            End If
        End If
        Set oCornerVectorDir = Nothing
    Next idx
    
    Set oSupgSurf = Nothing
    Set oLine3D = Nothing
    Set oGeom3DFactory = Nothing
    
'    'Now move the root point and the corner points by this distance along the axisDir
'    'vp2 = vp1 + k.uvn
'    dP0x = dP0x + dProjDis * dAx
'    dP0y = dP0y + dProjDis * dAy
'    dP0z = dP0z + dProjDis * dAz
' This nt needed as it is casuing extar xform. THis should be taken care in the translation part of the matrix below.
    
    'Now create the matrix for the occurrence
    'Cretae vectors along each axis of the Pad- ASsuming xy plane
    ' to be the pad plane and thickness along z.
    'So Zvector is along the surface normal
    Set oPadZVec = New DVector
    oPadZVec.x = dN0x
    oPadZVec.y = dN0y
    oPadZVec.z = dN0z
    
    'Yvector is along the members y-axis, so can be derived from the member loctoglo matrix
    Set oPadYVec = New DVector
    oPadYVec.x = 0#
    oPadYVec.y = 1#
    oPadYVec.z = 0#
    Set oPadYVec = oLocToGloXform.TransformVector(oPadYVec)
    Set oLocToGloXform = Nothing
    
    'yvec could be xcrossz
    Set oPadXVec = New DVector
    Set oPadXVec = oPadYVec.Cross(oPadZVec)
    oPadXVec.Length = 1     ' normalize
    
    'TODO: check if the direction of the vectrs is corrrect for both pad at start and end of member
    'Now load the identity and start building the matrix
    oMat.LoadIdentity
    'local x-axis of the pad is along the surface normal
    oAxisDir.Length = 1
    oMat.IndexValue(0) = oPadXVec.x
    oMat.IndexValue(1) = oPadXVec.y
    oMat.IndexValue(2) = oPadXVec.z
    
    'local y-ax is of the pad
    oMat.IndexValue(4) = oPadYVec.x
    oMat.IndexValue(5) = oPadYVec.y
    oMat.IndexValue(6) = oPadYVec.z
    
    'local z-axis of the pad
    oMat.IndexValue(8) = oPadZVec.x
    oMat.IndexValue(9) = oPadZVec.y
    oMat.IndexValue(10) = oPadZVec.z
    
    'xlation component. center of the plate is P0
    oMat.IndexValue(12) = dP0x
    oMat.IndexValue(13) = dP0y
    oMat.IndexValue(14) = dP0z
    
    oOcc.Matrix = oMat
   
    Set oAxisDir = Nothing
    Set oPadXVec = Nothing
    Set oPadYVec = Nothing
    Set oPadZVec = Nothing
    Set oMat = Nothing
    
    'Set oLocalizer = Nothing
Exit Sub

ErrorHandler:
    ' For errors logged with E_FAIL, a todo list error will be generated so we should not
    '   be logging anything to the error log
    If Err.Number = SPS_MACRO_WARNING Then
        Err.Raise SPS_MACRO_WARNING
    Else
        Err.Raise ReportError(Err, m_strSourceFile, METHOD).Number
    End If
End Sub





'
' The following methods are generic for all the Custom assembly
'
'
Public Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, ByVal defParams As Variant, ByVal ActiveConnection As Object) As Object
  ' This method is in charge of the creation of the symbol definition object
  ' You can keep the current design unchanged for basic VB symbol definition.
  On Error GoTo ErrorHandler
  sMethod = "IJDUserSymbolServices_InstanciateDefinition"
  sERROR = ""
    
  sERROR = "Create the item definition"
  Dim pDefinition As IJDSymbolDefinition
  Dim pFact As IJCAFactory
  Set pFact = New CAFactory
  Set pDefinition = pFact.CreateCAD(ActiveConnection)
  
  ' Set definition progId and codebase
  pDefinition.ProgId = m_ProgId
  pDefinition.CodeBase = CodeBase
    
  ' Initialize the definition
  sERROR = "Initialize and named definition"
  IJDUserSymbolServices_InitializeSymbolDefinition pDefinition
  pDefinition.Name = IJDUserSymbolServices_GetDefinitionName(defParams)
  
  ' Persistence behavior
  pDefinition.SupportOnlyOption = igSYMBOL_NOT_SUPPORT_ONLY
'  pDefinition.MetaDataOption = igSYMBOL_DYNAMIC_METADATA
    
  'returned symbol definition
  Set IJDUserSymbolServices_InstanciateDefinition = pDefinition
  
  Exit Function
ErrorHandler:
  MsgBox Err.Source & ": " & Trim$(Str$(Err.Number)) & " - " & Err.Description & " - " & sERROR
  Debug.Assert False
  Err.Raise SYMBOL_E_FAIL, MODULE + " - " + sMethod + " - " + sERROR
End Function
Public Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
  ' Name should be unique
  IJDUserSymbolServices_GetDefinitionName = m_ProgId
End Function
Public Sub IJDUserSymbolServices_InvokeRepresentation(ByVal sblOcc As Object, ByVal repName As String, ByVal outputcoll As Object, ByRef arrayOfInputs())
 ' Obsolete method.
End Sub

Public Function IJDUserSymbolServices_EditOccurence(ByRef pSymbolOccurence As Object, ByVal transactionMgr As Object) As Boolean
 ' Obsolete method. Instead you can record your custom command within the definition (see IJDCommandDescription interface)
 IJDUserSymbolServices_EditOccurence = False
End Function


'Attribute Management Stuff
Private Function IJUserAttributeMgmt_OnAttributeChange(ByVal pIJDAttrs As SPSMembers.IJDAttributes, ByVal CollAllDisplayedValues As Object, ByVal pAttrToChange As SPSMembers.IJAttributeDescriptor, ByVal varNewAttrValue As Variant) As String

Const METHOD = "IJUserAttributeMgmt_OnAttributeChange"
On Error GoTo ErrorHandler
    IJUserAttributeMgmt_OnAttributeChange = "ERROR"
    
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
            IJUserAttributeMgmt_OnAttributeChange = ErrStr
            Exit Function
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
    IJUserAttributeMgmt_OnPreLoad = "ERROR"
    bOnPreLoad = True ' optimization to avoid value validation in OnAttrChange
    
    Dim i As Integer
    Dim pAttrColl As Collection
    Dim pAttrDescr As IJAttributeDescriptor
    Dim attrName As String
    Dim ErrStr As String

    Dim lSizingRule As Long
    
    'get the pad plate and its attributes
    lSizingRule = pIJDAttrs.CollectionOfAttributes("IJUAPlateSizeParams").Item("SizingRule").Value
    
    Set pAttrColl = CollAllDisplayedValues
    
    'Make the length, width and cornerradius readonly if sizing rule = ByRule
    If Not lSizingRule = 2 Then
        For i = 1 To pAttrColl.Count
            Set pAttrDescr = pAttrColl.Item(i)
            If pAttrDescr.attrName = "Width" Or pAttrDescr.attrName = "Length" Or pAttrDescr.attrName = "CornerRadius" Then
                pAttrDescr.AttrState = pAttrDescr.AttrState Or AttributeDescriptor_ReadOnly
            End If
        Next
    End If
    
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
    UserAttributeMgmt_Validate = "ERROR"

    Dim dInputs As IJDInputs
    Dim CurrentInput As IJDInput
    Dim oAttribute As IJDAttribute
    Dim PC As DParameterContent
    Dim bvalid As Boolean
    Dim oSymbolOcc As IJDSymbol
    Set oSymbolOcc = pIJDAttrs
    Dim oSymbolDef As IJDSymbolDefinition
    Dim ErrMessage As String
    Dim oLocalizer As IJLocalizer
    
    Set oLocalizer = New IMSLocalizer.Localizer
    oLocalizer.Initialize App.Path & "\" & App.EXEName
    
    Set oSymbolDef = oSymbolOcc.IJDSymbolDefinition(2)
    Set dInputs = oSymbolDef.IJDInputs
    Set PC = New DParameterContent
    
    'the pad dimensions (Length /Width ) need to be all  > 0, except for cornerradius which can be 0
    
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
        UserAttributeMgmt_Validate = oLocalizer.GetString(IDS_PARTMACROS_SCHEMAERROR, "SchemaERROR")
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
                    If (sAttributeName = "Length") Or (sAttributeName = "Width") Then 'length and width should be >0
                        If (varAttributeValue <= 0#) Then
                            UserAttributeMgmt_Validate = oLocalizer.GetString(IDS_PARTMACROS_VALUE_NEGATIVE, "Negative Attribute Value")
                            Set oAttributeMetaData = Nothing
                            Exit Function
                        End If
                    ElseIf sAttributeName = "CornerRadius" Then 'cornerradius  should not be < 0
                        If (varAttributeValue < 0#) Then
                            UserAttributeMgmt_Validate = oLocalizer.GetString(IDS_PARTMACROS_VALUE_NEGATIVE, "Negative Attribute Value")
                            Set oAttributeMetaData = Nothing
                            Exit Function
                        End If
                    End If
            
            End Select

        End If
    Next
    
    UserAttributeMgmt_Validate = ""
    Set oAttributeMetaData = Nothing
    Set oLocalizer = Nothing

Exit Function
ErrorHandler:  HandleError MODULE, METHOD
End Function

'*************************************************************************
'Function
'CMMigrateAggregator
'
'Abstract
'Callback method to migrate aggergator in case of a split.
'
'Arguments
'pAggregatorDescription -   provides the aggregator definition information provided
'                           at the time when the assembly connection was bulkloaded
'                           (see IJDUserSymbolServices_InitializeSymbolDefinition)
'
'Exceptions
'
'***************************************************************************
Public Sub CMMigrateAggregator(oAggregatorDesc As IJDAggregatorDescription, oMigrateHelper As IJMigrateHelper)
    Const METHOD = "CMMigrateAggregator"
    On Error GoTo ErrorHandler
    Dim oSmartOcc As IJSmartOccurrence
    Dim oRefColl As IJDReferencesCollection
    Dim oObjectsReplacing() As Object
    Dim bIsInputMigrated As Boolean
    
    Set oSmartOcc = oAggregatorDesc.CAO
    Set oRefColl = GetRefCollFromSmartOccurrence(oSmartOcc)
    
    MigrateRefColl oRefColl, oMigrateHelper, oObjectsReplacing, bIsInputMigrated
    
    If bIsInputMigrated Then
        'If any of the inputs are indeed migrated, reset them on the ref coll
        Call oRefColl.IJDEditJDArgument.RemoveAll
        oRefColl.IJDEditJDArgument.SetEntity 1, oObjectsReplacing(1), CONST_ISPSSplitAxisEndPort, "RefColl"
        oRefColl.IJDEditJDArgument.SetEntity 2, oObjectsReplacing(2), ConstIJSurface, "SurfaceToRC_DEST"
    End If
    
    Exit Sub

ErrorHandler:
    HandleError MODULE, METHOD
End Sub

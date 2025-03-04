VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CompFrameRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'***************************************************************************
'  Copyright (C) 2000, Intergraph Corporation.  All rights reserved.
'
'  Project: GSCADStrSpaceNameRules
'
'  Abstract: The file contains an implementation of the Frame naming rule
'            for the Compartment objectin Compartmentation UE.
'           It finds out the bounding Frames of the range of the compartment and
'           assigns the names of the compartment as per these frames.
'
'  History:
'  Arnab            10th September 2004                 Creation
'***************************************************************************

Option Explicit

Dim m_oErrors As IJEditErrors

Private Const E_FAIL = -2147467259

Implements IJNameRule

Private Const MODULE = "CompFrameRule: "
Private Const strCountFormat = "0000"   'define fixed-width number field

Private Sub Class_Initialize()
    Set m_oErrors = New IMSErrorLog.JServerErrors
End Sub

Private Sub Class_Terminate()
    Set m_oErrors = Nothing
End Sub

'*********************************************************************************************
' Description:
'   Creates a name for the object passed in. The name is based on the parents
'   name and object name. It is assumed that all Naming Parents and the Object implement IJNamedItem.
'   The Naming Parents are added in AddNamingParents() of the same interface.
'   Both these methods are called from the naming rule semantic.
'
' Notes: ZoneName  = Zone description + Unique Index
'***************************************************************************
Private Sub IJNameRule_ComputeName(ByVal pEntity As Object, ByVal pParents As IJElements, ByVal pActiveEntity As Object)
    Const METHOD = "IJNameRule_ComputeName"
    On Error GoTo ErrorHandler
    Dim oNamedItem                  As IJNamedItem
    Dim strName                     As String

    Dim JContext As IJContext
    Dim oDBTypeConfig As IJDBTypeConfiguration
    Dim oConnectMiddle As IJDAccessMiddle
    Dim strModelDBID As String
    Dim oModelResourceMgr As IUnknown
    Dim oNameCounter As IJNameCounter
    Dim strLocation As String
    Dim nCount As Long
    
    'Get the connection to the model database
    Set JContext = GetJContext()
    
    Set oDBTypeConfig = JContext.GetService("DBTypeConfiguration")
    Set oConnectMiddle = JContext.GetService("ConnectMiddle")
    
    strModelDBID = oDBTypeConfig.get_DataBaseFromDBType("Model")
    Set oModelResourceMgr = oConnectMiddle.GetResourceManager(strModelDBID)
    
    Set oNameCounter = New GSCADNameGenerator.NameGeneratorService
    Set oNamedItem = pEntity
    
    
    strName = GetBoundingFrameNames(oModelResourceMgr, pEntity)
    
    strLocation = vbNullString
    'GetCountEx:Returns the number of occurrence of a string in addtion to the LocationID
    nCount = oNameCounter.GetCountEx(oModelResourceMgr, strName, strLocation)
    
    'Add LocationID, if available
    If strLocation <> vbNullString Then
        strName = strName & "-" & strLocation & "-" & Format(nCount, strCountFormat)
    Else
        strName = strName & "-" & Format(nCount, strCountFormat)
    End If
    
    oNamedItem.Name = strName
    
    Set oNamedItem = Nothing
    Set JContext = Nothing
    Set oDBTypeConfig = Nothing
    Set oConnectMiddle = Nothing
    Set oModelResourceMgr = Nothing
    Set oNameCounter = Nothing
    
Exit Sub
ErrorHandler:
    m_oErrors.Add Err.Number, "CompSpaceRule::IJNameRule_ComputeName", Err.Description
    Err.Raise CompartLogError(Err, MODULE, METHOD, , , CMPART_CUSTOMERRORS_NAMINGRULES_FAILED_FRAMERULE)
End Sub

'****************************************************************************************************
'Description
'   All the Naming Parents that need to participate in an objects naming are added here to the
'   IJElements collection. Dummy function which does nothing
'****************************************************************************************************
Private Function IJNameRule_GetNamingParents(ByVal pEntity As Object) As IJElements
    Const METHOD = "IJNameRule_GetNamingParents"
    On Error GoTo ErrorHandler
    Set IJNameRule_GetNamingParents = New IMSCoreCollections.JObjectCollection
    
    Dim oCompartEntity          As IJCompartEntity
'    Dim oCompartEntity          As IJRangeAlias
'
'    Set oRange = pEntity
'
''    Set oCompartGeom = oCompartEntity.SolidGeometry
'
'    If Not oRange Is Nothing Then
'        IJNameRule_GetNamingParents.Add oRange
'    End If
'
'    Set oCompartEntity = Nothing
'    Set oRange = Nothing

Exit Function
ErrorHandler:
    m_oErrors.Add Err.Number, "CompFrameRule::IJNameRule_GetNamingParents", Err.Description
    Err.Raise CompartLogError(Err, MODULE, METHOD)
End Function

Private Function GetBoundingFrameNames(oModelResourceMgr As IUnknown, oEntity As Object) As String
    Const METHOD = "GetBoundingFrameNames"
    On Error GoTo ErrorHandler
    
    Const dTol = 0.000000001
    
    Dim oFrameMiddleHelper              As SPGMiddleHelper
    Dim oRange                          As IJRangeAlias
    Dim oFrameColl                      As IJElements
    Dim oNamedItem                      As IJNamedItem
    Dim oFirstFrame                     As ISPGNavigate
    Dim oLastFrame                      As ISPGNavigate
    Dim oBoundingFrameStart             As IJNamedItem
    Dim oBoundingFrameEnd               As IJNamedItem
    Dim eRefPosition                    As ReferencePosition
    Dim strName                         As String
    Dim boolFramesExist                 As Boolean
    
    Set oRange = oEntity
    
    eRefPosition = GetReferencePosition(oRange)

    If eRefPosition = rpPortside Then
        strName = "Port"
    ElseIf eRefPosition = rpStarboard Then
        strName = "StarBoard"
    Else
        strName = "Center"
    End If
    
    Set oFrameMiddleHelper = New SPGMiddleHelper
    
    'For X Axis
    oFrameMiddleHelper.EnumPlanesInRange oModelResourceMgr, oRange, Nothing, "", x, True, 4, oFrameColl
    If oFrameColl.Count > 0 Then
        Set oFirstFrame = oFrameColl.Item(1)
        Set oLastFrame = oFrameColl.Item(oFrameColl.Count)
        
        If IsPlaneTouchingRangeBox(oRange, oFirstFrame) = True Then
            Set oBoundingFrameStart = oFirstFrame
        Else
            oFirstFrame.GetReference Previous, NestingLevelType.Any, oBoundingFrameStart
        End If
        
        If Not oBoundingFrameStart Is Nothing Then
            GetBoundingFrameNames = GetBoundingFrameNames & oBoundingFrameStart.Name
        End If
        
        If IsPlaneTouchingRangeBox(oRange, oLastFrame) = True Then
            Set oBoundingFrameEnd = oLastFrame
        Else
            oLastFrame.GetReference ReferenceType.Next, NestingLevelType.Any, oBoundingFrameEnd
        End If
        
        If Not oBoundingFrameEnd Is Nothing Then
            If Not oBoundingFrameStart Is Nothing Then
                GetBoundingFrameNames = GetBoundingFrameNames & "-" & oBoundingFrameEnd.Name
            Else
                GetBoundingFrameNames = GetBoundingFrameNames & oBoundingFrameEnd.Name
            End If
        End If
    End If
    
    If (Not oBoundingFrameStart Is Nothing Or Not oBoundingFrameEnd Is Nothing) Then
        boolFramesExist = True
    End If
    
    'For Y Axis
    oFrameMiddleHelper.EnumPlanesInRange oModelResourceMgr, oRange, Nothing, "", y, True, 4, oFrameColl
    If oFrameColl.Count > 0 Then
        Set oFirstFrame = oFrameColl.Item(1)
        Set oLastFrame = oFrameColl.Item(oFrameColl.Count)
        
        If IsPlaneTouchingRangeBox(oRange, oFirstFrame) = True Then
            Set oBoundingFrameStart = oFirstFrame
        Else
            oFirstFrame.GetReference Previous, NestingLevelType.Any, oBoundingFrameStart
        End If
        
        If Not oBoundingFrameStart Is Nothing Then
            If boolFramesExist = True Then
                GetBoundingFrameNames = GetBoundingFrameNames & "::" & oBoundingFrameStart.Name
            Else
                GetBoundingFrameNames = GetBoundingFrameNames & oBoundingFrameStart.Name
            End If
        End If
        
        If IsPlaneTouchingRangeBox(oRange, oLastFrame) = True Then
            Set oBoundingFrameEnd = oLastFrame
        Else
            oLastFrame.GetReference ReferenceType.Next, NestingLevelType.Any, oBoundingFrameEnd
        End If
        
         If Not oBoundingFrameEnd Is Nothing Then
            If Not oBoundingFrameStart Is Nothing Then
                GetBoundingFrameNames = GetBoundingFrameNames & "-" & oBoundingFrameEnd.Name
            Else
                GetBoundingFrameNames = GetBoundingFrameNames & oBoundingFrameEnd.Name
            End If
        End If
    End If
    
    boolFramesExist = False
    
    
    If (Not oBoundingFrameStart Is Nothing Or Not oBoundingFrameEnd Is Nothing) Then
        boolFramesExist = True
    End If
    
    'For Z Axis
    oFrameMiddleHelper.EnumPlanesInRange oModelResourceMgr, oRange, Nothing, "", z, True, 4, oFrameColl
    If oFrameColl.Count > 0 Then
        Set oFirstFrame = oFrameColl.Item(1)
        Set oLastFrame = oFrameColl.Item(oFrameColl.Count)
        
        If IsPlaneTouchingRangeBox(oRange, oFirstFrame) = True Then
            Set oBoundingFrameStart = oFirstFrame
        Else
            oFirstFrame.GetReference Previous, NestingLevelType.Any, oBoundingFrameStart
        End If
        
         If boolFramesExist = True Then
                GetBoundingFrameNames = GetBoundingFrameNames & "::" & oBoundingFrameStart.Name
            Else
                GetBoundingFrameNames = GetBoundingFrameNames & oBoundingFrameStart.Name
            End If
        
        If IsPlaneTouchingRangeBox(oRange, oLastFrame) = True Then
            Set oBoundingFrameEnd = oLastFrame
        Else
            oLastFrame.GetReference ReferenceType.Next, NestingLevelType.Any, oBoundingFrameEnd
        End If
        
          If Not oBoundingFrameEnd Is Nothing Then
            If Not oBoundingFrameStart Is Nothing Then
                GetBoundingFrameNames = GetBoundingFrameNames & "-" & oBoundingFrameEnd.Name
            Else
                GetBoundingFrameNames = GetBoundingFrameNames & oBoundingFrameEnd.Name
            End If
        End If
    End If

    If GetBoundingFrameNames = "" Then
        GetBoundingFrameNames = strName & GetBoundingFrameNames
    Else
        GetBoundingFrameNames = strName & "_" & GetBoundingFrameNames
    End If

    
Exit Function
ErrorHandler:
    m_oErrors.Add Err.Number, "CompFrameRule::GetBoundingFrameNames", Err.Description
    Err.Raise CompartLogError(Err, MODULE, METHOD)
End Function

'Returns True if the plane just touches the range. Returns false if it intersects the range.
Private Function IsPlaneTouchingRangeBox(oRange As IJRangeAlias, oFrame As IJPlaneFacelet) As Boolean
    Const METHOD = "IsPlaneTouchingRangeBox"
    On Error GoTo ErrorHandler
    
    Dim gRangeBox           As GBox
    Dim oPosition(1 To 8)   As IJDPosition
    Dim i                   As Long
    Dim oRootPt             As IJDPosition
    Dim x                   As Double
    Dim y                   As Double
    Dim z                   As Double
    Dim oNormal             As IJDVector
    Dim oNewVector          As IJDVector
    Dim direction           As Double
    
    gRangeBox = oRange.GetRange
    IsPlaneTouchingRangeBox = True
    
    For i = 1 To 8
        Set oPosition(i) = New DPosition
    Next i
    
    oPosition(1).Set gRangeBox.m_low.x, gRangeBox.m_low.y, gRangeBox.m_low.z
    oPosition(2).Set gRangeBox.m_low.x, gRangeBox.m_low.y, gRangeBox.m_high.z
    oPosition(3).Set gRangeBox.m_low.x, gRangeBox.m_high.y, gRangeBox.m_low.z
    oPosition(4).Set gRangeBox.m_low.x, gRangeBox.m_high.y, gRangeBox.m_high.z
    oPosition(5).Set gRangeBox.m_high.x, gRangeBox.m_low.y, gRangeBox.m_low.z
    oPosition(6).Set gRangeBox.m_high.x, gRangeBox.m_low.y, gRangeBox.m_high.z
    oPosition(7).Set gRangeBox.m_high.x, gRangeBox.m_high.y, gRangeBox.m_low.z
    oPosition(8).Set gRangeBox.m_high.x, gRangeBox.m_high.y, gRangeBox.m_high.z
    
    oFrame.GetRootPoint x, y, z
    
    Set oRootPt = New DPosition
    oRootPt.Set x, y, z
    
    Set oNormal = New DVector
    oFrame.GetNormal x, y, z
    oNormal.Set x, y, z
    
    Set oNewVector = oPosition(1).Subtract(oRootPt)
    direction = oNewVector.Dot(oNormal)
    
    For i = 2 To 8
        Set oNewVector = oPosition(i).Subtract(oRootPt)
        If ((oNewVector.Dot(oNormal)) * direction < 0) Then
            IsPlaneTouchingRangeBox = False
            Exit For
        End If
    Next i
    
    For i = 1 To 8
        Set oPosition(i) = Nothing
    Next i

Exit Function
ErrorHandler:
    m_oErrors.Add Err.Number, "CompFrameRule::IsPlaneTouchingRangeBox", Err.Description
    Err.Raise CompartLogError(Err, MODULE, METHOD)
End Function

'Returns True if the plane just touches the range. Returns false if it intersects the range.
Private Function GetReferencePosition(oRange As IJRangeAlias) As ReferencePosition
    Const METHOD = "GetReferencePosition"
    On Error GoTo ErrorHandler
    
    Dim gRangeBox           As GBox
    
    gRangeBox = oRange.GetRange
    GetReferencePosition = rpUndefined

    If gRangeBox.m_low.y >= 0 And gRangeBox.m_high.y >= 0 Then
        GetReferencePosition = ReferencePosition.rpPortside
    ElseIf gRangeBox.m_low.y <= 0 And gRangeBox.m_high.y <= 0 Then
        GetReferencePosition = ReferencePosition.rpStarboard
    End If
    

Exit Function
ErrorHandler:
    m_oErrors.Add Err.Number, "CompFrameRule::GetReferencePosition", Err.Description
    Err.Raise CompartLogError(Err, MODULE, METHOD)
End Function

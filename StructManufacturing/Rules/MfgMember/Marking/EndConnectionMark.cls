VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EndConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2002, Intergraph Corp.  All rights reserved.
'
' Project: MfgMemberMarking
' Module: Marking Rules
'
' Description:  Determines the marking settings for the mfg profile
'
' Author:
'
' Comments:
'*******************************************************************************
Option Explicit

Implements IJDMfgSystemMarkingRule

Private Const MODULE = "MfgMemberMarking.EndConnection"

Private Sub Class_Initialize()
    'Initialize the most used objects and helpers
    Helpers.Initialize
End Sub

Private Sub Class_Terminate()
    'Clean up
    Helpers.UnInitialize
End Sub

Private Function IJDMfgSystemMarkingRule_CreateAfterUnfold(ByVal Part As Object, ByVal UpSide As Long, ByVal bSelectiveRecompute As Boolean, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias) As GSCADMfgRulesDefinitions.IJMfgGeomCol2d

End Function

Private Function IJDMfgSystemMarkingRule_CreateBeforeUnfold(ByVal Part As Object, ByVal UpSide As Long, ByVal bSelectiveRecompute As Boolean, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias) As GSCADMfgRulesDefinitions.IJMfgGeomCol3d
    Const METHOD = "EndConnection: IJDMfgSystemMarkingRule_CreateBeforeUnfold"
    
    On Error GoTo ErrorHandler
    
    'This line is added because the method "Connection_ContourProfileEnd" is not implemented for beam part.
    'It can be removed, once it is implemented
    If TypeOf Part Is IJBeamPart Then Exit Function
    
    'Create the SD profile Wrapper and initialize it
    Dim oSDMemberWrapper As Object
    'Create the SD profile Wrapper and initialize it
    Set oSDMemberWrapper = New StructDetailObjects.MemberPart
    Set oSDMemberWrapper.object = Part
    
    Dim oProfileWrapper As MfgRuleHelpers.ProfilePartHlpr
    Set oProfileWrapper = New MfgRuleHelpers.ProfilePartHlpr
    Set oProfileWrapper.object = Part
    
    'Get the Profile Part Physically Connected Objects
    Dim oConObjsCol As Collection
    Set oConObjsCol = GetPhysicalConnectionData(Part, ReferenceObjColl, False)
    
    If oConObjsCol Is Nothing Then
        'Since there is no connecting structure we can leave the rule
        GoTo CleanUp
    End If
    
    Dim Item As Object
    Dim oConnectionData As ConnectionData
    
    Dim nIndex As Long, nWBIndex As Long
    Dim oSDConProfileWrapper As Object
    Dim oSDPhysicalConn As StructDetailObjects.PhysicalConn
    
    ' Loop thru each Physical Connections
    Dim bContourEnd As Boolean
    Dim oWBCollection As Collection
    Dim oWB As IJWireBody
    Dim oCS As IJComplexString
    Dim oSystemMark As IJMfgSystemMark
    Dim oObjSystemMark As IUnknown
    Dim oMoniker As IMoniker
    Dim oMarkingInfo As MarkingInfo
    Dim oGeom3d As IJMfgGeom3d
    Dim lGeomCount As Long
    lGeomCount = 1
    
    Dim oResourceManager As IUnknown
    Set oResourceManager = GetActiveConnection.GetResourceManager(GetActiveConnectionName)
            
    Dim oGeomCol3d As IJMfgGeomCol3d
    Set oGeomCol3d = m_oGeomCol3dFactory.Create(oResourceManager)
    
    Dim oMfgProfilePart As IJMfgProfilePart
    Dim oMfgProfileOutput As IJMfgProfileOutput
        
    If Not oProfileWrapper.ProfileHasMfgPart(oMfgProfilePart) Then Exit Function
    Set oMfgProfileOutput = oMfgProfilePart
    
    Dim oFaces() As Variant
    Dim varFace As Variant
    Dim lFace As Long
    
    Dim oElems As IJElements
    
    'Getting All the Active Faces of the Member to be Manufactured on which marks can be put
    oMfgProfileOutput.GetPartFaces StructMfgAnySide, StructMfgAsAny, oFaces
    
    If Not (oConObjsCol Is Nothing) Then
        For nIndex = 1 To oConObjsCol.Count
            oConnectionData = oConObjsCol.Item(nIndex)
            
            Dim oStructPort As IJStructPort
            Dim res As Integer
            Set oStructPort = oConnectionData.ConnectingPort
            
            res = oStructPort.ContextID And CTX_LATERAL
                
            If res <= 0 Then GoTo NextItem
            
            'Check if the connected object is a profile
            If Not (TypeOf oConnectionData.ToConnectable Is IJProfilePart Or _
                     TypeOf oConnectionData.ToConnectable Is ISPSMemberPartPrismatic) Then
                 GoTo NextItem
            End If
            
            'Initialize the profile wrapper and the Physical Connection wrapper
            If TypeOf oConnectionData.ToConnectable Is IJProfilePart Then
                Set oSDConProfileWrapper = New StructDetailObjects.ProfilePart
            Else
                Set oSDConProfileWrapper = New StructDetailObjects.MemberPart
            End If
            Set oSDConProfileWrapper.object = oConnectionData.ToConnectable
                
            Set oSDPhysicalConn = New StructDetailObjects.PhysicalConn
            Set oSDPhysicalConn.object = oConnectionData.AppConnection
            bContourEnd = oSDMemberWrapper.Connection_ContourProfileEnd(oConnectionData.AppConnection, oWBCollection)
            If ((bContourEnd = True) And Not (oWBCollection Is Nothing)) Then
                'Convert the IJWireBody to a IJComplexString
                'Fix for TR 54992
                'iterate through all the elements in the collection
'                For nWBIndex = 1 To oWBCollection.Count
'                    Set oWB = oWBCollection.Item(nWBIndex)

                If oWBCollection.Count = 0 Then
                    'Since elements count is zero
                    GoTo NextItem
                End If

                'If we iterate through each wire body and convert it into complex string then
                'it will result into multiple Marking lines representing same connection which
                'is not wrong but the part monitor will have so many lables around the connection mark
                'showing the Connected part name.
                'Work around is to converting the wirebodies collection into one single complex string
                'and create one Geom3d marking line for the one End connection. This will result into
                'one Marking line and only one lable showing the Part name in Part monitor.
                'Fix for TR 63496.
                    
                'Merging collection of wirebodies into single wirebody
                Set oWB = m_oMfgRuleHelper.MergeWireBodyCollection(oWBCollection)
                
                
                'Finding the Intersection of the Contour Wirebody on the Ports to be marked
                
                
                If UBound(oFaces) >= LBound(oFaces) Then
                    For Each varFace In oFaces
                        lFace = varFace
                        
                        If lFace = JXSEC_BOTTOM Or lFace = JXSEC_TOP Or lFace = JXSEC_WEB_LEFT Or lFace = JXSEC_WEB_RIGHT Then
                            Dim oConnPortSurfaceBody As IJSurfaceBody
                            Dim oIntersectedWB As IJWireBody
                            Set oConnPortSurfaceBody = oProfileWrapper.GetSurfacePort(lFace).Geometry
                            
                            On Error Resume Next
                            Set oIntersectedWB = GetIntersection(oConnPortSurfaceBody, oWB)
                            On Error GoTo ErrorHandler
                            
                            If Not oIntersectedWB Is Nothing Then
                            
                                'Converting wire body into complex string
                                Set oCS = m_oMfgRuleHelper.WireBodyToComplexString(oIntersectedWB)
                                
                                oCS.GetCurves oElems
                                
                                If oElems.Count > 1 Then
                    
                                    'Create a SystemMark object to store additional information
                                    Set oSystemMark = m_oSystemMarkFactory.Create(oResourceManager)
                                    
                                    'Dim oSubPort As IJStructPort
                                    'Set oSubPort = oProfileWrapper.GetProfileSubPort(oConnectionData.ConnectingPort, oConnectionData.AppConnection)
                                    
                                    'Set the marking side
                                    'Dim MarkingFace As Long
                                    'MarkingFace = oProfileWrapper.GetSide(oSubPort)
                                    oSystemMark.SetMarkingSide lFace
                                    
                                    'Set oSubPort = Nothing
                        
                                    'QI for the MarkingInfo object on the SystemMark
                                    Set oMarkingInfo = oSystemMark
                        
                                    oMarkingInfo.name = oSDConProfileWrapper.name
                                    oMarkingInfo.thickness = oSDConProfileWrapper.WebThickness
                        
                                    oMarkingInfo.FittingAngle = oSDPhysicalConn.MountingAngle
                        
                                    Set oGeom3d = m_oGeom3dFactory.Create(oResourceManager)
                                    oGeom3d.PutGeometry oCS
                                    oGeom3d.PutGeometrytype STRMFG_END_MARK
                                    oGeom3d.FaceId = lFace
                                    Set oObjSystemMark = oSystemMark
                                    Set oMoniker = m_oMfgRuleHelper.GetMoniker(oConnectionData.AppConnection)
                                    oGeom3d.PutMoniker oMoniker
                            
                                    oSystemMark.Set3dGeometry oGeom3d
                                        
                                    oGeomCol3d.AddGeometry lGeomCount, oGeom3d
                                    lGeomCount = lGeomCount + 1
                                End If
                            End If
                            Set oConnPortSurfaceBody = Nothing
                            Set oIntersectedWB = Nothing
                        End If
                    Next
                End If
            End If
NextItem:
                
            Set oWBCollection = Nothing
            Set oWB = Nothing
            Set oSystemMark = Nothing
            Set oMarkingInfo = Nothing
            Set oGeom3d = Nothing
            Set oSDPhysicalConn = Nothing
        Next nIndex
    End If
    'Return the 3d collection
    Set IJDMfgSystemMarkingRule_CreateBeforeUnfold = oGeomCol3d
    
CleanUp:
    Set oSDMemberWrapper = Nothing
    Set oProfileWrapper = Nothing
    Set oConObjsCol = Nothing
    Set Item = Nothing
    Set oCS = Nothing
    Set oMoniker = Nothing
    Set oGeomCol3d = Nothing
    Set oElems = Nothing
    Set oMfgProfilePart = Nothing
    Set oMfgProfileOutput = Nothing

Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 2003, , "RULES")
    GoTo CleanUp
End Function


 

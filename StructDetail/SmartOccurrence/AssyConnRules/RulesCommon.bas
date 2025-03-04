Attribute VB_Name = "RulesCommon"
Option Explicit

Public Const INPUT_BOUNDED_OR_PENETRATED_OBJECT = "ConnectedObject1"
Public Const INPUT_BOUNDING_OR_PENETRATING_OBJECT = "ConnectedObject2"

Public Const IID_IJPlate = "{53CF4EA0-91BF-11D1-BE56-080036B3A103}"
Public Const IID_IJStructureMaterial = "{E790A7C0-2DBA-11D2-96DC-0060974FF15B}"
Public Const IID_IJCollarPart = "{138C021D-7089-11D5-B0D9-006008676515}"

Public Const QUES_ENDCONDITION As String = "EndCondition"
Public Const CL_ENDCONDITION As String = "EndConditionCodeList"

Public Const QUES_ENDCUTTYPE As String = "EndCutType"
Public Const CL_ENDCUTTYPE As String = "EndCutTypeCodeList"


' End cut EndCondition Constants
Public Const gsFixed = "Fixed"
Public Const gsFree = "Free"
Public Const gsFlangeFree = "FlangeFree"

' EndCut Type Constants
Public Const gsW = "W"
Public Const gsC = "C"
Public Const gsF = "F"
Public Const gsS = "S"
Public Const gsFV = "FV"
Public Const gsR = "R"
Public Const gsRV = "RV"

' Assembly Method Constants
Public Const gsDrop = "Drop"
Public Const gsSlide = "Slide"


' Stress Level constants
Public Const gsHigh = "High"
Public Const gsMedium = "Medium"
Public Const gsLow = "Low"

Public Enum enmProfilePortType
    PROFILE_PORTTYPE_EDGE = 1
    PROFILE_PORTTYPE_FACE
End Enum
    
Public Enum enmPortBasicContext
    PORT_BASIC_CONTEXT_BASE = 1
    PORT_BASIC_CONTEXT_OFFSET
    PORT_BASIC_CONTEXT_LATERAL
End Enum

Public Const CMLIBRARY_ASSYCONNRULES As String = "AssyConnRules.AssyConnDefCM"
Public Const CMLIBRARY_ASSYCONNSEL As String = "AssyConnRules.AssyConnSelCM"
Public Const LIBRARY_SOURCE_ID = "AssyConnRules.PlateByPlateSel"

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\AssyConnRules\RulesCommon.bas"

Public Function ProfilePartPortType(ProfilePartCrossSectionType As String, _
                                                   CrossSectionEntity As IMSProfileEntity.JXSEC_CODE) As enmProfilePortType
    On Error GoTo ErrorHandler

    Dim strError As String
    
    Select Case ProfilePartCrossSectionType
        Case "B"
            If CrossSectionEntity = JXSEC_BOTTOM Then
                ProfilePartPortType = PROFILE_PORTTYPE_EDGE
            Else
                ProfilePartPortType = PROFILE_PORTTYPE_FACE
            End If
            
        Case "FB"
            If CrossSectionEntity = JXSEC_TOP Or _
                CrossSectionEntity = JXSEC_BOTTOM Then
                ProfilePartPortType = PROFILE_PORTTYPE_EDGE
            Else
                ProfilePartPortType = PROFILE_PORTTYPE_FACE
            End If
            
        Case "EA", "UA"
            If CrossSectionEntity = JXSEC_BOTTOM Or _
                CrossSectionEntity = JXSEC_WEB_RIGHT_BOTTOM_CORNER Or _
                CrossSectionEntity = JXSEC_TOP_FLANGE_RIGHT Or _
                CrossSectionEntity = JXSEC_TOP_FLANGE_RIGHT_BOTTOM_CORNER Then
                ProfilePartPortType = PROFILE_PORTTYPE_EDGE
            Else
                ProfilePartPortType = PROFILE_PORTTYPE_FACE
            End If
        
        Case "T", "TSType", "BUT", "BUTL2", "T_XType", "T_Xtype"
            If CrossSectionEntity = JXSEC_BOTTOM Or _
                CrossSectionEntity = JXSEC_TOP_FLANGE_LEFT Or _
                CrossSectionEntity = JXSEC_TOP_FLANGE_RIGHT Then
                ProfilePartPortType = PROFILE_PORTTYPE_EDGE
            Else
                ProfilePartPortType = PROFILE_PORTTYPE_FACE
            End If
        
        Case "I", "ISType"
            If CrossSectionEntity = JXSEC_TOP_FLANGE_LEFT Or _
                CrossSectionEntity = JXSEC_TOP_FLANGE_LEFT_BOTTOM_CORNER Or _
                CrossSectionEntity = JXSEC_TOP_FLANGE_RIGHT Or _
                CrossSectionEntity = JXSEC_TOP_FLANGE_RIGHT_BOTTOM_CORNER Or _
                CrossSectionEntity = JXSEC_BOTTOM_FLANGE_LEFT Or _
                CrossSectionEntity = JXSEC_BOTTOM_FLANGE_LEFT_TOP_CORNER Or _
                CrossSectionEntity = JXSEC_BOTTOM_FLANGE_RIGHT Or _
                CrossSectionEntity = JXSEC_BOTTOM_FLANGE_RIGHT_TOP_CORNER Then
                ProfilePartPortType = PROFILE_PORTTYPE_EDGE
            Else
                ProfilePartPortType = PROFILE_PORTTYPE_FACE
            End If
        
        Case "C_SS", "CSType"
            If CrossSectionEntity = JXSEC_TOP_FLANGE_RIGHT Or _
                CrossSectionEntity = JXSEC_TOP_FLANGE_RIGHT_BOTTOM_CORNER Or _
                CrossSectionEntity = JXSEC_BOTTOM_FLANGE_RIGHT Or _
                CrossSectionEntity = JXSEC_BOTTOM_FLANGE_RIGHT_TOP_CORNER Then
                ProfilePartPortType = PROFILE_PORTTYPE_EDGE
            Else
                ProfilePartPortType = PROFILE_PORTTYPE_FACE
            End If
        
        Case "BUTL3"
            If CrossSectionEntity = JXSEC_TOP Or _
                CrossSectionEntity = JXSEC_BOTTOM Or _
                CrossSectionEntity = JXSEC_TOP_FLANGE_RIGHT Then
                ProfilePartPortType = PROFILE_PORTTYPE_EDGE
            Else
                ProfilePartPortType = PROFILE_PORTTYPE_FACE
            End If
        
        Case "HalfR", "R", "P", "RT", "SB"
            ProfilePartPortType = PROFILE_PORTTYPE_FACE
        
        Case Else
            strError = "Unknown profile part cross section type (" & ProfilePartCrossSectionType & ")."
            GoTo ErrorHandler
    End Select

    Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, "ProfilePartPortType", strError).Number
End Function

Public Function PortBasicContext(PortTopologicalContext As IMSStructConnection.eUSER_CTX_FLAGS) As enmPortBasicContext

    On Error GoTo ErrorHandler
    
    Dim strError As String

    If (PortTopologicalContext = CTX_INVALID) Then
        strError = "PortTopologicalContext is Invalid."
        GoTo ErrorHandler
    ElseIf (PortTopologicalContext And CTX_BASE) Then
       PortBasicContext = PORT_BASIC_CONTEXT_BASE
        
    ElseIf (PortTopologicalContext And CTX_OFFSET) Then
       PortBasicContext = PORT_BASIC_CONTEXT_OFFSET
        
    ElseIf (PortTopologicalContext And CTX_LATERAL) Then
       PortBasicContext = PORT_BASIC_CONTEXT_LATERAL
        
    Else
        strError = "PortTopologicalContext doesn't correspond to Base, Offset, or Lateral."
        GoTo ErrorHandler
    End If

    Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, "PortBasicContext", strError).Number
End Function

'********************************************************************
' ' Routine: LogError
'
' Description:  default Error Reporter
'********************************************************************
Public Function LogError(oErrObject As ErrObject, _
                            Optional strSourceFile As String = "", _
                            Optional strMethod As String = "", _
                            Optional strExtraInfo As String = "") As IJError
     
    Dim strErrSource As String
    Dim strErrDesc As String
    Dim lErrNumber As Long
    Dim oEditErrors As IJEditErrors
     
    lErrNumber = oErrObject.Number
    strErrSource = oErrObject.Source
    strErrDesc = oErrObject.Description
     
     ' retrieve the error service
    Set oEditErrors = GetJContext().GetService("Errors")
       
    ' add the error to the service : the error is also logged to the file specified by
    ' "HKEY_LOCAL_MACHINE/SOFTWARE/Intergraph/Sp3D/Core/OperationParameter/ReportErrors_Log"
    Set LogError = oEditErrors.Add(lErrNumber, _
                                      strErrSource, _
                                      strErrDesc, _
                                      , _
                                      , _
                                      , _
                                      strMethod & ": " & strExtraInfo, _
                                      , _
                                      strSourceFile)
    Set oEditErrors = Nothing
End Function



Public Sub SetQuestionEndCondition(ByRef pQH As IJDQuestionsHelper)
    On Error GoTo ErrorHandler
     
    Dim strError As String
    Dim colAnswers As New Collection
    
'    colAnswers.Add "Fixed"
'    colAnswers.Add "Free"
'    colAnswers.Add "FlangeFree"
'    colAnswers.Add "Undefined"
'
'    strError = "Defining codelist."
'    pQH.DefineCodeList CL_ENDCONDITION, colAnswers
    
    strError = "Setting question."
    pQH.SetQuestion QUES_ENDCONDITION, gsFixed, CL_ENDCONDITION
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SetQuestionEndCondition", strError).Number
End Sub

Public Sub SetQuestionEndCutType(ByRef pQH As IJDQuestionsHelper)
    On Error GoTo ErrorHandler
     
    Dim strError As String
'    Dim colAnswers As New Collection
'
'    colAnswers.Add "W"
'    colAnswers.Add "C"
'    colAnswers.Add "F"
'    colAnswers.Add "S"
'    colAnswers.Add "FV"
''    colAnswers.Add "R"
''    colAnswers.Add "RV"
'
'    strError = "Defining codelist."
'    pQH.DefineCodeList CL_ENDCUTTYPE, colAnswers
    
    strError = "Setting question."
    pQH.SetQuestion QUES_ENDCUTTYPE, gsW, CL_ENDCUTTYPE
zMsgBox "   pQH.SetQuestion..." & QUES_ENDCUTTYPE & "..." & gsW
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SetQuestionEndCutType", strError).Number
End Sub

' =========================================================================
' =========================================================================
' =========================================================================
Public Sub zvMsgBox(vText As Variant, _
                   Optional sDumArg1 As String, Optional sTitle As String)
On Error Resume Next
Dim sText As String
    sText = vText
    zMsgBox sText, sDumArg1, sTitle

End Sub

' =========================================================================
' =========================================================================
' =========================================================================
Public Sub zMsgBox(sText As String, _
                   Optional sDumArg1 As String, Optional sTitle As String)
On Error Resume Next

Dim iFileNumber
Dim sFileName As String

Exit Sub
'$$$Debug $$$    MsgBox sText, , sTitle

    iFileNumber = FreeFile
    sFileName = "C:\Temp\TraceFile.txt"
    Open sFileName For Append Shared As #iFileNumber
    
    If Len(Trim(sTitle)) > 0 Then
        Write #iFileNumber, sTitle
    End If
    
    Write #iFileNumber, sText
    Close #iFileNumber
End Sub

' =========================================================================
' =========================================================================
' =========================================================================
Public Sub UpdateWebCutForFlangeCuts(pMemberDescription As IJDMemberDescription)
    On Error GoTo ErrorHandler
     
    Dim strError As String
    
    ' Force an Update on the WebCut using the same interface,IJStructGeometry,
    ' as is used when placing the WebCut as an input to the FlangeCuts
    ' This appears to allow Assoc to always recompute the WebCut before FlangeCuts
    Dim oSDO_WebCut As StructDetailObjects.WebCut
    Dim pMemberObjects As IJDMemberObjects
    
    strError = "Calling Structdetailobjects.WebCut::ForceUpdateForFlangeCuts"
    Set pMemberObjects = pMemberDescription.CAO
    Set oSDO_WebCut = New StructDetailObjects.WebCut
    Set oSDO_WebCut.object = pMemberObjects.Item(1)
    oSDO_WebCut.ForceUpdateForFlangeCuts
    Set oSDO_WebCut = Nothing
    Set pMemberObjects = Nothing
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "UpdateWebCutForFlangeCuts", strError).Number
End Sub

' =========================================================================
' =========================================================================
' =========================================================================
Public Function Get_CrossSectionType(ByRef pMD As IJDMemberDescription, _
                        Optional bConnectedObject1 As Boolean = True) As String
Const sMETHOD = "Get_CrossSectionType"
    On Error GoTo ErrorHandler

    Dim sError As String
    
    Get_CrossSectionType = ""
    
    ' Initialize wrapper class and get the bounded profile
    sError = "Setting Assembly Connection "
    Dim pAssyConn As StructDetailObjects.AssemblyConn
    Set pAssyConn = New StructDetailObjects.AssemblyConn
    Set pAssyConn.object = pMD.CAO
     
    ' Initialize wrapper class and get the 2 ports
    sError = "Getting Assembly bounded profile"
    Dim oBoundedProfile As IJConnectable
    If bConnectedObject1 Then
        Set oBoundedProfile = pAssyConn.ConnectedObject1
    Else
        Set oBoundedProfile = pAssyConn.ConnectedObject2
    End If
    
    Set pAssyConn = Nothing
    
    If TypeOf oBoundedProfile Is IJStiffener Then
        Dim oProfile As StructDetailObjects.ProfilePart
        Set oProfile = New StructDetailObjects.ProfilePart
        Set oProfile.object = oBoundedProfile
        
        Get_CrossSectionType = oProfile.SectionType
        
        Set oProfile = Nothing
        Set oBoundedProfile = Nothing
    Else
        Dim oBeam As StructDetailObjects.BeamPart
        Set oBeam = New StructDetailObjects.BeamPart
        Set oBeam.object = oBoundedProfile
        
        Get_CrossSectionType = oBeam.SectionType
        
        Set oBeam = Nothing
        Set oBoundedProfile = Nothing
    End If
        
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number
End Function

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'===============================================================================
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Public Sub zTest_KnuckleBoundPorts(oBoundedPort As Object, _
                                   oBoundingPort As Object, _
                                   oEndCutPosition As IJDPosition)
Const sMETHOD = "zTest_KnuckleBoundPorts"
    On Error GoTo ErrorHandler
    Dim sError As String
    
    Dim eContextid As Long
    Dim dDistance1 As Double
    Dim dDistance2 As Double
    
    Dim oPoint As IJDPosition
    Dim oVector As IJDVector
    
    Dim oPort As IJPort
    Dim oBasePort As IJPort
    Dim oOffsetPort As IJPort
    Dim oStructPort As IJStructPort
    Dim oConnectable As IJConnectable
    
    Dim oSDO_ProfilePort As StructDetailObjects.ProfilePart
    Dim oTopologyLocate As GSCADStructGeomUtilities.TopologyLocate
        
    If TypeOf oBoundedPort Is IJStructPort Then
        Set oStructPort = oBoundedPort
        eContextid = oStructPort.ContextID
        If eContextid And eUSER_CTX_FLAGS.CTX_LATERAL Then
            ' For Split Knuckle Cases (Miter or Boxed),
            ' the Lateral Port is not a valid Port
            ' switch this Port with a Base or Offset Port
            Set oPort = oBoundedPort
            Set oConnectable = oPort.Connectable
            
            Set oSDO_ProfilePort = New StructDetailObjects.ProfilePart
            Set oSDO_ProfilePort.object = oConnectable
            
            Set oBasePort = oSDO_ProfilePort.BasePort(BPT_Base)
            Set oOffsetPort = oSDO_ProfilePort.BasePort(BPT_Offset)
            
            Set oTopologyLocate = New GSCADStructGeomUtilities.TopologyLocate
            oTopologyLocate.GetProjectedPointOnModelBody oBasePort.Geometry, _
                                                         oEndCutPosition, _
                                                         oPoint, oVector
            dDistance1 = oEndCutPosition.DistPt(oPoint)
            
            oTopologyLocate.GetProjectedPointOnModelBody oOffsetPort.Geometry, _
                                                         oEndCutPosition, _
                                                         oPoint, oVector
            dDistance2 = oEndCutPosition.DistPt(oPoint)
            
            If dDistance1 < dDistance2 Then
                Set oBoundedPort = oSDO_ProfilePort.BasePortBeforeTrim(BPT_Base)
            Else
                Set oBoundedPort = oSDO_ProfilePort.BasePortBeforeTrim(BPT_Offset)
            End If
            
            Set oTopologyLocate = Nothing
            Set oSDO_ProfilePort = Nothing
            
        End If
    End If
    
    If TypeOf oBoundingPort Is IJStructPort Then
        Set oStructPort = oBoundingPort
        eContextid = oStructPort.ContextID
        If eContextid And eUSER_CTX_FLAGS.CTX_LATERAL Then
            ' For Split Knuckle Case (Miter or Boxed),
            ' the Lateral Port is not a valid Port
            ' switch this Port with a Base or Offset Port
            Set oPort = oBoundingPort
            Set oConnectable = oPort.Connectable
            
            Set oSDO_ProfilePort = New StructDetailObjects.ProfilePart
            Set oSDO_ProfilePort.object = oConnectable
            
            Set oBasePort = oSDO_ProfilePort.BasePort(BPT_Base)
            Set oOffsetPort = oSDO_ProfilePort.BasePort(BPT_Offset)
            
            Set oTopologyLocate = New GSCADStructGeomUtilities.TopologyLocate
            oTopologyLocate.GetProjectedPointOnModelBody oBasePort.Geometry, _
                                                         oEndCutPosition, _
                                                         oPoint, oVector
            dDistance1 = oEndCutPosition.DistPt(oPoint)
            
            oTopologyLocate.GetProjectedPointOnModelBody oOffsetPort.Geometry, _
                                                         oEndCutPosition, _
                                                         oPoint, oVector
            dDistance2 = oEndCutPosition.DistPt(oPoint)
            
            If dDistance1 < dDistance2 Then
                Set oBoundingPort = oSDO_ProfilePort.BasePortBeforeTrim(BPT_Base)
            Else
                Set oBoundingPort = oSDO_ProfilePort.BasePortBeforeTrim(BPT_Offset)
            End If
            
            Set oTopologyLocate = Nothing
            Set oSDO_ProfilePort = Nothing
            
        End If
    End If
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number
End Sub

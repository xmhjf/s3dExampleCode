Attribute VB_Name = "RulesCommon"
'*******************************************************************
'
'  Copyright (C) 2013 Intergraph Corporation. All rights reserved.
'
'  Author      : Alligators
'
'  History     :
'    8/12/2011 ----- TR-175607--------  GH   -----------------
'        5/Mar/2012------CR-208794 -------Alligators
'                                        SMAssyConRules force updating features out of the the range of chnage
'        3/Jul/2013   TR-235070 -------Alligators
'                                      IsEggCrateCondition method is corrected for Invert logical connection
'                                      based on 'SplitEffect' attribute.
'        23/Aug/2013   TR-229070 -------Alligators
'                                      Created ConstructPC_MemberAndPlateFace() method and added _
'                                      GetFacePortsOfMemberOverlappingWithPlate() from MemberByPlateSel to RulesCommon
'        5/Nov/2013 - vb/svsmylav
'            DI-CP-240506 changes are done in 'UpdateDependentCornerSeam' method for performance improvement.
'       8/Nov/2013 - vb/svsmylav
'            DI-CP-240506 Added new method 'IsACisWithinRangeOfChange' for seam-movement case.
'
'       27/Dec/2013 - GH
'            TI-CP-245577 Modified ConstructPC_MemberAndPlateFace() method to fill Port2 for PC Creation
'       12/April/2016  - dsmamidi    DM-292461 removed IsEggCrateCondition()
'               26/April/2016  - GHM             TR-219720 and TR-293343 Modified GetMiterPlaneForMutualBound() to handle all Miter cases
'                                                                        for both stiffeners and ERs with all mounting faces and all orientations
'                                                                        Created new method CheckIfBothPlanesAreSame()
'*********************************************************************************************
Option Explicit

Public Const m_sProjectName As String = CUSTOMERID + "AssyConRul"
Public Const m_sProjectPath As String = "S:\StructDetail\Data\SmartOccurrence\" + m_sProjectName + "\"

Public Const INPUT_BOUNDED_OR_PENETRATED_OBJECT = "ConnectedObject1"
Public Const INPUT_BOUNDING_OR_PENETRATING_OBJECT = "ConnectedObject2"

Public Const IID_IJPlate = "{53CF4EA0-91BF-11D1-BE56-080036B3A103}"
Public Const IID_IJStructureMaterial = "{E790A7C0-2DBA-11D2-96DC-0060974FF15B}"
Public Const IID_IJCollarPart = "{138C021D-7089-11D5-B0D9-006008676515}"

Public Const QUES_ENDCONDITION As String = "EndCondition"
Public Const CL_ENDCONDITION As String = "EndConditionCodeList"

Public Const QUES_ENDCUTTYPE As String = "EndCutType"
Public Const CL_ENDCUTTYPE As String = "EndCutTypeCodeList"

' Tolerance used to detrmine valid overlap between sheets
Private Const SHEET_OVERLAP_TOLERANCE As Double = 0.0045

' End cut EndCondition Constants
Public Const gsFixed = "Fixed"
Public Const gsFree = "Free"
Public Const gsFlangeFree = "FlangeFree"

' EndCut Type Constants
Public Const gsW = "Welded"
Public Const gsC = "C"
Public Const gsF = "F"
Public Const gsS = "S"
Public Const gsFV = "FV"
Public Const gsR = "R"
Public Const gsRV = "RV"

' Assembly Method Constants
Public Const gsDrop = "Drop"
Public Const gsSlide = "Slide"
Public Const gsDefaultValue = "Default value"


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

Public Const CMLIBRARY_ASSYCONNRULES As String = CUSTOMERID + "AssyConRul.AssyConnDefCM"
Public Const CMLIBRARY_ASSYCONNSEL As String = CUSTOMERID + "AssyConRul.AssyConnSelCM"
Public Const LIBRARY_SOURCE_ID = CUSTOMERID + "AssyConRul.PlateByPlateSel"

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\" + CUSTOMERID + "AssyConRul\RulesCommon.bas"

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
        
        Case "I", "ISType", "H"
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
''********************************************************************
'' ' Routine: LogError
''
'' Description:  default Error Reporter
''********************************************************************
'Public Function LogError(oErrObject As ErrObject, _
'                            Optional strSourceFile As String = "", _
'                            Optional strMethod As String = "", _
'                            Optional strExtraInfo As String = "") As IJError
'
'    Dim strErrSource As String
'    Dim strErrDesc As String
'    Dim lErrNumber As Long
'    Dim oEditErrors As IJEditErrors
'
'    lErrNumber = oErrObject.Number
'    strErrSource = oErrObject.Source
'    strErrDesc = oErrObject.Description
'
'     ' retrieve the error service
'    Set oEditErrors = GetJContext().GetService("Errors")
'
'    ' add the error to the service : the error is also logged to the file specified by
'    ' "HKEY_LOCAL_MACHINE/SOFTWARE/Intergraph/Sp3D/Core/OperationParameter/ReportErrors_Log"
'    Set LogError = oEditErrors.Add(lErrNumber, _
'                                      strErrSource, _
'                                      strErrDesc, _
'                                      , _
'                                      , _
'                                      , _
'                                      strMethod & ": " & strExtraInfo, _
'                                      , _
'                                      strSourceFile)
'    Set oEditErrors = Nothing
'End Function



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
'
'' =========================================================================
'' =========================================================================
'' =========================================================================
'Public Sub zMsgBox(sText As String, _
'                   Optional sDumArg1 As String, Optional sTitle As String)
'On Error Resume Next
'
'Dim iFileNumber
'Dim sFileName As String
'
'Exit Sub
''$$$Debug $$$    MsgBox sText, , sTitle
'
'    iFileNumber = FreeFile
'    sFileName = "C:\Temp\TraceFile.txt"
'    Open sFileName For Append Shared As #iFileNumber
'
'    If Len(Trim(sTitle)) > 0 Then
'        Write #iFileNumber, sTitle
'    End If
'
'    Write #iFileNumber, sText
'    Close #iFileNumber
'End Sub
'
'' =========================================================================
'' =========================================================================
'' =========================================================================
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
        
        Get_CrossSectionType = oProfile.sectionType
        
        Set oProfile = Nothing
        Set oBoundedProfile = Nothing
    Else
        Dim oBeam As StructDetailObjects.BeamPart
        Set oBeam = New StructDetailObjects.BeamPart
        Set oBeam.object = oBoundedProfile
        
        Get_CrossSectionType = oBeam.sectionType
        
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
            
            Set oBasePort = oSDO_ProfilePort.baseport(BPT_Base)
            Set oOffsetPort = oSDO_ProfilePort.baseport(BPT_Offset)
            
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
            
            Set oBasePort = oSDO_ProfilePort.baseport(BPT_Base)
            Set oOffsetPort = oSDO_ProfilePort.baseport(BPT_Offset)
            
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

Public Function IsLapped(ByRef oBoundedPort As IJPort, Optional ByRef dLapDist As Double) As Boolean

    IsLapped = False
    
    ' ------------------------------------
    ' Get the bounded object
    ' ------------------------------------
    Dim oBoundedConn As Object
    Set oBoundedConn = oBoundedPort.Connectable
    
    ' -----------------------------
    ' Get the root geometry
    ' -----------------------------
    Dim oParentSystem As Object
    Dim oStructDetailHelper As IJStructDetailHelper
    Set oStructDetailHelper = New GSCADStructDetailUtil.StructDetailHelper
    
    If TypeOf oBoundedConn Is IJStructGraph Then
        oStructDetailHelper.IsPartDerivedFromSystem oBoundedConn, oParentSystem, True
    End If
        
    If oParentSystem Is Nothing Then
        Set oParentSystem = oBoundedConn
    End If
    
    Dim oRootLC As Object
    Dim oProfileAttr As IJProfileAttributes
    Set oProfileAttr = New ProfileUtils
    oProfileAttr.GetRootLandingCurveFromProfile oParentSystem, oRootLC
    
    ' ---------------------------------
    ' Get the operation that created it
    ' ---------------------------------
    Dim oProfileGraphicRepr As IJStructGraphicRepresentation
    Set oProfileGraphicRepr = oRootLC
    
    Dim oLCActiveEntity As Object
    Dim oAEHelper As StructGenericUtilities.StructBOHelper
    Set oAEHelper = New StructBOHelper
    
    oAEHelper.GetCreateActiveEntity oProfileGraphicRepr, oLCActiveEntity

    ' ------------------------------------------
    ' If by elements, it is a tripping stiffener
    ' ------------------------------------------
    If TypeOf oLCActiveEntity Is IJLandCrvBetweenElements_AE Then
        
        ' ----------------------------------------
        ' Determine which end we are on
        ' ----------------------------------------
        Dim oBoundedStructPort As IJStructPort
        Set oBoundedStructPort = oBoundedPort
        
        Dim eCtx As eUSER_CTX_FLAGS
        eCtx = oBoundedStructPort.ContextID
        
        Dim bIsStart As Boolean
        bIsStart = False
        If eCtx And CTX_BASE Then
            bIsStart = True
        End If
        
        ' -------------------------------
        ' Determine if lapped at this end
        ' -------------------------------
        Dim oLCBE As IJLandCrvBetweenElements_AE
        Set oLCBE = oLCActiveEntity
        
        Dim endCond As LandingCrvAttachmentMethod
        endCond = oLCBE.AttachmentType(bIsStart)
    
        If endCond = LCA_LAPPED Then
            IsLapped = True
            
            Dim uOff As Double
            Dim vOff As Double
            uOff = Abs(oLCBE.AttachmentUOffset(bIsStart))
            vOff = Abs(oLCBE.AttachmentVOffset(bIsStart))
            
            ' Expect that one of these is zero and one is greater than zero
            If uOff > vOff Then
                dLapDist = uOff
            Else
                dLapDist = vOff
            End If
        End If
    End If
End Function

Public Sub CheckAndCreateBoundingPlane( _
                     ByVal oMD As IJDMemberDescription, _
                     ByVal oBoundedPort As Object, _
                     ByVal oBoundingPort As Object, _
                     ByRef oBoundingPlane As Object)
   
   Dim oCommonHelper As CommonHelper
   Dim sAlongGlobalAxis As String
   
   Set oCommonHelper = New CommonHelper
   Dim oSO As IJSmartOccurrence
   Dim oSI As IJSmartItem
   Dim oSC As IJSmartClass
   
   Set oSO = oMD.CAO
   Set oSI = oSO.ItemObject
   Set oSO = Nothing
   Set oSC = oSI.Parent
   Set oSI = Nothing
   
   On Error Resume Next
   sAlongGlobalAxis = oCommonHelper.GetAnswer( _
                                          oMD.CAO, _
                                          oSC.SelectionRuleDef, _
                                          "SplitEndToEndCase")
   Set oSC = Nothing
   Set oCommonHelper = Nothing

   If sAlongGlobalAxis = "AlongGlobalAxis" Then
      CreateBoundingPlaneAlongAxis _
                             oBoundedPort, _
                             oBoundingPort, _
                             oBoundingPlane
   End If
   
   Exit Sub
End Sub

' Create bounding plane along major axis
Public Sub CreateBoundingPlaneAlongAxis( _
                        ByVal oBoundedPort As Object, _
                        ByVal oBoundingObject As Object, _
                        ByRef oNewBoundingPlane As IJPlane, _
                        Optional ByRef oOldBoundingPlane As IJPlane = Nothing)
   Const METHOD = ":CreateBoundingPlaneAlongAxis"
   On Error GoTo ErrorHandler
   
   '
   ' Determine bounding plane root point and normal
   '
   Dim oPort As IJPort
   Dim oStructDetailHelper As New StructDetailHelper
   Dim oBoundedSystem As IJSystem
   
   Set oPort = oBoundedPort
   oStructDetailHelper.IsPartDerivedFromSystem _
                                   oPort.Connectable, _
                                   oBoundedSystem, _
                                   False
   
   Dim oSGOMBUtil As New SGOModelBodyUtilities
   Dim oBoundedSurfaceBody As IJSurfaceBody
   Dim oPointOnSystem As IJDPosition
   Dim oPointOnSurface As IJDPosition
   Dim dDistance As Double
   
   Set oBoundedSurfaceBody = oPort.Geometry
   oSGOMBUtil.GetClosestPointsBetweenTwoBodies oBoundedSystem, _
                                               oBoundedSurfaceBody, _
                                               oPointOnSystem, _
                                               oPointOnSurface, _
                                               dDistance
                                               
   Dim oBoundedPortNormal As IJDVector
   
   oBoundedSurfaceBody.GetNormalFromPosition _
                                   oPointOnSurface, _
                                   oBoundedPortNormal

   Dim dPortNormalX As Double
   Dim dPortNormalY As Double
   Dim dPortNormalZ As Double
   
   dPortNormalX = Abs(oBoundedPortNormal.x)
   dPortNormalY = Abs(oBoundedPortNormal.y)
   dPortNormalZ = Abs(oBoundedPortNormal.z)
   
   ' Select an Axis
   Dim ix As Long
   Dim iy As Long
   Dim iz As Long
   
   ' Determine the Global Major Axis
   If dPortNormalX > dPortNormalY Then
       If dPortNormalX > dPortNormalZ Then
           ' Major Axis is X
           ix = 8
           iy = 4
           iz = 0
       ElseIf dPortNormalZ > dPortNormalX Then
           ' Major Axis is Z
           ix = 0
           iy = 4
           iz = 8
       Else
           ix = -1
       End If
   ElseIf dPortNormalY > dPortNormalZ Then
       ' Major Axis is Y
           ix = 4
           iy = 8
           iz = 0
   ElseIf dPortNormalZ > dPortNormalX Then
       ' Major Axis is Z
           ix = 0
           iy = 4
           iz = 8
   Else
       ix = -1
   End If
   
   ' Create a new Bounding Plane base on the Global Major Axis
   If ix > -1 Then
       Dim oMatrix As AutoMath.DT4x4
       Dim oNewPlane As IJPlane

       Set oMatrix = New AutoMath.DT4x4
       oMatrix.LoadIdentity
       
       oMatrix.IndexValue(iz) = 0#
       oMatrix.IndexValue(iz + 1) = 0#
       oMatrix.IndexValue(iz + 2) = 1#
       
       oMatrix.IndexValue(iy) = 0#
       oMatrix.IndexValue(iy + 1) = 1#
       oMatrix.IndexValue(iy + 2) = 0#
       
       oMatrix.IndexValue(ix) = 1#
       oMatrix.IndexValue(ix + 1) = 0#
       oMatrix.IndexValue(ix + 2) = 0#
       
       oMatrix.IndexValue(12) = oPointOnSystem.x
       oMatrix.IndexValue(13) = oPointOnSystem.y
       oMatrix.IndexValue(14) = oPointOnSystem.z
       
       Dim oMajorAxisNormal As AutoMath.dVector
       
       Set oMajorAxisNormal = New AutoMath.dVector
       oMajorAxisNormal.Set oMatrix.IndexValue(8), oMatrix.IndexValue(9), oMatrix.IndexValue(10)
       If Not oBoundedPortNormal Is Nothing Then
           Dim dDot As Double
       
           dDot = oMajorAxisNormal.Dot(oBoundedPortNormal)
           If dDot > 0 Then
               oMatrix.IndexValue(8) = -oMatrix.IndexValue(8)
               oMatrix.IndexValue(9) = -oMatrix.IndexValue(9)
               oMatrix.IndexValue(10) = -oMatrix.IndexValue(10)
           End If
       End If
               
       ' Set the Bounding Object to be the Global Major Axis Plane
       If oOldBoundingPlane Is Nothing Then
           Dim oStructPlaneHelper As New StructPlane.StructPlaneHelper
           Dim oSDOHelper As New StructDetailObjects.Helper
           Dim oResourceManager As IUnknown

           Set oResourceManager = oSDOHelper.GetResourceManagerFromObject(oBoundedSystem)
           Set oNewPlane = oStructPlaneHelper.CreateStructPlane(oResourceManager)
       Else
           Set oNewPlane = oOldBoundingPlane
       End If
           
       ' Set the Bounding Object to be the Global Major Axis Plane
       oNewPlane.DefineByPointNormal _
                                 oMatrix.IndexValue(12), _
                                 oMatrix.IndexValue(13), _
                                 oMatrix.IndexValue(14), _
                                 oMatrix.IndexValue(8), _
                                 oMatrix.IndexValue(9), _
                                 oMatrix.IndexValue(10)

       oNewPlane.SetUDirection oMatrix.IndexValue(0), _
                              oMatrix.IndexValue(1), _
                              oMatrix.IndexValue(2)
                              
       Set oNewBoundingPlane = oNewPlane
       
   End If
   
   Exit Sub
   
ErrorHandler:
   Err.Raise LogError(Err, MODULE, METHOD).Number
End Sub

Public Sub ReplaceEndToEndCutBoundingObject( _
             ByVal oACMemberDescription As IJDMemberDescription)
    
   ' Get answer to AC's SplitEndToEndCase question
   Dim oCommonHelper As DefinitionHlprs.CommonHelper
   Dim sAlongGlobalAxis As Variant
   Dim oSO As IJSmartOccurrence
   Dim oSI As IJSmartItem
   Dim oSC As IJSmartClass
    
   Set oSO = oACMemberDescription.CAO
   Set oSI = oSO.SmartItemObject
   Set oSC = oSI.Parent
    
   Set oCommonHelper = New DefinitionHlprs.CommonHelper
   sAlongGlobalAxis = oCommonHelper.GetAnswer( _
                                     oSO, _
                                     oSC.SelectionRuleDef, _
                                     "SplitEndToEndCase")
   '
   ' In end to end case, assembly connection has following members:
   '
   '1 Part1 Web Cut
   '2 Part1 Top Flange Cut
   '3 Part1 Bottom Flange Cut
   '4 Part2 Web Cut
   '5 Part2 Top Flange Cut
   '6 Part2 Bottom Flange Cut
   '
   ' Applicable to both profile and beam

   Dim oMemberObjects As IJDMemberObjects
   Dim oExistingEndCut As Object
   Dim nMemberDispID As Long
   Dim bIsCutOnPart1 As Boolean
   Dim bIsWebCut As Boolean
   
   nMemberDispID = oACMemberDescription.dispid
   bIsCutOnPart1 = False
   Select Case nMemberDispID
      Case 1, 2, 3
         bIsCutOnPart1 = True
      Case 4, 5, 6
         bIsCutOnPart1 = False
   End Select
   
   If nMemberDispID = 1 Or nMemberDispID = 4 Then
      bIsWebCut = True
   Else
      bIsWebCut = False
   End If
   
   Dim oWebCutRelatedToFlangeCut As Object
   
   Set oMemberObjects = oACMemberDescription.CAO
   If bIsWebCut = True Then
      Set oWebCutRelatedToFlangeCut = Nothing
   Else
      If bIsCutOnPart1 = True Then
         Set oWebCutRelatedToFlangeCut = oMemberObjects.Item(1)
      Else
         Set oWebCutRelatedToFlangeCut = oMemberObjects.Item(4)
      End If
      
   End If
   
   Set oExistingEndCut = oMemberObjects.Item(nMemberDispID)
   
   If Not oExistingEndCut Is Nothing Then
      ReplaceEndToEndCutBoundingObjectInternal _
               oExistingEndCut, _
               sAlongGlobalAxis, _
               bIsWebCut, _
               bIsCutOnPart1, _
               oWebCutRelatedToFlangeCut
   End If
   
   Exit Sub
   
End Sub

Public Sub ReplaceEndToEndCutBoundingObjectInternal( _
             ByVal oEndCut As Object, _
             ByVal sAlongGlobalAxis As String, _
             ByVal bIsWebCut As Boolean, _
             ByVal bIsCutOnPart1 As Boolean, _
             ByVal oWebCut As Object)
   Dim oDesignChild As IJDesignChild
   Dim oEndCutParent As Object
    
   Set oDesignChild = oEndCut
   Set oEndCutParent = oDesignChild.GetParent
    
   Dim oBoundingObject As Object
   Dim oBoundedObject As Object

   If bIsWebCut = True Then
      Dim oWebCutWrapper As New StructDetailObjects.WebCut
   
      Set oWebCutWrapper.object = oEndCut
      Set oBoundingObject = oWebCutWrapper.BoundingPort
      Set oBoundedObject = oWebCutWrapper.BoundedPort
   Else
      Dim oFlangeCutWrapper As New StructDetailObjects.FlangeCut
      
      Set oFlangeCutWrapper.object = oEndCut
      Set oBoundingObject = oFlangeCutWrapper.BoundingPort
      Set oBoundedObject = oFlangeCutWrapper.BoundedPort
   End If

   Dim bReplaceBounding As Boolean
   Dim oNewBoundingObject As Object
   
   bReplaceBounding = False
   If TypeOf oBoundingObject Is IJPort Then
      If sAlongGlobalAxis = "AlongGlobalAxis" Then
         bReplaceBounding = True
         
         ' Replace port with plane
         CreateBoundingPlaneAlongAxis _
                                 oBoundedObject, _
                                 oBoundingObject, _
                                 oNewBoundingObject
      End If
   ElseIf TypeOf oBoundingObject Is IJPlane Then
      If LCase(sAlongGlobalAxis) <> LCase("AlongGlobalAxis") Then
         bReplaceBounding = True
         
         ' Replace plane with port from assembly connection
         Dim oACWrapper As New StructDetailObjects.AssemblyConn
          
         Set oACWrapper.object = oEndCutParent
         If bIsCutOnPart1 = True Then
            Set oNewBoundingObject = oACWrapper.Port2 ' Ref end cut creation CM
         Else
            Set oNewBoundingObject = oACWrapper.Port1 ' Ref end cut creation CM
         End If
      End If
  End If

  If bReplaceBounding = True Then
     Dim oSO As IJSmartOccurrence
     Dim oDObject As IJDObject
     Dim oSDFeatureUtil As New SDFeatureUtils
    
     Set oSO = oEndCut
     Set oDObject = oEndCut

     If bIsWebCut = True Then
        oSDFeatureUtil.CreateWebCut _
                           oDObject.ResourceManager, _
                           oNewBoundingObject, _
                           oBoundedObject, _
                           oSO.RootSelection, _
                           oEndCutParent, _
                           oEndCut
     Else
        oSDFeatureUtil.CreateFlangeCut _
                           oDObject.ResourceManager, _
                           oNewBoundingObject, _
                           oBoundedObject, _
                           oWebCut, _
                           oSO.RootSelection, _
                           oEndCutParent, _
                           oEndCut
     End If
  End If
   
  Exit Sub
   
End Sub

Public Sub IsLengthEdgeOfBuiltup(oPort As IJPort, oBuiltupMember As ISPSDesignedMember, mBoolean As Boolean)

        Dim oMemberPart As New StructDetailObjects.MemberPart
        Dim oParent As IJDesignParent
        Dim ppChildren As IJDObjectCollection
        Dim oIndex As Integer
        Dim oCOunt As Integer
        Dim oObject As Object
        Dim oModelBody1 As IJModelBody
        Dim oTopologyLocate As GSCADStructGeomUtilities.TopologyLocate

        Dim x As Double
        Dim y As Double
        Dim z As Double
    
        Dim dDotProduct As Double
    
        Dim oPlane1 As IJPlane
        Dim oPlate As IJPlate
        
        Dim oPortNormal As IJDVector
        Dim bApproxUsed As Boolean
        Dim oPartInfo As GSCADStructGeomUtilities.PartInfo
        Dim oWebNormal As IJDVector
    
        Set oParent = oBuiltupMember
        oParent.GetChildren ppChildren
        oCOunt = ppChildren.Count
        Set oTopologyLocate = New GSCADStructGeomUtilities.TopologyLocate
        For Each oObject In ppChildren
            Set oPlate = oObject
            If oPlate.plateType = WebPlate Then
'             Set oPlatePart = oObject
             Set oModelBody1 = oTopologyLocate.GetPlateParentBodyModel(oPlate)
                If TypeOf oModelBody1 Is IJPlane Then
                    Set oPlane1 = oModelBody1
                    oPlane1.GetNormal x, y, z
                End If
            End If
        Next oObject
        Set oWebNormal = New dVector
        oWebNormal.x = x
        oWebNormal.y = y
        oWebNormal.z = z
        oWebNormal.Length = 1

        Set oPartInfo = New GSCADStructGeomUtilities.PartInfo
        Set oPortNormal = oPartInfo.GetPortNormal(oPort, bApproxUsed)
        oPortNormal.Length = 1

        dDotProduct = oWebNormal.Dot(oPortNormal)
        If dDotProduct = 1 Or dDotProduct = -1 Then
            mBoolean = True
        End If
        
End Sub
'***********************************************************************
' METHOD:  IsPortFromBuiltUpMember
'
' DESCRIPTION:  Compare the normal of the molded surface of two
'               connected plate parts. If they point to the same
'               direction, then return true. If not, return false.
'***********************************************************************
Public Sub IsPortFromBuiltUpMember(oPort As IJPort, _
                                   bFromBuiltUp As Boolean, _
                                   Optional oBuiltupMember As ISPSDesignedMember)
Const METHOD = "::IsPortFromBuiltUpMember"
On Error GoTo ErrorHandler
    
    bFromBuiltUp = False
    
    Dim oParentObject As Object
    Dim oPlateSystem As IJPlateSystem
    Dim oSDO_PlatePart As StructDetailObjects.PlatePart
    Dim oSDO_PlateSystem As StructDetailObjects.PlateSystem
    
    ' Check if given port is from a PlatePart
    If TypeOf oPort.Connectable Is IJPlatePart Then
        ' Given Port's Connectable is IJplatePart
        ' Get the Plate Part's Parent Object
        Set oSDO_PlatePart = New StructDetailObjects.PlatePart
        Set oSDO_PlatePart.object = oPort.Connectable
        Set oParentObject = oSDO_PlatePart.ParentSystem
        Set oSDO_PlatePart = Nothing
            
        ' Check if the Plate Part's Parent object is IJPlateSystem
        If TypeOf oParentObject Is IJPlateSystem Then
            ' Plate Part's Parent object is IJPlateSystem
            ' Get the Plate Systems's Parent object
            Set oSDO_PlateSystem = New StructDetailObjects.PlateSystem
            Set oSDO_PlateSystem.object = oParentObject
            Set oParentObject = oSDO_PlateSystem.ParentSystem
            Set oSDO_PlateSystem = Nothing
            
            ' Check if the Plate System's Parent object is IJPlateSystem
            If TypeOf oParentObject Is IJPlateSystem Then
                Set oPlateSystem = oParentObject
                bFromBuiltUp = oPlateSystem.IsBuiltupPlateSystem
                
                If bFromBuiltUp Then
                    Set oBuiltupMember = oPlateSystem.ParentBuiltup
                End If
            End If
            
        End If
    End If
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
End Sub


Public Function IsCanInvolved(oPort1 As IJPort, oPort2 As IJPort, PrimaryMemberSystem As ISPSMemberSystem) As Boolean

Const METHOD = "::IsCanByCanCase"
On Error GoTo ErrorHandler
            
            Dim bBuiltUp1 As Boolean
            Dim bBuiltUp2 As Boolean
            Dim oBuiltupMember1 As ISPSDesignedMember
            Dim oBuiltupMember2 As ISPSDesignedMember
            
            Dim oCan1 As New StructDetailObjects.MemberPart
            Dim oCan2 As New StructDetailObjects.MemberPart
            Dim oCanRule As ISPSCanRule
            Dim oSpsCanRuleStatus As SPSCanRuleStatus
            Dim pLine As IJLine
            Dim oMemberLine As IJLine
            
            IsPortFromBuiltUpMember oPort1, bBuiltUp1, oBuiltupMember1
            IsPortFromBuiltUpMember oPort2, bBuiltUp2, oBuiltupMember2
            
            If bBuiltUp1 And bBuiltUp2 Then 'introduce builtup2 and AND
                Set oCan1.object = oBuiltupMember1
                Set oCan2.object = oBuiltupMember2
              'check needed to ensure it is a can type.
                Dim oSectionName1 As String
                Dim oSectionName2 As String
                oSectionName1 = oCan1.sectionType
                oSectionName2 = oCan2.sectionType
                If oSectionName1 = "BUCan" Or oSectionName2 = "BUCan" Then
                    IsCanInvolved = True
                End If
                'get the normal of the built up parent
                If oSectionName1 = "BUCan" Then
                    Set oCanRule = oCan1.CanRule
                    oSpsCanRuleStatus = oCanRule.GetPrimaryMemberSystem(PrimaryMemberSystem)
                ElseIf oSectionName2 = "BUCan" Then
                    Set oCanRule = oCan2.CanRule
                    oSpsCanRuleStatus = oCanRule.GetPrimaryMemberSystem(PrimaryMemberSystem)
                End If
           End If
           Set oCan1 = Nothing
           Set oCan2 = Nothing
           Set oBuiltupMember1 = Nothing
           Set oBuiltupMember2 = Nothing
       Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
End Function
Function Position_FromLine(pLine As IJLine, iIndex As Integer) As IJDPosition
    ' create new position
    Dim pPosition As IJDPosition
    Set pPosition = New DPosition
    
    ' retrieve coordinates
    Dim x As Double, y As Double, z As Double
    If iIndex = 0 Then
        Call pLine.GetStartPoint(x, y, z)
    Else
        Call pLine.GetEndPoint(x, y, z)
    End If
    
    ' set coordinates
    Call pPosition.Set(x, y, z)
    
    ' return result
    Set Position_FromLine = pPosition
End Function
Public Function Line_FromPositions(pPOM As IJDPOM, pPositionOfStartPoint As IJDPosition, pPositionOfEndPoint As IJDPosition) As IJLine
    ' create new line
    Dim pGeometryFactory As New GeometryFactory
    Set Line_FromPositions = pGeometryFactory.Lines3d.CreateBy2Points(pPOM, _
        pPositionOfStartPoint.x, pPositionOfStartPoint.y, pPositionOfStartPoint.z, _
        pPositionOfEndPoint.x, pPositionOfEndPoint.y, pPositionOfEndPoint.z)
End Function
Function Vector_FromLine(pLine As IJLine) As IJDVector
    ' get vector
    Dim u As Double, v As Double, w As Double
    Call pLine.GetDirection(u, v, w)
    
    ' create result
    Dim pVector As New dVector
    Call pVector.Set(u, v, w)
    
    ' return result
    Set Vector_FromLine = pVector
End Function

' Thickness difference along tube normal is calculated.

'the thickness of the plate inclined is adjusted.
'In cases where CANS are involved,it is necessary to determine if the chamfer is on
'cone portion of CAN or cylinder portion of CAN.
'the tube normals of the plates are taken and compared to the normal of the PrimaryMemberSystem
'along which the CAN is placed.

'code is written to make sure that the thickness of the cone part is always adjusted.

Public Sub GetThicknessAlongTubeNormal(oChamferPart1 As StructDetailObjects.PlatePart, oChamferPart2 As StructDetailObjects.PlatePart, PrimaryMemberSystem As ISPSMemberSystem, oAngle As Double, dThicknessDiff_TN As Double)
             
            Dim oLineDir1 As IJDVector
            Dim pLine As IJLine
            Dim oMemberLine As IJLine
            
            Dim pPOM As IJDPOM
            Set pPOM = Nothing
            
            Set pLine = PrimaryMemberSystem.LogicalAxis.CurveGeometry
            Set oMemberLine = Line_FromPositions(pPOM, Position_FromLine(pLine, 0), Position_FromLine(pLine, 1))
            'This gives the Direction of the PrimaryMemberSystem along which the CAN is placed.
            Set oLineDir1 = Vector_FromLine(oMemberLine)
                   
            Dim oTopologyLocate As GSCADStructGeomUtilities.TopologyLocate
            Set oTopologyLocate = New GSCADStructGeomUtilities.TopologyLocate

            Dim oModelBody1 As IJModelBody
            Dim oModelBody2 As IJModelBody
           
            Dim oTopoIntersect As IJDTopologyIntersect
            Set oTopoIntersect = New DGeomOpsIntersect
            
            Dim oStPos As IJDPosition
            Dim oEndPos As IJDPosition
            Dim oCommonIntersection As Object
            Dim oWireBody As IJWireBody
           
            Set oModelBody1 = oTopologyLocate.GetPlateParentBodyModel(oChamferPart1.ParentSystem)
            Set oModelBody2 = oTopologyLocate.GetPlateParentBodyModel(oChamferPart2.ParentSystem)
            
            Dim oSurfaceBody1 As IJSurfaceBody
            Dim oSurfaceBody2 As IJSurfaceBody

            Dim ppNormal1 As IJDVector
            Dim ppNormal2 As IJDVector
           
            Dim oShipGeomOps As GSCADShipGeomOps.SGOModelBodyUtilities
            Set oShipGeomOps = New GSCADShipGeomOps.SGOModelBodyUtilities
            
            Dim ppPointOnFirstBody As IJDPosition
            Dim ppPointOnSecondBody As IJDPosition
            Dim ppdistance As Double
            
            oShipGeomOps.GetClosestPointsBetweenTwoBodies oModelBody1, oModelBody2, ppPointOnFirstBody, ppPointOnSecondBody, ppdistance
            
            If TypeOf oModelBody1 Is IJSurfaceBody Then
                Set oSurfaceBody1 = oModelBody1
                oSurfaceBody1.GetNormalFromPosition ppPointOnFirstBody, ppNormal1
            End If

            If TypeOf oModelBody2 Is IJSurfaceBody Then
                Set oSurfaceBody2 = oModelBody2
                oSurfaceBody2.GetNormalFromPosition ppPointOnSecondBody, ppNormal2
            End If
            
            ppNormal1.Length = 1
            ppNormal2.Length = 1
                      
           If Not oLineDir1 Is Nothing Then
                If Not Round(oLineDir1.Dot(ppNormal1), 3) = 0 Then
                   dThicknessDiff_TN = oChamferPart2.PlateThickness - (oChamferPart1.PlateThickness / Cos(oAngle))
                ElseIf Not Round(oLineDir1.Dot(ppNormal2), 3) = 0 Then
                    dThicknessDiff_TN = (oChamferPart2.PlateThickness / Cos(oAngle)) - oChamferPart1.PlateThickness
                ElseIf oChamferPart2.PlateThickness > oChamferPart1.PlateThickness Then
                    dThicknessDiff_TN = oChamferPart2.PlateThickness - oChamferPart1.PlateThickness
                ElseIf oChamferPart1.PlateThickness > oChamferPart2.PlateThickness Then
                    dThicknessDiff_TN = oChamferPart1.PlateThickness - oChamferPart2.PlateThickness
                End If
           End If
          
End Sub

'***********************************************************************
' METHOD:  GetPortsForBracketInAC
'
' DESCRIPTION:  It gets the ports of the Tripping Bracket for the
'               AC in which we are
'
' oBracketPort is Port of the Bracket
' oSupportPort is Port of the Support to which AC is connected
' oBracketPart is the Bracket object itself
' oSupportRoot is Support object itself
'***********************************************************************
 Public Sub GetPortsForBracketInAC(ByRef pMD As IJDMemberDescription, _
                                  ByRef oBracketPort As IJPort, _
                                  ByRef oSupportPort As IJPort, _
                                  ByRef oBracketPart As Object, _
                                  ByRef oSupportRoot As Object)

    On Error GoTo ErrorHandler
'    sMETHOD = "GetPortsForBracketInAC"
 
    Dim oSDO_AC As New StructDetailObjects.AssemblyConn
    Set oSDO_AC.object = pMD.CAO
 
    Dim oPart1 As Object
    Dim oPart2 As Object
    Set oPart1 = oSDO_AC.ConnectedObject1
    Set oPart2 = oSDO_AC.ConnectedObject2
    
    Set oBracketPort = Nothing
    Set oSupportPort = Nothing
    Set oBracketPart = Nothing
    Set oSupportRoot = Nothing
    
    If Not IsBracket(oPart1) And Not IsBracket(oPart2) Then
        Exit Sub
    End If
    
    ' Find the bounded bracket (both parts could be brackets, if one is a boundary for the other)
    If IsBracket(oPart1) And Not IsBracket(oPart2) Then
        Set oBracketPort = oSDO_AC.Port1
        Set oSupportPort = oSDO_AC.Port2
    ElseIf Not IsBracket(oPart1) And IsBracket(oPart2) Then
        Set oBracketPort = oSDO_AC.Port2
        Set oSupportPort = oSDO_AC.Port1
    Else
        Dim oPort1 As IJStructPort
        Dim oPort2 As IJStructPort
        Set oPort1 = oSDO_AC.Port1
        Set oPort2 = oSDO_AC.Port2
        
        If (oPort1.ContextID And CTX_LATERAL) And Not (oPort2.ContextID And CTX_LATERAL) Then
            Set oBracketPort = oSDO_AC.Port1
            Set oSupportPort = oSDO_AC.Port2
        ElseIf Not (oPort1.ContextID And CTX_LATERAL) And (oPort2.ContextID And CTX_LATERAL) Then
            Set oBracketPort = oSDO_AC.Port2
            Set oSupportPort = oSDO_AC.Port1
        Else
            ' this must be a seam in the bracket.  We're not interested in these.
            Exit Sub
        End If
    End If
    
    Set oBracketPart = oBracketPort.Connectable
        
    Dim oStructDetailHelper As StructDetailHelper
    Set oStructDetailHelper = New StructDetailHelper
    oStructDetailHelper.IsPartDerivedFromSystem oSupportPort.Connectable, oSupportRoot, True
    
    If oSupportRoot Is Nothing Then
        Set oSupportRoot = oSupportPort.Connectable
    End If
    
    Exit Sub
    
ErrorHandler:
'    Err.Raise LogError(Err, MODULE, sMETHOD).Number
End Sub



' *********************************************************************************
'
' Returns root object of each support, whether a part, port, or system
' Callers of this method will compare these supports to the objects that created an AC port
' Those objects will be the root objects
' Presuming here that tripping brackets supports are already root objects
' If that is not true, this code will need some modification
'
 ' When Bracket is passed in this method gets its all the supports and corresponding
   ' Bracket ports
   '   oS1 is the Support 1 of the Bracket
   '   oS2 is the Support 2 of the Bracket
   '   oS3 is the Support 3 of the Bracket(If required...in case of 3S Tripping Bracket)
   '   oS4 is the Support 4 of the Bracket(If required...in case of Bracket by Plane)
   '   oS5 is the Support 5 of the Bracket(If required...in case of Bracket by Plane)
   '   oS1Port is the port of the bracket connected to support 1
   '   oS2Port is the port of the bracket connected to support 2
   '   oS3Port is the port of the bracket connected to support 3
   '   oS4Port is the port of the bracket connected to support 4
   '   oS5Port is the port of the bracket connected to support 5
'***********************************************************************************
Public Sub GetBracketSupports(oBracket As Object, _
                              oS1 As Object, _
                              oS2 As Object, _
                              oS3 As Object, _
                              oS4 As Object, _
                              oS5 As Object, _
                              oS1Port As IJPort, _
                              oS2Port As IJPort, _
                              oS3Port As IJPort, _
                              oS4Port As IJPort, _
                              oS5Port As IJPort)

    On Error GoTo ErrorHandler
'    sMETHOD = "GetBracketSupports"

    ' ----------------------------------------
    ' If the object is not a plate, return now
    ' ----------------------------------------
    If Not TypeOf oBracket Is IJPlate Then
        Exit Sub
    End If
        
    ' --------------------------------
    ' If a part, see if system-derived
    ' --------------------------------
    Dim oStructDetailHelper As StructDetailHelper
    Set oStructDetailHelper = New StructDetailHelper
    
    Dim oRootBracket As Object ' either standalone or system
    If TypeOf oBracket Is IJPlatePart Then
        oStructDetailHelper.IsPartDerivedFromSystem oBracket, oRootBracket, True

        If oRootBracket Is Nothing Then
            Set oRootBracket = oBracket
        End If
    End If
    
    ' ----------------------------------------------------------------------
    ' If root object is a bracket-type plate part, it is a standalone object
    ' ----------------------------------------------------------------------
    Dim oPlateUtil As IJPlateAttributes
    Set oPlateUtil = New PlateUtils
    
    Dim plateType As StructPlateType
    plateType = CollarPlate
    If TypeOf oBracket Is IJPlate Then
        Dim oPlate As IJPlate
        Set oPlate = oBracket
        plateType = oPlate.plateType
    End If
        
    If TypeOf oRootBracket Is IJPlatePart And plateType = BracketPlate Then
        Dim oSDOBracket As New StructDetailObjects.Bracket
        Set oSDOBracket.object = oBracket

        Dim oPort1 As IJPort
        Dim oPort2 As IJPort
        Dim nSupports As Long
        Dim oRefPlane As IJPlane

        oSDOBracket.GetInputs nSupports, oRefPlane, oPort1, oPort2, oS3, oS4, oS5
             
       ' Return the root object of the support
        Dim oRootParent As Object

        Set oS1 = oPort1.Connectable
        oStructDetailHelper.IsPartDerivedFromSystem oS1, oRootParent, True
        If Not oRootParent Is Nothing Then
            Set oS1 = oRootParent
            Set oRootParent = Nothing
        End If
        
        Set oS2 = oPort2.Connectable
        oStructDetailHelper.IsPartDerivedFromSystem oS2, oRootParent, True
        If Not oRootParent Is Nothing Then
            Set oS2 = oRootParent
            Set oRootParent = Nothing
        End If
        
        Dim oPort As IJPort
        If Not oS3 Is Nothing Then
            If TypeOf oS3 Is IJPort Then
                Set oPort = oS3
                Set oS3 = oPort.Connectable
            End If

            oStructDetailHelper.IsPartDerivedFromSystem oS3, oRootParent, True
            If Not oRootParent Is Nothing Then
                Set oS3 = oRootParent
                Set oRootParent = Nothing
            End If
        End If

        If Not oS4 Is Nothing Then
            If TypeOf oS4 Is IJPort Then
                Set oPort = oS4
                Set oS4 = oPort.Connectable
            End If

            oStructDetailHelper.IsPartDerivedFromSystem oS4, oRootParent, True
            If Not oRootParent Is Nothing Then
                Set oS4 = oRootParent
                Set oRootParent = Nothing
            End If
        End If

        If Not oS5 Is Nothing Then
            If TypeOf oS5 Is IJPort Then
                Set oPort = oS5
                Set oS5 = oPort.Connectable
            End If

            oStructDetailHelper.IsPartDerivedFromSystem oS5, oRootParent, True
            If Not oRootParent Is Nothing Then
                Set oS5 = oRootParent
                Set oRootParent = Nothing
            End If
        End If
        
    ' --------------------------------------
    ' Otherwise, check if a tripping bracket
    ' --------------------------------------
    ElseIf oPlateUtil.IsTrippingBracket(oRootBracket) Then
        Dim oBracketU As IJDVector
        Dim oBracketV As IJDVector
        Dim oAE As IJPlaneByElements_AE

        oPlateUtil.GetInput_TrippingBracket oRootBracket, oBracketU, oBracketV, oS1, oS2, oS3, oAE
    ' ------------------------------------
    ' Finally, check if a bracket-by-plane
    ' ------------------------------------
    ElseIf oPlateUtil.IsBracketByPlane(oRootBracket) Then
        Dim oPlane As IJPlane
        Dim oBracketUPt As IJPoint
        Dim oBracketVPt As IJPoint
        Dim strRootSel As String
        Dim oSupports As IJElements
        
        oPlateUtil.GetInput_BracketByPlane oRootBracket, oPlane, oBracketUPt, oBracketVPt, strRootSel, oSupports
        
        Dim i As Long
        For i = 1 To oSupports.Count
            Select Case i
                Case 1
                    Set oS1 = oSupports.Item(i)
                Case 2
                    Set oS2 = oSupports.Item(i)
                Case 3
                    Set oS3 = oSupports.Item(i)
                Case 4
                    Set oS4 = oSupports.Item(i)
                Case 5
                    Set oS5 = oSupports.Item(i)
            End Select
        Next i
    End If
    
    ' --------------------------------------------
    ' Find part ports associated with each support
    ' --------------------------------------------
    Dim ePortType As JS_TOPOLOGY_FILTER_TYPE
    ePortType = JS_TOPOLOGY_FILTER_SOLID_LATERAL_LFACES
    
    Dim GeomUtils As IJTopologyLocate
    Set GeomUtils = New TopologyLocate
     
    Dim oListOfPorts As IEnumUnknown
    GeomUtils.GetNamedPorts oBracket, ePortType, oListOfPorts
    
    Dim oCollectionOfPorts As Collection
    Dim ConvertUtils As New CCollectionConversions
    ConvertUtils.CreateVBCollectionFromIEnumUnknown oListOfPorts, oCollectionOfPorts
    
    Dim oOperation As Object
    Dim oOperator As Object
    
    Dim oStructPort As IJStructPort
    Dim oOperatorPort As IJPort
    
    For Each oPort In oCollectionOfPorts
        
        
        Set oStructPort = oPort
        oStructDetailHelper.FindOperatorForOperationInGraphByID oBracket, _
                                                                oStructPort.OperationID, _
                                                                oStructPort.OperatorID, _
                                                                oOperation, _
                                                                oOperator
        If Not oOperator Is Nothing Then
            If TypeOf oOperator Is IJPort Then
                Set oOperatorPort = oOperator
                Set oOperator = oOperatorPort.Connectable
            End If
                
            Dim oStiffener As IJStiffener
            
            If IsSupportOrPlateForStiffenerSupport(oOperator, oS1) Then
                Set oS1Port = oPort
            ElseIf IsSupportOrPlateForStiffenerSupport(oOperator, oS2) Then
                Set oS2Port = oPort
            ElseIf IsSupportOrPlateForStiffenerSupport(oOperator, oS3) Then
                Set oS3Port = oPort
            ElseIf IsSupportOrPlateForStiffenerSupport(oOperator, oS4) Then
                Set oS4Port = oPort
            ElseIf IsSupportOrPlateForStiffenerSupport(oOperator, oS5) Then
                Set oS5Port = oPort
            End If
        Else
            ' FindOperatorForOperationInGraphByID, but doesn't work for standalone brackets, but it appears
            ' (dangerous assumption?) that we can count on *all* the ACs being there when we evaluate
            ' each one... or I've been lucky.  Needs further testing.  What we really want is a way to ask
            ' for the operator for the port, which wouldn't rely on the AC.
            If IsConnectedToObjectOrChildren(oPort, oS1) Then
                Set oS1Port = oPort
            ElseIf IsConnectedToObjectOrChildren(oPort, oS2) Then
                Set oS2Port = oPort
            ElseIf IsConnectedToObjectOrChildren(oPort, oS3) Then
                Set oS3Port = oPort
            ElseIf IsConnectedToObjectOrChildren(oPort, oS4) Then
                Set oS4Port = oPort
            ElseIf IsConnectedToObjectOrChildren(oPort, oS5) Then
                Set oS5Port = oPort
            End If
        End If
        
    Next oPort
    
    Exit Sub
    
ErrorHandler:
'    Err.Raise LogError(Err, MODULE, sMETHOD).Number
    
End Sub


Public Function IsSupportOrPlateForStiffenerSupport(oObject As Object, oSupport As Object) As Boolean

IsSupportOrPlateForStiffenerSupport = False

If oSupport Is Nothing Then
    Exit Function
End If

    'The type of oSupport is plate or stiffener object
    If IsObjectAndSupportConnected(oObject, oSupport) = False Then
        If oObject Is oSupport Then
    IsSupportOrPlateForStiffenerSupport = True

        ElseIf TypeOf oSupport Is IJStiffener Then
    Dim oSupportStiffener As IJStiffener
    Set oSupportStiffener = oSupport
    If oSupportStiffener.PlateSystem Is oObject Then
        IsSupportOrPlateForStiffenerSupport = True
                
            End If
        End If
    Else 'The type of oSupport IJPort
        IsSupportOrPlateForStiffenerSupport = True
    End If

End Function

'*************************************************************************
'Function
'   IsObjectAndSupportConnected
'
'Purpose:
'   Judge whether Object and Support is connected
'Inputs :
'   Two objects
'Returns:
'   True
'   False
'*********************************************************************************************
Public Function IsObjectAndSupportConnected(oObject As Object, oSupport As Object) As Boolean
    Const sMETHOD = "IsObjectAndSupportConnected"
    On Error GoTo ErrorHandler

    IsObjectAndSupportConnected = False

    If oSupport Is Nothing Then
        Exit Function
    End If

    Dim oHelper As New StructDetailHelper
    
    Dim oSupportPort As IJPort
    Dim oSupportObject As Object
    Dim zConnectionData() As ConnectionData
    Dim nConnected As Long
    
    Dim nBktConnected As Long
    Dim zBktConnectionData() As ConnectionData

    If TypeOf oSupport Is IJPort Then
    
        Set oSupportPort = oSupport
        Set oSupportObject = oSupportPort.Connectable
        
        If oSupportObject Is oObject Then  'The support and the object are the same one object
            IsObjectAndSupportConnected = True
            
        'The support and the object are two diffrent object,when they are connected reture true,or reture false
        Else
            nConnected = GetAllConnectables(oSupportObject, AppConnectionType_Unknown, zConnectionData)
            
            Dim iIndex As Long
            'Compare all the objects connected with oSupportObject to oSupport
            For iIndex = 1 To nConnected
            
            Dim oConnectData As ConnectionData
            oConnectData = zConnectionData(iIndex)
            
            'Two objects is connected
            If oConnectData.ConnectingPort Is oSupport Then
                Dim oOperatorSys As Object
                Dim oSConnectSys As Object
                
                On Error Resume Next
                oHelper.IsPartDerivedFromSystem oConnectData.ToConnectable, oSConnectSys, True
                oHelper.IsPartDerivedFromSystem oObject, oOperatorSys, True
                Err.Clear
                
                'Make sure oSConnectSys and oOperatorSys are all at the same level(Object)
                If oSConnectSys Is Nothing Then
                    Set oSConnectSys = oConnectData.ToConnectable
                End If
                
                If oOperatorSys Is Nothing Then
                    Set oOperatorSys = oObject
                End If
                
                If oSConnectSys Is oOperatorSys Then
                    IsObjectAndSupportConnected = True
                End If
    End If
            
            Next
        End If
    Else
        IsObjectAndSupportConnected = False
    End If
            

    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD).Number

End Function

'********************************************************************
' Routine: GetAllConnectables
'
' Description:
'   Given a Connectable (Parts)
'   return a list of All Connectables (Parts) that are connected
'
' Inputs:
'       Connectable1 -  Connectable Object,
'
'       sConnectionType -  type of Connection Objects to check
'                          "Assembly" or "Physical" or "Logical" or (other)
'
' Outputs:
'       zConnectionData -   list of all connected Parts connection data
'
'
' Notes:
'   two Connectables can be connected thru a
'   an "Assembly Connection" object,  or a "Physical Connection" object
'   or a "Logical Connection" object
'
'********************************************************************
Public Function GetAllConnectables(Connectable1 As IJConnectable, _
                                   eConnectionType As eAppConnectionTypes, _
                                   zConnectionData() As ConnectionData) As Long
Const MT = "GetAllConnectables"
On Error GoTo ErrorHandler

Dim bReleated As Boolean

Dim nConnectables As Long
Dim nAppConnections As Long

Dim oPort1 As IJPort
Dim oStructPort As IJStructPort
Dim oRelatedPort1 As IJPort
Dim oConnectedPort1 As IJPort
Dim oAppConnection1 As IJAppConnection
Dim oStructConnectable1 As IJStructConnectable
Dim ConnectableConnectedPorts1 As IJElements

On Error Resume Next
nConnectables = 0
GetAllConnectables = 0

' Retrieve a list of Connected Ports from the given Connectable object
If TypeOf Connectable1 Is IJStructConnectable Then
    Set oStructConnectable1 = Connectable1
    oStructConnectable1.enumConnectedPortsInGraph ConnectableConnectedPorts1, 0
    Set oStructConnectable1 = Nothing
Else
    Connectable1.enumConnectedPorts ConnectableConnectedPorts1, 0
End If

' Loop thru each Connected Port from the given Connectable object
For Each oConnectedPort1 In ConnectableConnectedPorts1
    ' For current Port
    ' Find all AppConnections of requested type (Assembly,Physical,Logical)
    ' for each AppConnection,
    '   get the to (other) connectable and Port
    Set oPort1 = oConnectedPort1
    If Not oPort1 Is Nothing Then
        nAppConnections = GetPortAppConnections(oPort1, _
                                                eConnectionType, _
                                                nConnectables, _
                                                zConnectionData, _
                                                Connectable1)

        Set oPort1 = Nothing
    End If
    
    Set oConnectedPort1 = Nothing
Next

ConnectableConnectedPorts1.Clear
Set ConnectableConnectedPorts1 = Nothing
    
GetAllConnectables = nConnectables

Exit Function
    
ErrorHandler:
   Err.Raise LogError(Err, MODULE, MT).Number
End Function

'********************************************************************
' Routine: GetPortsAppConnection
'
' Description:
'   Given a Port object
'   return the AppConnection object
'
' Inputs:
'       Port1           -  Port to find AppConnections for
'       sConnectionType -  type of AppConnection Objects to find
'                          "Assembly" or "Physical" or "Logical" or (other)
'       nConnectionData -   size of zConnectionData array
'
'   Optional
'       vConnectable1   -   Port1 Connectalbe object
'
' Outputs:
'       nConnectionData -   size of zConnectionData array
'       zConnectionData -   array (list) of connection data
'                               AppConnection   - AppConnection object
'                               ConnectingPort  - Connected from Port (= Port1)
'                               ToConnectable   - Connected to Object
'                               ToConnectedPort - Connected to Port
'
'
' Notes:
'   two Connectables can be connected thru a
'   an "Assembly Connection" object,  or a "Physical Connection" object
'   or a "Logical Connection" object
'
'********************************************************************
Public Function GetPortAppConnections(Port1 As IJPort, _
                                      eConnectionType As eAppConnectionTypes, _
                                      nConnectionData As Long, _
                                      zConnectionData() As ConnectionData, _
                                      Optional vConnectable1 As Variant = Null) As Long
Const MT = "ConnectedParts.GetPortsAppConnection"
On Error GoTo ErrorHandler

Dim oConnectable1 As IJConnectable

Dim oPort11 As IJPort
Dim oPort12 As IJPort
Dim oConnectable11 As IJConnectable
Dim oConnectable12 As IJConnectable

Dim oAppConnection1 As IJAppConnection
Dim PortAppConnections1 As IJElements
Dim AppConnectionPorts1 As IJElements

Dim nConnected As Long
Dim nAppConnections As Long
Dim oAssemblyConnection As IJAssemblyConnection
Dim oStructPhysicalConnection As IJStructPhysicalConnection
    
On Error Resume Next
'$$$ ***********************************************************************
Dim iStart As Long
iStart = nConnectionData
'$$$ ***********************************************************************

nAppConnections = 0
GetPortAppConnections = 0
    
' Retrieve the Connectable object from the given Port
Set oConnectable1 = vConnectable1
If oConnectable1 Is Nothing Then
    Set oConnectable1 = Port1.Connectable
End If

' Retrieve a list of AppConnections from the first given Port
' (a list of all AppConnections that the current Port is involved in)
Port1.enumConnections PortAppConnections1, 0, 0
        
If Not PortAppConnections1 Is Nothing Then
  If PortAppConnections1.Count > 0 Then
    ' Loop thru each AppConnection that the current Port is involved in
    ' (each AppConnection should have 2 Ports involved in the Connection)
    For Each oAppConnection1 In PortAppConnections1
                
        If Not oAppConnection1 Is Nothing Then
            Set oAssemblyConnection = oAppConnection1
            Set oStructPhysicalConnection = oAppConnection1
            If eConnectionType = AppConnectionType_Assembly Then
                If oAssemblyConnection Is Nothing Then
                    Set oAppConnection1 = Nothing
                End If
            ElseIf eConnectionType = AppConnectionType_Physical Then
                If oStructPhysicalConnection Is Nothing Then
                    Set oAppConnection1 = Nothing
                End If
            ElseIf eConnectionType = AppConnectionType_Logical Then
                If Not oAssemblyConnection Is Nothing Or _
                   Not oStructPhysicalConnection Is Nothing Then
                    Set oAppConnection1 = Nothing
                End If
            End If
                        
            Set oAssemblyConnection = Nothing
            Set oStructPhysicalConnection = Nothing
        End If
            
        If Not oAppConnection1 Is Nothing Then
            ' Retrieve a list of Ports from the AppConnection
            ' (should contain the 2 Ports involved in the Connection)
            oAppConnection1.enumPorts AppConnectionPorts1
            If AppConnectionPorts1.Count > 0 Then
                ' Retrieve the Connectable objects from the
                ' the 2 Ports involved in the Connection
                ' The returned Ports and Connectables are ordered
                '   Connectable11 = Connectable1
                '   Port11 is releated to Connectable11
                '   Connectable12 = other Connectable
                '   Port12 is releated to Connectable12
                nConnected = GetPortConnections(AppConnectionPorts1, _
                                               oPort11, oConnectable11, _
                                               oPort12, oConnectable12, _
                                               oConnectable1)
                                               
                'check if the current AppConnection is the requested
                'i.e.
                '   Port1 and Port2 are the same as Port11 and Port12
                '   Connectable1/Connectable1 = Connectable11/Connectable12
                
                If nConnected = 2 Then
                    nAppConnections = nAppConnections + 1
                    AddConnectionData nConnectionData, _
                                      oAppConnection1, _
                                      oPort11, _
                                      oConnectable12, oPort12, _
                                      zConnectionData
                End If
                        
                Set oPort11 = Nothing
                Set oPort12 = Nothing
                Set oConnectable11 = Nothing
                Set oConnectable12 = Nothing
            End If
                    
            AppConnectionPorts1.Clear
    End If
        
    Set oAppConnection1 = Nothing
    Next
  End If
End If

'$$$ ***********************************************************************
Dim sText As String
Dim iIndex As Long
Dim oNamedItem As IJNamedItem

If oConnectable1 Is Nothing Then
    sText = "Connectable1: is Nothing"
Else
    Set oNamedItem = oConnectable1
    sText = "Connectable1: " & oNamedItem.Name
    Set oNamedItem = Nothing
End If
sText = sText & " ... nAppConnections= " & Str(nAppConnections)

If nAppConnections > 0 Then
    iStart = iStart + 1
    For iIndex = iStart To nConnectionData
        Set oNamedItem = zConnectionData(iIndex).ToConnectable
        sText = sText & vbCrLf & _
                "iIndex ... " & Str(iIndex) & _
                "    oToConnectable ... " & oNamedItem.Name & _
                " ... " & Str(iStart)
        Set oNamedItem = Nothing
    Next iIndex
End If

'$$$Debug $$$ _
MsgBox sText, , "ConnectedParts::GetPortAppConnections"
'$$$ ***********************************************************************
        
If Not PortAppConnections1 Is Nothing Then
    PortAppConnections1.Clear
End If
Set oConnectable1 = Nothing

GetPortAppConnections = nAppConnections


Exit Function
    
ErrorHandler:
   Err.Raise LogError(Err, MODULE, MT).Number
End Function


'********************************************************************
' Routine: GetPortConnections
'
' Description: GetPortConnections
'
' Inputs:
'       Ports -  List of Ports used in connecting objects
'
'       Optional vFirstConnectable -  optional arugument
'                                     Connectable object that is used to
'                                     order the returned output
'                                     Port1,Connectionable1 will be the
'                                     vFirstConnectable object
'
' Outputs:
'       Port1           -  Port used in connecting objects
'       Connectable1    -  Connectable that Port1 belongs to
'       Port2           -  Port used in connecting objects
'       Connectable2    -  Connectable that Port2 belongs to
'
'
' Notes:
'
'********************************************************************
Public Function GetPortConnections(Ports As IJElements, _
                        Port1 As IJPort, _
                        Connectable1 As IJConnectable, _
                        Port2 As IJPort, _
                        Connectable2 As IJConnectable, _
                        Optional vFirstConnectable As Variant = Null) As Long
Const MT = "GetPortConnections"
On Error GoTo ErrorHandler

Dim nPorts As Integer

Dim Port As IJPort
Dim Connectable As IJConnectable
Dim FirstConnectable As IJConnectable

On Error Resume Next
Set Connectable1 = Nothing
Set Connectable2 = Nothing
GetPortConnections = 0

' Verify that the Ports list contain Ports
nPorts = 0
If Ports.Count < 1 Then
    Exit Function
End If
    
' Loop thru each Port in the given Ports list
For Each Port In Ports
    ' Retrieve the Connectable object from the Port
    Set Connectable = Port.Connectable
    If Not Connectable Is Nothing Then
        
        ' save first and second Connectable objects
        nPorts = nPorts + 1
        If nPorts = 1 Then
            Set Port1 = Port
            Set Connectable1 = Connectable
        ElseIf nPorts = 2 Then
            Set Port2 = Port
            Set Connectable2 = Connectable
        End If
            
        Set Connectable = Nothing
    End If
            
Next
    
' set the number Ports processed
GetPortConnections = nPorts

' a Valid Connection should consist of exactly 2 Ports
If nPorts = 2 Then
    ' check if the order of connectables to be returned
    ' is based on optional Connectable input
    Set FirstConnectable = vFirstConnectable
    If Not FirstConnectable Is Nothing Then
        
        ' check if the order of Connectable is to be reversed
        If Connectable2 Is FirstConnectable Then
            Set Port = Nothing
            Set Connectable = Nothing
            Set Port = Port1
            Set Connectable = Connectable1
            
            Set Port1 = Nothing
            Set Connectable1 = Nothing
            Set Port1 = Port2
            Set Connectable1 = Connectable2
            
            Set Port2 = Nothing
            Set Connectable2 = Nothing
            Set Port2 = Port
            Set Connectable2 = Connectable
            
            Set Port = Nothing
            Set Connectable = Nothing
        End If
        
    End If
End If
    
Exit Function
    
ErrorHandler:
   Err.Raise LogError(Err, MODULE, MT).Number
End Function

'********************************************************************
' Routine: AddConnectionData
'
' Description:
'   add connection data to ConnectionData structure array
'
' Inputs:
'       lSize            -  current size of ConnectionData array
'       oAppConnection   -  Connection object used to connect the objects
'       oConnectingPort  -  Port from the Connecting from object
'       oToConnectable   -  the Connected To object
'       oToConnectedPort -  Port from Connected To object
'
'       zConnectionData -  array of ConnectionData structures
'
' Outputs:
'       zConnectionData -  array of ConnectionData structures
'
'
' Notes:
'   the size of the ConnectionData array is re-dimensioned
'
'********************************************************************
Public Function AddConnectionData(lSize As Long, _
                                  oAppConnection As IJAppConnection, _
                                  oConnectingPort As IJPort, _
                                  oToConnectable As IJConnectable, _
                                  oToConnectedPort As IJPort, _
                                  ByRef zConnectionData() As ConnectionData)
Const MT = "ConnectedParts.GetPortConnections"
On Error GoTo ErrorHandler

    On Error GoTo ErrorHandler

    lSize = lSize + 1
    ReDim Preserve zConnectionData(lSize)

    Set zConnectionData(lSize).AppConnection = oAppConnection
    Set zConnectionData(lSize).ConnectingPort = oConnectingPort
    Set zConnectionData(lSize).ToConnectable = oToConnectable
    Set zConnectionData(lSize).ToConnectedPort = oToConnectedPort

Exit Function
    
ErrorHandler:
   Err.Raise LogError(Err, MODULE, MT).Number
End Function

Public Function IsReinforcedBracket(oObject As Object, _
                                     existingType As ShpStrBracketReinforcementType) As Boolean
    On Error GoTo ErrorHandler
    Const MT = "IsReinforcedBracket"
    
    Dim exists As Boolean
    existingType = BRACKETREINFORCEMENTTYPE_None
    exists = False
   
    If (Not TypeOf oObject Is IJPlate) And _
             (Not IsBracket(oObject)) Then
       Exit Function
    End If
   
   Dim oBracketRoot As IJPlate
   Dim oStructDetailHelper As StructDetailHelper
   Set oStructDetailHelper = New StructDetailHelper
   
    oStructDetailHelper.IsPartDerivedFromSystem oObject, oBracketRoot, True

    If oBracketRoot Is Nothing Then
          Set oBracketRoot = oObject
     End If
    
    Dim pChildrenColl As IJDObjectCollection
    Dim oParent As IJDesignParent
    
    Set oParent = oBracketRoot
    
    oParent.GetChildren pChildrenColl
    
    Dim oChild As Object
    
    For Each oChild In pChildrenColl
        Dim oChildObj As Object
        Set oChildObj = oChild
        If TypeOf oChildObj Is IJBracketReinforcementByRule Then
              exists = True
              If TypeOf oChildObj Is IJERSystem Then
                  existingType = BRACKETREINFORCEMENTTYPE_EdgeReinforcement
              Else
                  existingType = BRACKETREINFORCEMENTTYPE_BucklingStiffener
              End If
              Exit For
         End If
     Next oChild
    
       
    IsReinforcedBracket = exists
    
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, MT).Number
    
End Function

'******************************************************************************
'This Method doesnt work for Stand Alone Brackets....
'Works only for Tripping Bracket and Brackt By Plane
'Inputs -- Pass the Leaf plate part(Detailed Plate)
               

'*****************************************************************************
Public Function GetBracketFreeEdgeLateralFacePort(oBracket As IJPlate) As IJPort
  On Error GoTo ErrorHandler
    Const MT = "GetBracketFreeEdgeLateralFacePort"
    
    If Not IsBracket(oBracket) Then
       Exit Function
    End If
    
    Dim oFreeEdgePort As Object
    Dim oFreeEdgeStructPort As IJStructPort
    Dim oBracketFreeEdgeLateralFace As IJPort
    Dim oRootBracket As IJPlate
    Dim oStructDetailHelper As StructDetailHelper
    
    Set oStructDetailHelper = New StructDetailHelper
   
    oStructDetailHelper.IsPartDerivedFromSystem oBracket, oRootBracket, True

    If oRootBracket Is Nothing Then
          Set oRootBracket = oBracket
     End If
     
    Dim oBracketHlpr As GSCADCreateModifyUtilities.IJBracketAttributes
    Set oBracketHlpr = New GSCADCreateModifyUtilities.PlateUtils
    
    Set oFreeEdgePort = oBracketHlpr.GetBracketFreeEdgePort(oRootBracket)
    Set oBracketFreeEdgeLateralFace = Nothing
    
    If Not oFreeEdgePort Is Nothing Then
     If TypeOf oFreeEdgePort Is IJPort Then
       Set oFreeEdgeStructPort = oFreeEdgePort
      
       Dim ePortType As JS_TOPOLOGY_FILTER_TYPE
       ePortType = JS_TOPOLOGY_FILTER_SOLID_LATERAL_LFACES
    
       Dim GeomUtils As IJTopologyLocate
       Set GeomUtils = New TopologyLocate
     
       Dim oListOfPorts As IEnumUnknown
       GeomUtils.GetNamedPorts oBracket, ePortType, oListOfPorts
    
       Dim oCollectionOfPorts As Collection
       Dim ConvertUtils As New CCollectionConversions
       ConvertUtils.CreateVBCollectionFromIEnumUnknown oListOfPorts, oCollectionOfPorts
       
       Dim oPort As IJPort
       Dim oStructPort As IJStructPort
       
       If Not oCollectionOfPorts Is Nothing Then
          If oCollectionOfPorts.Count > 0 Then
              For Each oPort In oCollectionOfPorts
                  Set oStructPort = oPort
                  If oStructPort.OperationID = oFreeEdgeStructPort.OperationID And _
                     oStructPort.OperatorID = oFreeEdgeStructPort.OperatorID Then
                        Dim oTempPort As IJPort
                        Dim oModelBody As IJDModelBody
                        
                        Set oTempPort = oStructPort
                        Set oModelBody = oTempPort.Geometry
                        
                        Dim pStartPt As IJDPosition
                        Dim pEndPt As IJDPosition
                        Dim dMinDist As Double
                        
                        oModelBody.GetMinimumDistance oFreeEdgePort.Geometry, pStartPt, pEndPt, dMinDist
                        
                        If dMinDist < 0.00005 Then
                           Set oBracketFreeEdgeLateralFace = oTempPort
                           Exit For
                        End If
                  End If
              Next oPort
          End If
       End If
     End If
    End If
    
    Set GetBracketFreeEdgeLateralFacePort = oBracketFreeEdgeLateralFace
    
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, MT).Number
End Function
'*************************************************************************
'Function
'IsFlangeIn
'
'Purpose:
'   Determines is the flange for a giving support is in the direction of
'   the bracket.
'Inputs :
'   oProfileSupport - any profile which acts as support to the bracket
'   oBracket - a bracket to which the Flange Direction is Needed
'Returns:
'   True   - if Flange is IN (orientation is towards the direction of the bracket.)
'   False - if Flange is OUT (orientation is away from the direction of the bracket.)
'*********************************************************************************************
Public Function IsFlangeTowardsBracket(oProfileSupport As Object, oBracket As IJPlate) As Boolean

On Error GoTo ErrorHandler:

'Assume true...
IsFlangeTowardsBracket = True
                 
If Not IsBracket(oBracket) Then
    Exit Function
End If

'*********************************************************************************************
 '1. Get interstion point between supports and bracket plane
Dim oBracketUtils As GSCADCreateModifyUtilities.IJBracketAttributes
Set oBracketUtils = New GSCADCreateModifyUtilities.PlateUtils

Dim oBracketPlate As IJPlate
Set oBracketPlate = oBracket

Dim oIntersect As IJDPosition
Dim GeomUtils As New TopologyLocate

Dim oS1 As Object
Dim oS2 As Object
Dim oS3 As Object
Dim oS4 As Object
Dim oS5 As Object

Dim oS1Port As IJPort
Dim oS2Port As IJPort
Dim oS3Port As IJPort
Dim oS4Port As IJPort
Dim oS5Port As IJPort

Dim uBracketVec As IJDVector
Dim vBracketVec As IJDVector
Dim oRefPlane As IJPlane
Dim nNumSupports As Long

GetBracketSupports oBracketPlate, oS1, oS2, oS3, oS4, oS5, oS1Port, oS2Port, oS3Port, oS4Port, oS5Port

GetBracketPlaneAndUVVectors oBracketPlate, oRefPlane, uBracketVec, vBracketVec, nNumSupports



Dim oStructDetailHlpr As New StructDetailHelper
Dim oFlangeOrientVec As IJDVector
Dim oWebOrientVec As IJDVector

Dim oOrientationPosition As IJDPosition

Dim oRootProfile As Object
Dim oRootStiffenerSystem As IJStiffenerSystem

oStructDetailHlpr.IsPartDerivedFromSystem oProfileSupport, oRootProfile, True
If oRootProfile Is Nothing Then
    Set oRootProfile = oProfileSupport
End If

Set oIntersect = GeomUtils.FindIntersectionPoint(oRootProfile, oRefPlane)
'CMT if support 1 is a profile on a deck,the rootprofile doesn't necessarily intersect, so adding an if.
If oIntersect Is Nothing Then
    Set oIntersect = GeomUtils.FindIntersectionPoint(oRootProfile, oS1)
End If
Set oOrientationPosition = oIntersect
'*********************************************************************************************
'2. Get the direction vector for first and seondary orientation of profile sent...
Dim oProfilehelper As IJProfileAttributes
Set oProfilehelper = New ProfileUtils
oProfilehelper.GetProfileOrientation oRootProfile, oOrientationPosition, oFlangeOrientVec, oWebOrientVec
 
'*********************************************************************************************
'3. Found out what support was sent and choice either the U or V vector from the bracket...
Dim oBracketDir As IJDVector

If oRootProfile Is oS1 Then
    ' Bracket is attached to flange, therefore no need to determine flange direction...
    Exit Function
ElseIf (oRootProfile Is oS2) Or (oRootProfile Is oS3) Then
    
    Set oBracketDir = uBracketVec

ElseIf (oRootProfile Is oS4) Or (oRootProfile Is oS5) Then

    Set oBracketDir = vBracketVec
       
End If

'*********************************************************************************************
'4. Determine if the direction between the flange direction is in the same direction as the
'    the bracket...
Dim fDot As Double

fDot = oBracketDir.x * oFlangeOrientVec.x + _
           oBracketDir.y * oFlangeOrientVec.y + _
           oBracketDir.z * oFlangeOrientVec.z

If fDot < 0 Then
    IsFlangeTowardsBracket = True
Else
    IsFlangeTowardsBracket = False
End If

Exit Function
'Clean up...
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "IsFlangeTowardsBracket").Number
    Err.Clear
End Function




Public Sub GetBracketPlaneAndUVVectors(oBracket As Object, oRefPlane As IJPlane, uVector As IJDVector, vVector As IJDVector, lnumSupports As Long)
  On Error GoTo ErrorHandler
    Const MT = "GetBracketPlaneAndUVVectors"
    
    If Not TypeOf oBracket Is IJPlate And _
       Not IsBracket(oBracket) Then
      Exit Sub
    End If
    
    Dim oS1 As Object
    Dim oS2 As Object
    Dim oS3 As Object
    Dim oS4 As Object
    Dim oS5 As Object
    
    Dim oBracketU As IJDVector
    Dim oBracketV As IJDVector
    Dim oPlane As IJPlane
    
    Set oBracketU = Nothing
    Set oBracketV = Nothing
    Set oPlane = Nothing
    
     ' --------------------------------
    ' If a part, see if system-derived
    ' --------------------------------
    Dim oStructDetailHelper As StructDetailHelper
    Set oStructDetailHelper = New StructDetailHelper
    
    Dim oRootBracket As Object ' either standalone or system
    If TypeOf oBracket Is IJPlatePart Then
        oStructDetailHelper.IsPartDerivedFromSystem oBracket, oRootBracket, True

        If oRootBracket Is Nothing Then
            Set oRootBracket = oBracket
        End If
    End If
    
    ' ----------------------------------------------------------------------
    ' If root object is a bracket-type plate part, it is a standalone object
    ' ----------------------------------------------------------------------
    Dim oPlateUtil As IJPlateAttributes
    Set oPlateUtil = New PlateUtils
    
    Dim plateType As StructPlateType
    plateType = CollarPlate
    If TypeOf oBracket Is IJPlate Then
        Dim oPlate As IJPlate
        Set oPlate = oBracket
        plateType = oPlate.plateType
    End If
        
    If TypeOf oRootBracket Is IJPlatePart And plateType = BracketPlate Then
        Dim oSDOBracket As New StructDetailObjects.Bracket
        Set oSDOBracket.object = oBracket

        Dim oPort1 As IJPort
        Dim oPort2 As IJPort
        Dim nSupports As Long

        oSDOBracket.GetInputs nSupports, oPlane, oPort1, oPort2, oS3, oS4, oS5
        
        lnumSupports = nSupports
        
        Dim oPointUtils As IJPointAttributes
        Set oPointUtils = New PointUtils
        
        Dim ppIntersection As IJDPosition
        Dim oRootParent As Object
        
        Set oS1 = oPort1.Connectable
        Set oS2 = oPort2.Connectable
        
        oStructDetailHelper.IsPartDerivedFromSystem oS1, oRootParent, True
        If Not oRootParent Is Nothing Then
            Set oS1 = oRootParent
            Set oRootParent = Nothing
        End If
        
        oStructDetailHelper.IsPartDerivedFromSystem oS2, oRootParent, True
        If Not oRootParent Is Nothing Then
            Set oS2 = oRootParent
            Set oRootParent = Nothing
        End If
        oPointUtils.GetPlaneAndSupportsIntersection oS1, oS2, oPlane, ppIntersection
        
        Dim oToplogy As New TopologyLocate
        
        Dim oPointOnS1 As IJDPosition
        Dim oPointOnS2 As IJDPosition
        
        Set oPointOnS1 = oToplogy.FindIntersectionPoint(oPlane, oPort1)
        Set oPointOnS2 = oToplogy.FindIntersectionPoint(oPlane, oPort2)
        
        Dim oUApproxVector As IJDVector
        Dim oVApproxVector As IJDVector
        
        Set oUApproxVector = New dVector
        Set oVApproxVector = New dVector
        
        oUApproxVector.Set oPointOnS1.x - ppIntersection.x, _
                           oPointOnS1.y - ppIntersection.y, _
                           oPointOnS1.z - ppIntersection.z
                           
        oVApproxVector.Set oPointOnS2.x - ppIntersection.x, _
                           oPointOnS2.y - ppIntersection.y, _
                           oPointOnS2.z - ppIntersection.z
                           
        Dim oPartInfo As New PartInfo
        Dim oPort1Normal As IJDVector
        Dim oPort2Normal As IJDVector
        Dim bApproximationUsed As Boolean
        
        Set oPort1Normal = oPartInfo.GetPortNormal(oPort1, bApproximationUsed)
        Set oPort2Normal = oPartInfo.GetPortNormal(oPort2, bApproximationUsed)
        
        'Get Bracket Plane Nomal
        Dim oBracketPlaneNormal As IJDVector
        Set oBracketPlaneNormal = New dVector
        
        Dim dBrackX As Double
        Dim dBrackY As Double
        Dim dBrackZ As Double
        
        oPlane.GetNormal dBrackX, dBrackY, dBrackZ
        
        oBracketPlaneNormal.Set dBrackX, dBrackY, dBrackZ
       
        Dim oCrossS1Vector As IJDVector
        Dim oCrossS2Vector As IJDVector
        
        Set oCrossS1Vector = oBracketPlaneNormal.Cross(oPort1Normal)
        Set oCrossS2Vector = oBracketPlaneNormal.Cross(oPort2Normal)
        
        Set oBracketU = New dVector
        Set oBracketV = New dVector
        
        If oCrossS1Vector.Dot(oUApproxVector) > 0 Then
               oBracketU.Set oCrossS1Vector.x, _
                             oCrossS1Vector.y, _
                             oCrossS1Vector.z
        Else
              oBracketU.Set -oCrossS1Vector.x, _
                            -oCrossS1Vector.y, _
                            -oCrossS1Vector.z
        End If
        
        If oCrossS2Vector.Dot(oVApproxVector) > 0 Then
                oBracketV.Set oCrossS2Vector.x, _
                             oCrossS2Vector.y, _
                             oCrossS2Vector.z
        Else
               oBracketV.Set -oCrossS2Vector.x, _
                             -oCrossS2Vector.y, _
                             -oCrossS2Vector.z
        End If
    ' --------------------------------------
    ' Otherwise, check if a tripping bracket
    ' --------------------------------------
    ElseIf oPlateUtil.IsTrippingBracket(oRootBracket) Then
       
        Dim oAE As IJPlaneByElements_AE
    
        oPlateUtil.GetInput_TrippingBracket oRootBracket, oBracketU, oBracketV, oS1, oS2, oS3, oAE
        
        If TypeOf oRootBracket Is IJPlane Then
            Set oPlane = oRootBracket
        End If
        
        If Not oS3 Is Nothing Then
           lnumSupports = 3
        Else
           lnumSupports = 4
        End If
    ' ------------------------------------
    ' Finally, check if a bracket-by-plane
    ' ------------------------------------
    ElseIf oPlateUtil.IsBracketByPlane(oRootBracket) Then
        
        Dim oBracketPlateSystem As IJBracketPlateSystem
        Set oBracketPlateSystem = oRootBracket
        
        Set oPlane = oBracketPlateSystem.BracketSketchingPlane
        Set oBracketU = oBracketPlateSystem.BracketUVector
        Set oBracketV = oBracketPlateSystem.BracketVVector
        
        lnumSupports = oBracketPlateSystem.NumberOfSupports
        Set oBracketPlateSystem = Nothing
    End If
    
    Set oRefPlane = oPlane
    Set uVector = oBracketU
    Set vVector = oBracketV
    
  Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, MT).Number

End Sub

Public Function IsConnectedToObjectOrChildren(oPort As IJPort, oObject As Object) As Boolean

    On Error GoTo ErrorHandler
'    sMETHOD = "IsConnectedToObjectOrChildren"
    
    IsConnectedToObjectOrChildren = False
    
    If oObject Is Nothing Then
        Exit Function
    End If
    
    Dim oStructPortEx As IJStructPortEx
    Dim oConnPort As IJPort
    Set oStructPortEx = oPort
    Set oConnPort = oStructPortEx.ConnectablePort
    
    
    Dim oChildPartEnum As IEnumUnknown
    Dim oStructDetailHelper As StructDetailHelper
    Set oStructDetailHelper = New StructDetailHelper
    
    If TypeOf oObject Is IJSystem Then
        oStructDetailHelper.GetPartsDerivedFromSystem oObject, oChildPartEnum, True
    End If
    
    Dim oChildPartCol As Collection
    If oChildPartEnum Is Nothing Then
        Set oChildPartCol = New Collection
        oChildPartCol.Add oObject
    Else
        Dim ConvertUtils As CCollectionConversions
        Set ConvertUtils = New CCollectionConversions
        ConvertUtils.CreateVBCollectionFromIEnumUnknown oChildPartEnum, oChildPartCol
    End If
    
    Dim oConnections As IJElements
    Dim oConnection As IJAppConnection
    Dim oPortsInConn As IJElements
    Dim oPortInConn As IJPort
    Dim oConnectedObj As Object
    Dim oChildObj As Object
    
    oConnPort.enumConnections oConnections, ConnectionAssembly, ConnectionStandard
    
    If Not oConnections Is Nothing Then
        For Each oConnection In oConnections
            Set oPortsInConn = Nothing
            oConnection.enumPorts oPortsInConn
            
            For Each oPortInConn In oPortsInConn
                If Not oPortInConn Is oConnPort Then
                
                    Set oConnectedObj = oPortInConn.Connectable
                
                    For Each oChildObj In oChildPartCol
                        If oChildObj Is oConnectedObj Then
                            IsConnectedToObjectOrChildren = True
                            Exit Function
                        End If
                    Next oChildObj
                End If
            Next oPortInConn
        Next oConnection
    End If
    
    Exit Function
    
ErrorHandler:
'    Err.Raise LogError(Err, MODULE, sMETHOD).Number

End Function

'***********************************************************************************************
'    Function      : UpdateDependentCornersSeam
'
'    Description   : This method helps to change a non seam item to seam item
'
'***********************************************************************************************
Public Sub UpdateDependentCornersSeam(oAssyConn As Object)

Const sMETHOD = "UpdateDependentCornersSeam"
On Error GoTo ErrHandler
        
    ' ----------------------
    ' Check for valid inputs
    ' ----------------------
    If Not TypeOf oAssyConn Is IJAssemblyConnection Then
        Exit Sub
    End If

    Dim oSDOAssyConn As StructDetailObjects.AssemblyConn
    Set oSDOAssyConn = New StructDetailObjects.AssemblyConn
    Set oSDOAssyConn.object = oAssyConn
    
    If Not (oSDOAssyConn.FromDesignSeam Or oSDOAssyConn.FromPlanningSeam Or oSDOAssyConn.FromStrakingSeam) Then
        Exit Sub
    End If
    
    'Get seam and use it for checking distance from the CF
    Dim oStructDetailConnUtils As StructDetailConnectionUtil
        
    Set oStructDetailConnUtils = New StructDetailConnectionUtil
    Dim oSeamObject As Object
    Dim oSeamCurve As IJCurve
    On Error Resume Next
    Set oSeamObject = oStructDetailConnUtils.GetTheSeamFromConnection(oAssyConn)
    Set oSeamCurve = oSeamObject
    On Error GoTo ErrHandler
    If oSeamCurve Is Nothing Then GoTo ErrHandler
    
    ' -------------------------------------------------------------------------------
    ' Retreive all of the Assembly Connections for connected objects
    ' -------------------------------------------------------------------------------
    Dim nConnsToPart1 As Long
    Dim nConnsToPart2 As Long
    Dim aPart1ACData() As ConnectionData
    Dim aPart2ACData() As ConnectionData

    Dim oSDO_Helper As StructDetailObjects.Helper
    Set oSDO_Helper = New StructDetailObjects.Helper

    oSDO_Helper.Object_AppConnections oSDOAssyConn.ConnectedObject1, _
                                      AppConnectionType_Assembly, _
                                      nConnsToPart1, _
                                      aPart1ACData

    oSDO_Helper.Object_AppConnections oSDOAssyConn.ConnectedObject2, _
                                      AppConnectionType_Assembly, _
                                      nConnsToPart2, _
                                      aPart2ACData
                                      
    'Approach:
    ' Step 1: To-Connectable parts of aPart1ACData are added to collection.
    ' Step 2: Get each To-connectable from the other ACdata i.e. aPart2ACData array and if this part is found
    ' in the collection prepared in step 1 (=> the part has AC with oSDOAssyConn.ConnectedObject1 and also has AC with oSDOAssyConn.ConnectedObject2)
    ' check for corner features on this part.
    Dim i As Long
    Dim oPart1AC_ConnParts As IJDObjectCollection
    Set oPart1AC_ConnParts = New JObjectCollection
    For i = 1 To nConnsToPart1
        If Not TypeOf aPart1ACData(i).ToConnectable Is ISPSMemberPartPrismatic Then
            If Not oPart1AC_ConnParts.Contains(aPart1ACData(i).ToConnectable) Then
                oPart1AC_ConnParts.Add aPart1ACData(i).ToConnectable, i
            End If
        End If
    Next i

    ' -------------
    ' For each part
    ' -------------
    Dim oPartSupport As IJPartSupport
    Set oPartSupport = New PartSupport

    Dim oFeaturesList As Collection
    Dim oFeature As IUnknown
    Dim oStructFeature As IJStructFeature
    Dim featureType As StructFeatureTypes
    
    'Iterate each AC to get part
    For i = 1 To nConnsToPart2
        If oPart1AC_ConnParts.Contains(aPart2ACData(i).ToConnectable) Then
            Set oPartSupport.Part = aPart2ACData(i).ToConnectable
            oPartSupport.GetFeatures oFeaturesList
    
            ' ----------------------------
            ' For each feature on the part
            ' ----------------------------
            Dim nFeatures As Long
            nFeatures = oFeaturesList.Count
            
            Dim j As Long
            For j = 1 To nFeatures
                Set oFeature = oFeaturesList.Item(j)
                
                ' -------------------------
                ' If slot or corner feature
                ' -------------------------
                If TypeOf oFeature Is IJStructFeature Then
                    Set oStructFeature = oFeature
                    featureType = oStructFeature.get_StructFeatureType
    
                    Select Case featureType
                        Case SF_CornerFeature
                            Dim oLocation As IJDPosition
                            Dim oFeatureRng As IJRangeAlias
                            Dim gRngBox As GBox
                            Set oFeatureRng = oFeature
                            gRngBox = oFeatureRng.GetRange()
                            Set oLocation = New DPosition
                            
                            oLocation.Set (gRngBox.m_low.x + gRngBox.m_high.x) / 2, _
                                            (gRngBox.m_low.y + gRngBox.m_high.y) / 2, _
                                            (gRngBox.m_low.z + gRngBox.m_high.z) / 2
                            ' --------------------------------------------
                            ' Check distance tolerance and update the feature
                            ' --------------------------------------------
                            'Project the location of CF onto seam
                                Dim dDista As Double
                            Dim dSrcX As Double
                            Dim dSrcY As Double
                            Dim dSrcZ As Double
                            Dim dInX As Double
                            Dim dInY As Double
                            Dim dInZ As Double
                            oSeamCurve.DistanceBetween oLocation, dDista, dSrcX, dSrcY, dSrcZ, dInX, dInY, dInZ
                            If dDista < 1 Then
                                ForceUpdateSmartItem oFeature
                            End If
                    End Select
                End If
            Next j
        End If
    Next i
    
    Set oFeaturesList = Nothing

    Exit Sub

ErrHandler:
  Err.Raise LogError(Err, MODULE, sMETHOD).Number
End Sub



'***********************************************************************************************
'    Function      : UpdateCFWhenSeamIsDeleted
'
'    Description   : This method helps to change a seam item to a non seam item
'
'
'***********************************************************************************************
Public Sub UpdateCFWhenSeamIsDeleted(oAssyConn As Object)

Const sMETHOD = "UpdateCFWhenSeamIsDeleted"
On Error GoTo ErrHandler
       
    ' ----------------------
    ' Check for valid inputs
    ' ----------------------
    If Not TypeOf oAssyConn Is IJAssemblyConnection Then
        Exit Sub
    End If
    
    Dim oSDOAssyConn As New StructDetailObjects.AssemblyConn
    Set oSDOAssyConn.object = oAssyConn
    
   
    Dim oConnectedObject1 As Object
    Dim oConnectedObject2 As Object
    
    Set oConnectedObject1 = oSDOAssyConn.ConnectedObject1
    Set oConnectedObject2 = oSDOAssyConn.ConnectedObject2
    
    'assume the other connected object is object1
    
    Dim oPartSupport As IJPartSupport
    Set oPartSupport = New PartSupport
    
    Dim oFeaturesList As Collection
    Dim oFeature As IUnknown
    Dim oStructFeature As IJStructFeature
    Dim featureType As StructFeatureTypes
    
    Dim oCornerFeature As IJSDOCornerFeature
    
    Set oCornerFeature = New StructDetailObjectsex.CornerFeature
    
    Dim oStructFeatUtils As IJSDFeatureAttributes
    
    Set oPartSupport.Part = oConnectedObject1
    oPartSupport.GetFeatures oFeaturesList

    Dim nFeatures As Long
    nFeatures = oFeaturesList.Count
    
    Dim bForceUpdate As Boolean
    
    bForceUpdate = False
    
    Dim j As Long
    For j = 1 To nFeatures
        Set oFeature = oFeaturesList.Item(j)
        
        ' -------------------------
        ' If slot or corner feature
        ' -------------------------
        If TypeOf oFeature Is IJStructFeature Then
            Set oStructFeature = oFeature
            featureType = oStructFeature.get_StructFeatureType

            Select Case featureType
                Case SF_CornerFeature
                    Set oCornerFeature.object = oFeature
                    Set oStructFeatUtils = New SDFeatureUtils
                    
                    Dim oNamedItem As IJNamedItem
                    Set oNamedItem = oCornerFeature.object
                    
                    Dim oSmartItem As IJSmartItem
                    Dim oSmartOccurrence As DEFINITIONHELPERSINTFLib.IJSmartOccurrence
                    
                    Set oSmartOccurrence = oCornerFeature.object
                    
                    If Not oSmartOccurrence.Item = "" Then
                        Set oSmartItem = oSmartOccurrence.SmartItemObject
                    End If
                    
                    Dim oCFName As String
                    oCFName = ""
                    
                    If Not oSmartItem Is Nothing Then
                        oCFName = oSmartItem.Name
                    End If

                    If oCFName = "LongScallopWithSeam" Then
                        ForceUpdateSmartItem oFeature
                    End If
            End Select
        End If
    Next
        
Exit Sub

ErrHandler:
  Err.Raise LogError(Err, MODULE, sMETHOD).Number

End Sub

'This method is to return the system for a given object. It depends on the boolean bRoot.
'If Selected object is a profile part/plate part AND If bRoot is true, the root system will be returned.
'    If bRoot is false, leaf system will be returned.
'If Selected object is a Profile/Plate leaf system AND if bRoot is true, the root system will be returned.
'   If bRoot is False, the same system(leaf system) will be returned.
'If Selected object is a Root system, always a Root system will be selected no matter what the bRoot is set.

Public Function GetSystemForObject(oPartOrSystem As Object, bRoot As Boolean) As Object
 
Dim oSystem As IJSystem
Dim oDesignChild As IJDesignChild
Dim structOp As StructOperation
Dim oHelper As New StructDetailHelper

If (TypeOf oPartOrSystem Is IJPlate And Not TypeOf oPartOrSystem Is IJPlatePart) Or _
   (TypeOf oPartOrSystem Is IJStiffener And Not TypeOf oPartOrSystem Is IJStiffenerPart) Then ' Is a plate system or a  stiffener system
    
    Dim oPlate As IJPlate
    Dim oProfile As IJProfile
    
    If TypeOf oPartOrSystem Is IJPlate Then
        Set oPlate = oPartOrSystem
        structOp = oPlate.OperationHistory
    Else
        Set oProfile = oPartOrSystem
        structOp = oProfile.OperationHistory
    End If

    If bRoot And (structOp And SplitOperation) Then
        Set oDesignChild = oPartOrSystem
        Set oSystem = oDesignChild.GetParent
    Else
        Set oSystem = oPartOrSystem
    End If

ElseIf TypeOf oPartOrSystem Is IJPlatePart Or TypeOf oPartOrSystem Is IJStiffenerPart Then ' Is a plate or stiffener part
    Dim recurseFlag As Boolean
    recurseFlag = False
    If bRoot Then
        recurseFlag = True
    End If
    
    If recurseFlag = True Then
        oHelper.IsPartDerivedFromSystem oPartOrSystem, oSystem, True
    Else
        oHelper.IsPartDerivedFromSystem oPartOrSystem, oSystem, False
    End If

    ' Not shown: Use IsPartDerivedFromSystem using variable for recursive option
End If

Set GetSystemForObject = oSystem

Exit Function
ErrorHandler:
   Err.Raise LogError(Err, "Helper.bas", "GetSystemForObject").Number
End Function

'This method determines if a clip is valid or not in the below described test case

'Consider a case where the StiffThroughPlate AC is split into multiple ACs by splitting the penetrated plate with seams(or plates)
'In such case, there will be multiple features(as many ACs) that would form a single slot. Till now, the code allows to place as
'many Clips as many no.of slots. TR 180319 calls for restriction of creation of clips/collars in such a way that: inspite of AC being
'split into multiple ACs, and multiple slots being created, only one set of Clip/Collar should be created as a child of the penetrated
'part with which it forms lap PCs.
'Ther could be a max no.of 2 penetrated plate parts with which the collar could create lap PCs. So, the collar should be made as a
'child of the AC of which connected object1 is a penetrated part with which the Clip/Collar forms a lap PC and also the penetrated
'part should be touching base plate of the stiffener.

'So, the logic is:
'IF it is a primary clip, it should only be created if
    'a) it's parent AC.ConnectedObject1 is towards web right side of the profile and the most nearest part to the base plate.
'If it is a Secondary Clip, it should only be created if
    'a) it's parent AC.ConnectedObject1 is towards web left side of the profile and the most nearest part to the base plate.


Public Sub IsClipValid(ByRef pMD As IJDMemberDescription, ByRef bIsValid As Boolean, sCollarOrder As String)
    On Error GoTo ErrorHandler

        Dim oAssyConn As New StructDetailObjects.AssemblyConn
        Set oAssyConn.object = pMD.CAO
        
        Dim oProfile As StructDetailObjects.ProfilePart
        Dim oStiffObject As Object
        
        If oAssyConn.ConnectedObject2Type = SDOBJECT_STIFFENER Then
            Set oStiffObject = oAssyConn.ConnectedObject2
        Else
            'Default to True for now until we can create similar logic for penetrating plate
            bIsValid = True
            Exit Sub
        End If
                
        Dim oRootStiffenerSystem As Object
        Set oRootStiffenerSystem = GetSystemForObject(oStiffObject, True)
        
        Dim oStiffPlate As IJPlate
        
        Dim oProfileAttr As GSCADSDPartSupport.IJProfilePartSupport
        Set oProfileAttr = New GSCADSDPartSupport.ProfilePartSupport
        
        Dim oSDPartsupport As GSCADSDPartSupport.IJPartSupport
        Set oSDPartsupport = oProfileAttr
        
        'get the stiffened plate
        Set oSDPartsupport.Part = oAssyConn.ConnectedObject2
        Set oStiffPlate = oProfileAttr.StiffenedPlate
                       
        'get the penetrated plate part
        Dim oPenetratedPart As Object
        Set oPenetratedPart = oAssyConn.ConnectedObject1
        
        Dim oPeneSystem As New StructDetailObjects.PlatePart
        Set oPeneSystem.object = oPenetratedPart
        
    
        'project the oPenetration location to profile
        Dim oPenetrationLoc As IJDPosition
        Dim oTopo As New GSCADStructGeomUtilities.TopologyLocate
        Dim oProfileAttributes As GSCADCreateModifyUtilities.IJProfileAttributes
        Set oProfileAttributes = New GSCADCreateModifyUtilities.ProfileUtils
        Dim oShipGeomOps As GSCADShipGeomOps.SGOModelBodyUtilities
        Set oShipGeomOps = New GSCADShipGeomOps.SGOModelBodyUtilities
    Dim oWireBodyUtils As GSCADShipGeomOps.SGOWireBodyUtilities
    Set oWireBodyUtils = New GSCADShipGeomOps.SGOWireBodyUtilities
        
    'get the primary orientation and secondary orientation of the profile
        Dim pPriOrient As IJDVector
        Dim pSecOrient As IJDVector
        
        Set pPriOrient = Nothing
        Set pSecOrient = Nothing
        
        Set oPenetrationLoc = oAssyConn.PenetrationGlobalShipLocation
        oProfileAttributes.GetProfileOrientation oRootStiffenerSystem, oPenetrationLoc, pSecOrient, pPriOrient
        
    
        'project the penetration location onto profile
        Dim oPointOnProfile As IJDPosition
        Dim dDist As Double
        oShipGeomOps.GetClosestPointOnBody oStiffObject, oPenetrationLoc, oPointOnProfile, dDist
        
        'offset oPenetraionPos to 20mm in Profile Primary orientation
        pPriOrient.Length = 0.005
        
        Dim oOffSetPos As IJDPosition
        Set oOffSetPos = oPointOnProfile.Offset(pPriOrient)
                
        'offset the above point in direction of stiffener secondary orientation incase of primary clip and negative of secondary orientation
        'incase of secondary clip
        If sCollarOrder = "Secondary" Then
            pSecOrient.Length = -pSecOrient.Length
        End If
        
        pSecOrient.Length = 0.005
        Set oOffSetPos = oOffSetPos.Offset(pSecOrient)
        
        Dim oLeafPenePart As IJDModelBody
        Set oLeafPenePart = oPeneSystem.ParentSystem
                        
    Dim oPoint As IJDPosition
    Dim oTangent As IJDVector
    Dim oLandingCurve As Object
    Dim oPointOnLeafPenPlate As IJDPosition
        
    ' Get the Profile Landing Curve
    oProfileAttributes.GetLandingCurveFromProfile oRootStiffenerSystem, oLandingCurve
        
    oWireBodyUtils.GetClosestPointOnWire oLandingCurve, oOffSetPos, oPoint, oTangent
    oTopo.NearestProjectedPointOnSurface oLeafPenePart, oOffSetPos, oTangent, oPointOnLeafPenPlate
    
    If oPointOnLeafPenPlate Is Nothing Then
        Dim oOppTangent As IJDVector
        Set oOppTangent = New dVector
        
        ' Setting the Reverse direction vector of the  Tangent vector
        oOppTangent.Set -oTangent.x, -oTangent.y, -oTangent.z
        
        ' Projecting the point on the reverse dirction of th landing curve
        oTopo.NearestProjectedPointOnSurface oLeafPenePart, oOffSetPos, oOppTangent, oPointOnLeafPenPlate
        
        If Not oPointOnLeafPenPlate Is Nothing Then
                bIsValid = True
            Else
                bIsValid = False
            End If
    ElseIf Not oPointOnLeafPenPlate Is Nothing Then
        bIsValid = True
        End If
        
CleanUp:

    Set oAssyConn = Nothing
    Set oProfile = Nothing
    Set oStiffObject = Nothing
    Set oRootStiffenerSystem = Nothing
    Set oStiffPlate = Nothing
    Set oProfileAttr = Nothing
    Set oProfileAttributes = Nothing
    Set oSDPartsupport = Nothing
    Set oPenetratedPart = Nothing
    Set oTopo = Nothing
    Set oShipGeomOps = Nothing
    Set oOffSetPos = Nothing
    Set oPenetrationLoc = Nothing
    Set oLeafPenePart = Nothing
    Set oPointOnLeafPenPlate = Nothing
    Set oPointOnProfile = Nothing
    Set oWireBodyUtils = Nothing

    Exit Sub
ErrorHandler:
   Err.Raise LogError(Err, "Helper.bas", "GetSystemForObject").Number
End Sub
'********************************************************************
' Routine: CheckPartClassExist
' Description:  Check if the given PartClass exist in Catalog
'
'********************************************************************
Public Function CheckPartClassExist(sPartClassName As String) As Boolean
    Const sMETHOD = "CheckPartClassExist"
    On Error GoTo ErrorHandler
    
    Dim oCatalogQuery As IJSRDQuery
    Dim oSmartQuery As IJSmartQuery
    Dim oPartClass As IJSmartClass
    
    Set oCatalogQuery = New SRDQuery
    Set oSmartQuery = oCatalogQuery
    
    ' Query for SmartClass... Check if exist
    On Error Resume Next
    Set oPartClass = oSmartQuery.GetClassByName(sPartClassName)
    On Error GoTo ErrorHandler
    
    If Not oPartClass Is Nothing Then
        CheckPartClassExist = True
    End If
    
    Set oPartClass = Nothing
    Set oSmartQuery = Nothing
    Set oCatalogQuery = Nothing
    
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD).Number
    
End Function

'********************************************************************************************************
'METHOD: GetMemberPortCrossSectionCode
'
'DESCRIPTION: This method retrieves the Port cross section code based on the operator ID of the port.
'
'********************************************************************************************************
Public Sub GetMemberPortCrossSectionCode(ByVal oMemberPort As Object, _
                            ByRef ePortCrossSectionCode As IMSProfileEntity.JXSEC_CODE)
    
    On Error GoTo ErrorHandler
    
    Const METHOD_NAME = "GetMemberPortCrossSectionCode"
    
    ePortCrossSectionCode = JXSEC_UNKNOWN
    
    Dim lngContexID As Long
    Dim lngOperatorID As Long
    Dim lngOperationID As Long
    
    Dim ePortType As JS_TOPOLOGY_PROXY_TYPE
    
    Dim oMemberFactory As SPSMembers.SPSMemberFactory
    Set oMemberFactory = New SPSMembers.SPSMemberFactory

    Dim oMemberConnectionServices As SPSMembers.ISPSMemberConnectionServices
    Set oMemberConnectionServices = oMemberFactory.CreateConnectionServices

    oMemberConnectionServices.GetStructPortInfo oMemberPort, ePortType, _
                                        lngContexID, lngOperationID, lngOperatorID

    ' Check if ActiveEntity OID and Operation ID are equal
    If lngOperationID = 2000 Then
        ePortCrossSectionCode = lngOperatorID
    End If
    
PROC_EXIT:
    Set oMemberConnectionServices = Nothing
    Set oMemberFactory = Nothing
    
    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD_NAME).Number
    GoTo PROC_EXIT
End Sub

'********************************************************************************************************
'METHOD: MemberPartPortType
'
'DESCRIPTION: This method returns the member port type in terms of enmProfilePortType by considering
'             the member cross section type and member port cross section code.
'
'********************************************************************************************************
Public Sub MemberPartPortType(ByVal strMemberXSectionType As String, _
                            ByVal ePortCrossSectionCode As JXSEC_CODE, _
                            ByRef eMemberPortType As enmProfilePortType)

    On Error GoTo ErrorHandler
    Const METHOD_NAME = "MemberPartPortType"
    
    Dim strError As String
    
    Select Case strMemberXSectionType
        Case "2L"
            If ePortCrossSectionCode = JXSEC_TOP Or _
                ePortCrossSectionCode = JXSEC_BOTTOM_FLANGE_LEFT Or _
                ePortCrossSectionCode = JXSEC_BOTTOM_FLANGE_RIGHT Then
                eMemberPortType = PROFILE_PORTTYPE_EDGE
            Else
                eMemberPortType = PROFILE_PORTTYPE_FACE
            End If
            
        Case "C", "MC"
            If ePortCrossSectionCode = JXSEC_TOP_FLANGE_RIGHT Or _
                ePortCrossSectionCode = JXSEC_TOP_FLANGE_RIGHT_BOTTOM_CORNER Or _
                ePortCrossSectionCode = JXSEC_BOTTOM_FLANGE_RIGHT Or _
                ePortCrossSectionCode = JXSEC_BOTTOM_FLANGE_RIGHT_TOP_CORNER Then
                eMemberPortType = PROFILE_PORTTYPE_EDGE
            Else
                eMemberPortType = PROFILE_PORTTYPE_FACE
            End If

        Case "W", "M", "HP"
            If ePortCrossSectionCode = JXSEC_TOP_FLANGE_LEFT Or _
                ePortCrossSectionCode = JXSEC_TOP_FLANGE_LEFT_BOTTOM_CORNER Or _
                ePortCrossSectionCode = JXSEC_TOP_FLANGE_RIGHT Or _
                ePortCrossSectionCode = JXSEC_TOP_FLANGE_RIGHT_BOTTOM_CORNER Or _
                ePortCrossSectionCode = JXSEC_BOTTOM_FLANGE_LEFT Or _
                ePortCrossSectionCode = JXSEC_BOTTOM_FLANGE_LEFT_TOP_CORNER Or _
                ePortCrossSectionCode = JXSEC_BOTTOM_FLANGE_RIGHT Or _
                ePortCrossSectionCode = JXSEC_BOTTOM_FLANGE_RIGHT_TOP_CORNER Then
                eMemberPortType = PROFILE_PORTTYPE_EDGE
            Else
                eMemberPortType = PROFILE_PORTTYPE_FACE
            End If
                        
        Case "HSSC", "PIPE"
            eMemberPortType = PROFILE_PORTTYPE_FACE
            
        Case "HSSR"
            eMemberPortType = PROFILE_PORTTYPE_FACE
        
        Case "L"
            If ePortCrossSectionCode = JXSEC_TOP Or _
                ePortCrossSectionCode = JXSEC_BOTTOM_FLANGE_RIGHT Then
                eMemberPortType = PROFILE_PORTTYPE_EDGE
            Else
                eMemberPortType = PROFILE_PORTTYPE_FACE
            End If
            
        Case "MT", "WT"
            If ePortCrossSectionCode = JXSEC_BOTTOM Or _
                ePortCrossSectionCode = JXSEC_TOP_FLANGE_LEFT Or _
                ePortCrossSectionCode = JXSEC_TOP_FLANGE_RIGHT Then
                eMemberPortType = PROFILE_PORTTYPE_EDGE
            Else
                eMemberPortType = PROFILE_PORTTYPE_FACE
            End If
            
        Case "S"
            If ePortCrossSectionCode = JXSEC_TOP_FLANGE_LEFT Or _
                ePortCrossSectionCode = JXSEC_TOP_FLANGE_LEFT_BOTTOM_CORNER Or _
                ePortCrossSectionCode = JXSEC_TOP_FLANGE_RIGHT Or _
                ePortCrossSectionCode = JXSEC_TOP_FLANGE_RIGHT_BOTTOM_CORNER Or _
                ePortCrossSectionCode = JXSEC_BOTTOM_FLANGE_LEFT Or _
                ePortCrossSectionCode = JXSEC_BOTTOM_FLANGE_LEFT_TOP_CORNER Or _
                ePortCrossSectionCode = JXSEC_BOTTOM_FLANGE_RIGHT Or _
                ePortCrossSectionCode = JXSEC_BOTTOM_FLANGE_RIGHT_TOP_CORNER Then
                eMemberPortType = PROFILE_PORTTYPE_EDGE
            Else
                eMemberPortType = PROFILE_PORTTYPE_FACE
            End If

        Case "ST"
            If ePortCrossSectionCode = JXSEC_BOTTOM Or _
                ePortCrossSectionCode = JXSEC_TOP_FLANGE_LEFT Or _
                ePortCrossSectionCode = JXSEC_TOP_FLANGE_RIGHT Then
                eMemberPortType = PROFILE_PORTTYPE_EDGE
            Else
                eMemberPortType = PROFILE_PORTTYPE_FACE
            End If

        Case Else
            strError = "Unknown member part cross section type (" & strMemberXSectionType & ")."
            GoTo ErrorHandler
    End Select

PROC_EXIT:

    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD_NAME).Number
    GoTo PROC_EXIT
End Sub

Public Sub GetMiterPlaneForMutualBound(ByVal pResManager As IUnknown, oAssemblyConn As StructDetailObjects.AssemblyConn, ByRef oMiterPlane As Plane3d)
    On Error GoTo ErrorHandler
    Const METHOD_NAME = "GetMiterPlaneForMutualBound"

    Dim oProfileAtt As GSCADCreateModifyUtilities.IJProfileAttributes
    Dim oTopo As New GSCADStructGeomUtilities.TopologyLocate
    Dim oHelper As New StructDetailObjects.Helper
    Dim oUtil As GSCADShipGeomOps.SGOModelBodyUtilities
    
    Set oProfileAtt = New GSCADCreateModifyUtilities.ProfileUtils
    Set oUtil = New GSCADShipGeomOps.SGOModelBodyUtilities
    
    Dim oBoundedProfile As New StructDetailObjects.ProfilePart
    Dim oBoundingProfile As New StructDetailObjects.ProfilePart
    Dim oIntersecPos As IJDPosition
    Dim oSecDir1 As IJDVector, oSecDir2 As IJDVector, oPriDir1 As IJDVector, oPriDir2 As IJDVector
    
    Set oBoundedProfile.object = oAssemblyConn.ConnectedObject1
    Set oBoundingProfile.object = oAssemblyConn.ConnectedObject2
    Set oIntersecPos = New DPosition
    Set oSecDir1 = New dVector: Set oPriDir2 = New dVector: Set oPriDir1 = New dVector: Set oSecDir2 = New dVector

    Dim oRootStiff1 As Object
    Dim oRootStiff2 As Object
    Set oRootStiff1 = oHelper.Object_RootParentSystem(oBoundedProfile.object)
    Set oRootStiff2 = oHelper.Object_RootParentSystem(oBoundingProfile.object)

    Dim bEndToEnd As Boolean
    Dim oVector1 As IJDVector
    Dim oVector2 As IJDVector
    
    Set oVector1 = New dVector
    Set oVector2 = New dVector
    
    'Get the intersecting position of the two profiles; primary orientations of the two profiles
    Set oIntersecPos = oTopo.FindIntersectionPoint(oRootStiff1, oRootStiff2)
    
    Dim oPort1 As IJPort
    Dim oPort2 As IJPort
     
    Dim oSurface1 As IJSurfaceBody
    Dim oSurface2 As IJSurfaceBody
    
    Set oPort1 = oAssemblyConn.Port1
    Set oPort2 = oAssemblyConn.Port2
    
    Set oSurface1 = oPort1.Geometry
    Set oSurface2 = oPort2.Geometry
    
    'Get Profile End vectors
    oSurface1.GetNormalFromPosition oIntersecPos, oVector1
    oSurface2.GetNormalFromPosition oIntersecPos, oVector2
        
               
    oVector1.Length = oBoundedProfile.Height
    oVector2.Length = oBoundingProfile.Height
    
    Dim oCrossVec As IJDVector: Set oCrossVec = New dVector
    Dim oResultant As IJDVector: Set oResultant = New dVector
    Dim oNormal As IJDVector: Set oNormal = New dVector
    
    'steps to obtain miter plane
    Set oCrossVec = oVector1.Cross(oVector2)
    Set oResultant = oVector1.Add(oVector2)
    Set oNormal = oCrossVec.Cross(oResultant)
    
    'create a plane
    Dim oFactory As IngrGeom3D.IPlanes3d
    Set oFactory = New IngrGeom3D.GeometryFactory
    Set oFactory = Nothing
    
    Dim pPOM As IJDPOM
    Set pPOM = Nothing
    
    Dim StructPlaneHelper As StructPlane.StructPlaneHelper
    Set StructPlaneHelper = New StructPlaneHelper
    Dim oStructPlane3d As Object
    Set oStructPlane3d = StructPlaneHelper.CreateStructPlane(pResManager, Nothing)
    Set oMiterPlane = oStructPlane3d
    If Not oMiterPlane Is Nothing Then
       oMiterPlane.SetRootPoint oIntersecPos.x, oIntersecPos.y, oIntersecPos.z
       oMiterPlane.SetNormal oNormal.x, oNormal.y, oNormal.z
       Set oStructPlane3d = Nothing
    End If
    Set StructPlaneHelper = Nothing
 
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD_NAME).Number
    
End Sub
'This method currently handles only stiffeners and returns true if the two stiffeners are colinear
Public Function AreConnectedObjectsColinear(oAC As StructDetailObjects.AssemblyConn) As Boolean
    On Error GoTo ErrorHandler
    Const METHOD_NAME = "AreConnectedObjectsColinear"
    AreConnectedObjectsColinear = False
    
    Dim oWireBody2 As IJWireBody, oWireBody1 As IJWireBody
    If oAC.ConnectedObject1Type = SDOBJECT_STIFFENER Then
        Set oWireBody1 = GetProfilePartLandingCurve(oAC.ConnectedObject1)
    End If
    If oAC.ConnectedObject2Type = SDOBJECT_STIFFENER Then
        Set oWireBody2 = GetProfilePartLandingCurve(oAC.ConnectedObject2)
    End If
    
    If oWireBody1 Is Nothing Or oWireBody2 Is Nothing Then
        Exit Function
    End If
    
    Dim oWireUtil As IJSGOWireBodyUtilities
    Set oWireUtil = New SGOWireBodyUtilities

    Dim oAxis1 As IJDVector, oAxis2 As IJDVector
    Set oAxis1 = New dVector: Set oAxis2 = New dVector
    Dim oClosestPoint1 As IJDPosition, oClosestPoint2 As IJDPosition
    Set oClosestPoint1 = New DPosition: Set oClosestPoint2 = New DPosition
    
    oWireUtil.GetClosestPointOnWire oWireBody1, oAC.BoundGlobalShipLocation, oClosestPoint1, oAxis1
    oWireUtil.GetClosestPointOnWire oWireBody2, oAC.BoundGlobalShipLocation, oClosestPoint2, oAxis2
  
    If Equal(Abs(oAxis1.Dot(oAxis2)), 1) Then
        AreConnectedObjectsColinear = True
    End If
    
    
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD_NAME).Number
End Function

' ********************************************************************************
' Method:
'   LandingCurve
' Description:
'   Gets the landingcurve for the given Stiffener/ER/Beam
' ********************************************************************************
' This method copied from EndCutRules\Common.bas
' EndCuts should be made to use the MarineLibraryCommon.bas file, which is in a more public location
Private Function GetProfilePartLandingCurve(oProfilePart As Object) As IJWireBody
    On Error GoTo ErrorHandler

    Dim oStructDetailHelper As StructDetailHelper
    Dim oTopoLocate As IJTopologyLocate
    Dim oLandingCrv As IJWireBody

    Set oStructDetailHelper = New StructDetailHelper
    Set oTopoLocate = New TopologyLocate

    'Based on whether the part is derived from system or not,
    'we shall get the landing curve differently for performance

    Dim oIJStructGraph As IJStructGraph
    Set oIJStructGraph = oProfilePart

    If Not oIJStructGraph Is Nothing Then
        Dim oParentSystem As IJSystem
        oStructDetailHelper.IsPartDerivedFromSystem oIJStructGraph, oParentSystem

        'Derived from system?
        If Not oParentSystem Is Nothing Then
            Set oLandingCrv = oTopoLocate.GetProfileParentWireBody(oProfilePart)
        Else
            'Not derived from system
            Dim oPartSupport As GSCADSDPartSupport.IJPartSupport
            Dim oProfilePartSupport As GSCADSDPartSupport.IJProfilePartSupport

            Set oProfilePartSupport = New GSCADSDPartSupport.ProfilePartSupport
            Set oPartSupport = oProfilePartSupport
            Set oPartSupport.Part = oProfilePart

            ' Get the curve (default direction is SideUnspecified, which gives
            ' a curve through the load point)
            Dim oThicknessDir As IJDVector
            Dim bThicknessCentered As Boolean

            oProfilePartSupport.GetProfilePartLandingCurve oLandingCrv, _
                                                           oThicknessDir, _
                                                           bThicknessCentered

            Set oProfilePartSupport = Nothing
            Set oPartSupport = Nothing
            Set oThicknessDir = Nothing
        End If

        Set oParentSystem = Nothing
    End If

    Set GetProfilePartLandingCurve = oLandingCrv

    Set oLandingCrv = Nothing
    Set oStructDetailHelper = Nothing
    Set oTopoLocate = Nothing

    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "GetProfilePartLandingCurve").Number
End Function


'********************************************************************************************************
'METHOD: GetFacePortsOfMemberOverlappingWithPlate
'
'DESCRIPTION: This helper method gets all the LATERAL_FACE_PORTS of the member part.
'             Loops through each of these ports to check for an overlap with plate port.
'             Returns the collection of ports which are having a valid overlap.
'
'********************************************************************************************************
Public Function GetFacePortsOfMemberOverlappingWithPlate( _
                        ByVal oPlatePort As Object, _
                        ByVal oMemberPart As Object) As Collection
    
    Const METHOD_NAME = "GetFacePortsOfMemberOverlappingWithPlate"
    
    On Error GoTo ErrorHandler
    
    Dim oOverlapFacePortColl As Collection
    Set oOverlapFacePortColl = New Collection
    
    If oPlatePort Is Nothing _
        Or oMemberPart Is Nothing Then
        GoTo PROC_EXIT
    End If
    
    Dim oPlatePortSurface As IJSurfaceBody
    If TypeOf oPlatePort Is IJPort Then
        If TypeOf oPlatePort.Geometry Is IJSurfaceBody Then
            Set oPlatePortSurface = oPlatePort.Geometry
        End If
    End If

    Dim oStructGraphConnectable As IJStructGraphConnectable
   
    If TypeOf oMemberPart Is ISPSMemberPartPrismatic Then
        If TypeOf oMemberPart Is IJStructGraphConnectable Then
            Set oStructGraphConnectable = oMemberPart
        End If
    End If
        
    Dim oEnumPorts As IJElements
    If Not oStructGraphConnectable Is Nothing Then
        oStructGraphConnectable.enumPortsInGraphByTopologyFilter _
                                        oEnumPorts, _
                                        JS_TOPOLOGY_FILTER_SOLID_LATERAL_LFACES, _
                                        CurrentGeometry, _
                                        vbNull
    End If
    
    Dim oSurfaceBodyUtils As IJSGOSurfaceBodyUtilities
    Set oSurfaceBodyUtils = New GSCADShipGeomOps.SGOSurfaceBodyUtilities
    
    If Not oEnumPorts Is Nothing Then
        If oEnumPorts.Count > 0 Then
            Dim lPortCount As Long
            For lPortCount = 1 To oEnumPorts.Count
                
                Dim oMemberPort As IJPort
                Dim oPort As Object
                
                If TypeOf oEnumPorts.Item(lPortCount) Is IJPort Then
                    Set oPort = oEnumPorts.Item(lPortCount)
                    
                    If Not oPort Is Nothing Then
                        Set oMemberPort = oPort
                    End If
                End If
                
                Dim oMemberPortSurface As IJSurfaceBody
                If Not oMemberPort Is Nothing Then
                    If TypeOf oMemberPort.Geometry Is IJSurfaceBody Then
                        Set oMemberPortSurface = oMemberPort.Geometry
                    End If
                End If
                
                Dim oWireBody As IJWireBody
                                                                  
                If Not oSurfaceBodyUtils Is Nothing Then
                    If Not oPlatePortSurface Is Nothing And _
                        Not oMemberPortSurface Is Nothing Then
                        
                        On Error Resume Next
                        oSurfaceBodyUtils.FindSheetOverlapWithTolerance oPlatePortSurface, _
                                                            oMemberPortSurface, _
                                                            SHEET_OVERLAP_TOLERANCE, _
                                                            True, _
                                                            OVERLAP_ONLY_FACES_WITH_OPPOSITE_NORMALS, _
                                                            Nothing, _
                                                            Nothing, _
                                                            oWireBody
                        On Error GoTo ErrorHandler
                        
                        If Not oWireBody Is Nothing Then
                            oOverlapFacePortColl.Add oMemberPort
                        End If
                    End If
                End If
                
            Next
        End If
    End If
        
PROC_EXIT:
    Set oSurfaceBodyUtils = Nothing
    Set oPlatePortSurface = Nothing
    Set oMemberPortSurface = Nothing
    Set oPort = Nothing
    Set oStructGraphConnectable = Nothing
    Set oMemberPort = Nothing
    Set oEnumPorts = Nothing
    Set oWireBody = Nothing
    
    Set GetFacePortsOfMemberOverlappingWithPlate = oOverlapFacePortColl
    Set oOverlapFacePortColl = Nothing

Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD_NAME).Number
    GoTo PROC_EXIT
    
End Function


Public Function ConstructPC_MemberAndPlateFace(pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, strStartClass As String) As Object

  On Error GoTo ErrorHandler

    Dim sError As String
    Dim sMETHOD As String
    
    sError = "Constructing physical connection."
    sMETHOD = "CMConstructPC"
    
    ' Get Wrapper Class
    Dim pAssyConn As StructDetailObjects.AssemblyConn
    Set pAssyConn = New StructDetailObjects.AssemblyConn
    
    ' Initialize wrapper class and get the 2 ports
    sError = "Setting Assembly Connection Inputs"
    Set pAssyConn.object = pMemberDescription.CAO
    
    sError = "Getting Assembly Connection Ports"
    
    Dim oPort1 As IJPort
    Dim oPort2 As IJPort
    Set oPort1 = pAssyConn.Port1
    Set oPort2 = pAssyConn.Port2
    
    ' Get the Assembly connection, since it is the parent of the PC
    Dim pSystemParent As IJSystemChild
    sError = "Setting system parent to Member Description Custom Assembly"
    Set pSystemParent = pMemberDescription.CAO
       
    'Create PC
    Dim oPhysicalConnection As New PhysicalConn
    sError = "Creating Physical Connection"
    
    Dim oSLPort1 As IJPort
    Dim oSLPort2 As IJPort
    
    Dim oStructDetailObjectHelper As New StructDetailObjects.Helper
    Dim oFacePortColl As New Collection
    
    If TypeOf oPort1.Connectable Is IJPlate Then
        Set oSLPort1 = oStructDetailObjectHelper.GetEquivalentLastPort(oPort1)
        
        Set oFacePortColl = GetFacePortsOfMemberOverlappingWithPlate(oPort1, oPort2.Connectable)
                
        'Get Member Port
        If oFacePortColl.Count > 0 Then
            Set oSLPort2 = oStructDetailObjectHelper.GetEquivalentLastPort(oFacePortColl.Item(1))
        Else
            Set oSLPort2 = oStructDetailObjectHelper.GetEquivalentLastPort(oPort2)
        End If
    Else ' if port1 is Member
        Set oSLPort2 = oStructDetailObjectHelper.GetEquivalentLastPort(oPort2)
                
        Set oFacePortColl = GetFacePortsOfMemberOverlappingWithPlate(oPort2, oPort1.Connectable)
        
        'Get Member Port
        If oFacePortColl.Count > 0 Then
            Set oSLPort1 = oStructDetailObjectHelper.GetEquivalentLastPort(oFacePortColl.Item(1))
        Else
            Set oSLPort1 = oStructDetailObjectHelper.GetEquivalentLastPort(oPort1)
        End If
    End If
    
    Set oStructDetailObjectHelper = Nothing
    Set oPort1 = Nothing
    Set oPort2 = Nothing
    
    'Create PC
    Call oPhysicalConnection.Create(pResourceManager, oSLPort1, oSLPort2, _
                                    strStartClass, pSystemParent, ConnectionStandard)
        
    Set ConstructPC_MemberAndPlateFace = oPhysicalConnection.object

Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD).Number

End Function

'*************************************************************************
'Function
'   IsACisWithinRangeOfChange
'
'Abstract
'   Checks if given AC is completely within the range-of-change box
'
'Inputs
'   oAssyConn As Object
'   ogRngofChgBox As GBox
'
'Return
'   Value is 'True' if the AC is completely within the range-of-change box, otherwise 'False'.

'Exceptions
'
'***************************************************************************
Public Function IsACisWithinRangeOfChange(oAssyConn As Object, gRngofChgBox As GBox) As Boolean
  On Error GoTo ErrorHandler

    Dim sMETHOD As String
    sMETHOD = "IsACisWithinRangeOfChange"
    
    IsACisWithinRangeOfChange = False 'Initialize

    'Compare Rangebox of AC and Rangebox of the corner feature
    Dim oRange As IJRangeAlias
    Dim gACRngBox As GBox
    Set oRange = oAssyConn
    gACRngBox = oRange.GetRange()
    
    'If all of following six conditions are true, AC is within the given range-of-change(GBox)
    If GreaterThan(gACRngBox.m_low.x, gRngofChgBox.m_low.x) And LessThan(gACRngBox.m_high.x, gRngofChgBox.m_high.x) Then
        If GreaterThan(gACRngBox.m_low.y, gRngofChgBox.m_low.y) And LessThan(gACRngBox.m_high.y, gRngofChgBox.m_high.y) Then
            If GreaterThan(gACRngBox.m_low.z, gRngofChgBox.m_low.z) And LessThan(gACRngBox.m_high.z, gRngofChgBox.m_high.z) Then
                IsACisWithinRangeOfChange = True
            End If
        End If
    End If
    
CleanUp:
    Set oRange = Nothing
    
Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD).Number
End Function

'*************************************************************************
'Function
'   CheckIfBothPlanesAreSame
'
'Abstract
'   Checks if given both planes are same by checking Plane Root point and Normal
'
'Inputs
'   oPlane1 As IJPlane
'   oPlane2 As IJPlane
'
'Return
'   Value is 'True' if the both the planes are in same and if not value is 'Flase'
'
'Exceptions
'
'***************************************************************************
Public Function CheckIfBothPlanesAreSame(oPlane1 As IJPlane, oPlane2 As IJPlane) As Boolean

  On Error GoTo ErrorHandler

    Dim sMETHOD As String
    sMETHOD = "CheckIfBothPlanesAreSame"
    
    CheckIfBothPlanesAreSame = False

    Dim dx1 As Double
    Dim dy1 As Double
    Dim dz1 As Double
    
    Dim dX2 As Double
    Dim dY2 As Double
    Dim dZ2 As Double
    
    oPlane1.GetRootPoint dx1, dy1, dz1
    oPlane2.GetRootPoint dX2, dY2, dZ2
    
    'Check if both Old and new miter planes are same.
    If Equal(dx1, dX2) And Equal(dy1, dY2) And Equal(dz1, dZ2) Then
    
        oPlane1.GetNormal dx1, dy1, dz1
        oPlane2.GetNormal dX2, dY2, dZ2
        
        If Equal(dx1, dX2) And Equal(dy1, dY2) And Equal(dz1, dZ2) Then
            CheckIfBothPlanesAreSame = True
        End If
    End If
   
Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD).Number
    
End Function

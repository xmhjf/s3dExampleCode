Attribute VB_Name = "FullyParametricFun"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003-07, Intergraph Corporation. All rights reserved.
'
'   FullyParametricFun.bas
'   Author:
'   Creation Date:
'   Description:
'       TODO - fill in header description information
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'16 Feb 2003    SSP      Commented out End Preparations Checks for Crating Nozzle.
'6 June 2003    Symbol Team Added Function CreateRetrieveDynamicEquipmentNozzle and ReportUnanticipatedError3
                                            'Commented unused variable of type IJDConnection
'If Symbol uses only basGeom3d.bas file, for error report call function  "ReportUnanticipatedError2"
'If Symbol uses only FullyParametricFun.bas file, for error report call function  "ReportUnanticipatedError3"
'If Symbol uses both basGeom3d.bas and FullyParametricFun.bas files, for error report call function  "ReportUnanticipatedError2"
'
'   09.Jul.2003  SymbolTeam(India)   Copyright Information, Header  is added.
'   22.Aug.2003  SymbolTeam(India)   TR 46728 Change Port Index of a nozzle to 15 used to corrupt Equipment.
'                                    Modified port index logic. Added new function 'CreateRetrieveDynamicNozzleForEquipment'
'                                    which can be used for any port index.Modified "CreateRetrieveDynamicEquipmentNozzle" in same way.
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'  1.NOV.2007     RRK  CR-123952  Modfied CreateRetrieveDynamicNozzle, CreateRetrieveDynamicEquipmentNozzle, CreateRetrieveDynamicNozzleForEquipment
'                                 functions to give cptOffset value zero when port termination class is bolted using a new optional boolean
'                                 parameter as argument
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Const MODULE = "FullyParametricFun"


Public Function GetCatalogDBResourceManager() As IUnknown
    Const METHOD = "GetCatalogDBResourceManager"
    On Error GoTo ErrorHandler
    Dim oMidctx As IJMiddleContext
    Set oMidctx = New GSCADMiddleContextProj.GSCADMiddleContext
    Set GetCatalogDBResourceManager = oMidctx.GetResourceManager("Catalog")
    Exit Function

ErrorHandler:
    ReportUnanticipatedError3 MODULE, METHOD
End Function


' Proposed 'CreateRetrieveDynamicNozzle' function can be used for creation of fully dynamic nozzle,
' and for retrieval of following parameters: PipeDiam, FlangeThick, FlangeDiam, sptoffset, depth
' Notes:-
'   i. You need to assign nozzle length to nozzle object after calling this function,
'   ii. Prefix 'Def' in 'DefNPD' denotes 'Default NPD', the same is true with other parameters.
'   iii. While retriving the values for nozzles, the value for input parameter "PortIndex" doesn't matter, as it
'       sends the values for all nozzles in a single function call.
Public Function CreateRetrieveDynamicNozzle(PortIndex As Long, DefNPD As Double, DefNPDUnitType As String, _
                    DefEndPreparation As Long, DefScheduleThickness As Long, DefEndStandard As Long, _
                    DefPressureRating As Long, DefFlowDirection As DistribFlow, Npd As Double, _
                    NPDUnitType As String, EndPreparation As Long, ScheduleThickness As Long, _
                    EndStandard As Long, PressureRating As Long, FlowDirection As DistribFlow, _
                    Id As String, objOutputColl As Object, m_oCodeListMetadata As IJDCodeListMetaData, _
                    CreateFlag As Boolean, ByRef lPipeDiam() As Double, ByRef lFlangeThick() As Double, _
                    ByRef lFlangeDiam() As Double, ByRef lsptoffset() As Double, ByRef lDepth() As Double, _
                    Optional lBoltedPartDimIncludesFlgFaceProj = True) _
                    As GSCADNozzleEntities.IJDNozzle
    
    Const METHOD = "CreateRetrieveDynamicNozzle:"
    On Error GoTo ErrorHandler

    Static NumNozzles As Integer
    Dim iCount As Integer
    Dim TerminationClass    As Long
    Dim TerminationSubClass As Long
    Dim SchedulePractice    As Long
    Dim EndPractice         As Long
    Dim RatingPractice      As Long
    
    Dim PortStatus          As DistribPortStatus
    Dim oLogicalDistPort    As GSCADNozzleEntities.IJLogicalDistPort
    Dim oDistribPort        As GSCADNozzleEntities.IJDistribPort
    
    Dim oPipePort           As GSCADNozzleEntities.IJDPipePort
    Dim oNozzleFactory      As GSCADNozzleEntities.NozzleFactory
    Dim oNozzle             As GSCADNozzleEntities.IJDNozzle
    
    'For the time being supporting a maxi of 12no of nozzles
    Static pipeDia(1 To 12) As Double
    Static FlangeTk(1 To 12) As Double
    Static flangeDia(1 To 12) As Double
    Static sptOffset(1 To 12) As Double
    Static depth(1 To 12) As Double
    'If the function is called from Physical the create the Nozzle and store the Values viz., Pipe dia, flange dia
    ' in a buffer
    
    If CreateFlag = True Then
        If PortIndex > NumNozzles Then NumNozzles = PortIndex
        PortStatus = DistribPortStatus_BASE
        Set oNozzleFactory = New GSCADNozzleEntities.NozzleFactory
        'Dim oCatalogConnection  As IJDConnection

        Dim oResourceManager As IUnknown
        Set oResourceManager = GetCatalogDBResourceManager
        Dim oDir As New AutoMath.DVector
        Dim oPlacePoint As New AutoMath.DPosition
        
        If Npd <= 0 Then
            Npd = DefNPD
        End If
    
        If EndPreparation <= 0 Then
            EndPreparation = DefEndPreparation
        End If
    
        If ScheduleThickness <= 0 Then
            ScheduleThickness = DefScheduleThickness
        End If
    
        If EndStandard <= 0 Then
            EndStandard = DefEndStandard
        End If
    
        If PressureRating <= 0 Then
            PressureRating = DefPressureRating
        End If
    
        If FlowDirection <= 0 Then
            FlowDirection = DefFlowDirection
        End If
           
'        If EndPreparation > 1 And EndPreparation < 300 Then                    ('16 Feb 2003    SSP      Commented out End Preparations Checks for Crating Nozzle.)
'            If PressureRating > 60 Then
'                PressureRating = 35
'            End If
'        ElseIf EndPreparation >= 300 And EndPreparation < 400 Then
'            If ScheduleThickness < 5 Then
'                ScheduleThickness = 100
'            End If
'        ElseIf EndPreparation >= 400 Then
'           PressureRating = 110
'        End If
'
        If NPDUnitType = "" Then
            NPDUnitType = DefNPDUnitType
        End If
                  
        TerminationSubClass = m_oCodeListMetadata.ParentValueID("EndPreparation", EndPreparation)
        TerminationClass = m_oCodeListMetadata.ParentValueID("TerminationSubClass", TerminationSubClass)
        SchedulePractice = m_oCodeListMetadata.ParentValueID("ScheduleThickness", ScheduleThickness)
        EndPractice = m_oCodeListMetadata.ParentValueID("EndStandard", EndStandard)
        RatingPractice = m_oCodeListMetadata.ParentValueID("PressureRating", PressureRating)
        
    ''uncomment to make test with text inputs
        Set oNozzle = oNozzleFactory.CreatePipeNozzle(PortIndex, Npd, NPDUnitType, _
                                                EndPreparation, ScheduleThickness, EndStandard, _
                                                PressureRating, FlowDirection, PortStatus, Id, _
                                                TerminationClass, TerminationSubClass, SchedulePractice, _
                                                5, EndPractice, RatingPractice, False, objOutputColl.ResourceManager, oResourceManager)
        Set oPipePort = oNozzle

        '   Return nozzle and other values
        pipeDia(PortIndex) = oPipePort.PipingOutsideDiameter      'Static
        FlangeTk(PortIndex) = oPipePort.FlangeOrHubThickness
        flangeDia(PortIndex) = oPipePort.FlangeOrHubOutsideDiameter
        depth(PortIndex) = oPipePort.SeatingOrGrooveOrSocketDepth
        sptOffset(PortIndex) = oPipePort.FlangeProjectionOrSocketOffset
        ' Case when part dimension includes face projection and port termination class is 'Bolted'
        If lBoltedPartDimIncludesFlgFaceProj = True And _
                                        oPipePort.TerminationClass = 5 Then
            sptOffset(PortIndex) = 0
        End If
        
        'Return Values
        lPipeDiam(PortIndex) = pipeDia(PortIndex)
        lFlangeThick(PortIndex) = FlangeTk(PortIndex)
        lFlangeDiam(PortIndex) = flangeDia(PortIndex)
        lDepth(PortIndex) = depth(PortIndex)
        lsptoffset(PortIndex) = sptOffset(PortIndex)
        
        Set CreateRetrieveDynamicNozzle = oNozzle
        Set oNozzle = Nothing
    Else        ' if the function is called from places other than Physical, only return the values.
            For iCount = 1 To NumNozzles
                lPipeDiam(iCount) = pipeDia(iCount)
                lFlangeThick(iCount) = FlangeTk(iCount)
                lFlangeDiam(iCount) = flangeDia(iCount)
                lsptoffset(iCount) = sptOffset(iCount)
                lDepth(iCount) = depth(iCount)
            Next iCount
    End If
    Exit Function
ErrorHandler:
  ReportUnanticipatedError3 MODULE, METHOD
  
End Function


' Proposed 'CreateRetrieveDynamicNozzle' function can be used for creation of fully dynamic nozzle,
' and for retrieval of following parameters: PipeDiam, FlangeThick, FlangeDiam, sptoffset, depth
' Notes:-
'   i. You need to assign nozzle length to nozzle object after calling this function,
'   ii. Prefix 'Def' in 'DefNPD' denotes 'Default NPD', the same is true with other parameters.
'   iii. While retriving the values for nozzles, the value for input parameter "PortIndex" doesn't matter, as it
'       sends the values for all nozzles in a single function call.
'
'   Important Notes:
'   i. If CreateFlag is false this function will return nozzle properties as per the
'      sequence of their creation.
'   ii. Array size which is being passed into this function (for eg. lFlangeDiam() etc.)
'       should be same as the number of nozzles created using this function.
Public Function CreateRetrieveDynamicEquipmentNozzle(ByRef oPartFclt As PartFacelets.IJDPart, PortIndex As Long, Npd As Double, _
                    NPDUnitType As String, EndPreparation As Long, ScheduleThickness As Long, _
                    EndStandard As Long, PressureRating As Long, FlowDirection As DistribFlow, _
                    Id As String, objOutputColl As Object, m_oCodeListMetadata As IJDCodeListMetaData, _
                    CreateFlag As Boolean, ByRef lPipeDiam() As Double, ByRef lFlangeThick() As Double, _
                    ByRef lFlangeDiam() As Double, ByRef lsptoffset() As Double, ByRef lDepth() As Double, _
                    Optional lBoltedPartDimIncludesFlgFaceProj = True) _
                    As GSCADNozzleEntities.IJDNozzle
                    
    
    Const METHOD = "CreateRetrieveDynamicEquipmentNozzle:"
    Const E_FAIL = -2147467259

    On Error GoTo ErrorHandler

    Const MAX_NO_NOZ = 30
    Static nozzleCount As Integer
'    Static NumNozzles As Integer
    Dim iCount As Integer
    Dim TerminationClass    As Long
    Dim TerminationSubClass As Long
    Dim SchedulePractice    As Long
    Dim EndPractice         As Long
    Dim RatingPractice      As Long
    
    Dim PortStatus          As DistribPortStatus
    Dim oLogicalDistPort    As GSCADNozzleEntities.IJLogicalDistPort
    Dim oDistribPort        As GSCADNozzleEntities.IJDistribPort
    
    Dim oPipePort           As GSCADNozzleEntities.IJDPipePort
    Dim oNozzleFactory      As GSCADNozzleEntities.NozzleFactory
    Dim oNozzle             As GSCADNozzleEntities.IJDNozzle
    
    'For the time being supporting a maximum of 30 no of nozzles
    Static pipeDia(1 To MAX_NO_NOZ) As Double
    Static FlangeTk(1 To MAX_NO_NOZ) As Double
    Static flangeDia(1 To MAX_NO_NOZ) As Double
    Static sptOffset(1 To MAX_NO_NOZ) As Double
    Static depth(1 To MAX_NO_NOZ) As Double

    'If the function is called from Physical the create the Nozzle and store the Values viz., Pipe dia, flange dia
    ' in a buffer
    
    If CreateFlag = True Then
'        If PortIndex > NumNozzles Then NumNozzles = PortIndex
        PortStatus = DistribPortStatus_BASE
        Set oNozzleFactory = New GSCADNozzleEntities.NozzleFactory
        'Dim oCatalogConnection  As IJDConnection

        Dim oResourceManager As IUnknown
        Set oResourceManager = GetCatalogDBResourceManager
        Dim oDir As New AutoMath.DVector
        Dim oPlacePoint As New AutoMath.DPosition
        
'        If Npd <= 0 Then
'            Npd = DefNPD
'        End If
'
'        If EndPreparation <= 0 Then
'            EndPreparation = DefEndPreparation
'        End If
'
'        If ScheduleThickness <= 0 Then
'            ScheduleThickness = DefScheduleThickness
'        End If
'
'        If EndStandard <= 0 Then
'            EndStandard = DefEndStandard
'        End If
'
'        If PressureRating <= 0 Then
'            PressureRating = DefPressureRating
'        End If
'
'        If FlowDirection <= 0 Then
'            FlowDirection = DefFlowDirection
'        End If
'
'        If EndPreparation > 1 And EndPreparation < 300 Then                    ('16 Feb 2003    SSP      Commented out End Preparations Checks for Crating Nozzle.)
'            If PressureRating > 60 Then
'                PressureRating = 35
'            End If
'        ElseIf EndPreparation >= 300 And EndPreparation < 400 Then
'            If ScheduleThickness < 5 Then
'                ScheduleThickness = 100
'            End If
'        ElseIf EndPreparation >= 400 Then
'           PressureRating = 110
'        End If
'

''''        If NPDUnitType = "1" Or "" Then
''''            RetrieveNozzleParameters PortIndex, oPartFclt, objOutputColl, PartNozzleNpdUnit
''''            NPDUnitType = PartNozzleNpdUnit
''''        End If
        If nozzleCount = UBound(lPipeDiam) Then
            nozzleCount = 0
        End If

        Dim PartNozzleNpdUnit As String
        If NPDUnitType <> "mm" Or NPDUnitType <> "in" Or NPDUnitType <> "A" Or NPDUnitType = "1" Or NPDUnitType = "" Then
            RetrieveNozzleParameters nozzleCount + 1, oPartFclt, objOutputColl, PartNozzleNpdUnit
            NPDUnitType = PartNozzleNpdUnit
        End If
                  
        TerminationSubClass = m_oCodeListMetadata.ParentValueID("EndPreparation", EndPreparation)
        TerminationClass = m_oCodeListMetadata.ParentValueID("TerminationSubClass", TerminationSubClass)
        SchedulePractice = m_oCodeListMetadata.ParentValueID("ScheduleThickness", ScheduleThickness)
        EndPractice = m_oCodeListMetadata.ParentValueID("EndStandard", EndStandard)
        RatingPractice = m_oCodeListMetadata.ParentValueID("PressureRating", PressureRating)
        
    ''uncomment to make test with text inputs
        Set oNozzle = oNozzleFactory.CreatePipeNozzle(PortIndex, Npd, NPDUnitType, _
                                                EndPreparation, ScheduleThickness, EndStandard, _
                                                PressureRating, FlowDirection, PortStatus, Id, _
                                                TerminationClass, TerminationSubClass, SchedulePractice, _
                                                5, EndPractice, RatingPractice, False, objOutputColl.ResourceManager, oResourceManager)
        Set oPipePort = oNozzle

    'Return nozzle and other values
        nozzleCount = nozzleCount + 1
        If nozzleCount > MAX_NO_NOZ Then Err.Raise E_FAIL
        pipeDia(nozzleCount) = oPipePort.PipingOutsideDiameter      'Static
        FlangeTk(nozzleCount) = oPipePort.FlangeOrHubThickness
        flangeDia(nozzleCount) = oPipePort.FlangeOrHubOutsideDiameter
        depth(nozzleCount) = oPipePort.SeatingOrGrooveOrSocketDepth
        sptOffset(nozzleCount) = oPipePort.FlangeProjectionOrSocketOffset
        ' Case when part dimension includes face projection and port termination class is 'Bolted'
        If lBoltedPartDimIncludesFlgFaceProj = True And _
                                        oPipePort.EndPreparation = 5 Then
            sptOffset(nozzleCount) = 0
        End If
        'Return Values
        lPipeDiam(nozzleCount) = pipeDia(nozzleCount)
        lFlangeThick(nozzleCount) = FlangeTk(nozzleCount)
        lFlangeDiam(nozzleCount) = flangeDia(nozzleCount)
        lDepth(nozzleCount) = depth(nozzleCount)
        lsptoffset(nozzleCount) = sptOffset(nozzleCount)
        
        Set CreateRetrieveDynamicEquipmentNozzle = oNozzle
        Set oNozzle = Nothing
    Else
    ' if the function is called from places other than Physical, only return the values.
        For iCount = 1 To UBound(lPipeDiam)
            lPipeDiam(iCount) = pipeDia(iCount)
            lFlangeThick(iCount) = FlangeTk(iCount)
            lFlangeDiam(iCount) = flangeDia(iCount)
            lsptoffset(iCount) = sptOffset(iCount)
            lDepth(iCount) = depth(iCount)
        Next iCount
    End If
    Exit Function
ErrorHandler:
  ReportUnanticipatedError3 MODULE, METHOD
  
End Function

' Proposed 'CreateRetrieveDynamicNozzleForEquipment' function can be used for creation of fully dynamic nozzle for Equipment,
' and for retrieval of following parameters: PipeDiam, FlangeThick, FlangeDiam, sptoffset, depth
' Notes:-
'   i. You need to assign nozzle length to nozzle object after calling this function,
'   ii. Prefix 'Def' in 'DefNPD' denotes 'Default NPD', the same is true with other parameters.
'   iii. While retriving the values for nozzles, the value for input parameter "PortIndex" doesn't matter, as it
'       sends the values for all nozzles in a single function call.
'
'   Important Notes:
'   i. If CreateFlag is false this function will return nozzle properties as per the
'      sequence of their creation.
'   ii. Array size which is being passed into this function (for eg. lFlangeDiam() etc.)
'       should be same as the number of nozzles created using this function.
Public Function CreateRetrieveDynamicNozzleForEquipment(PortIndex As Long, DefNPD As Double, DefNPDUnitType As String, _
                    DefEndPreparation As Long, DefScheduleThickness As Long, DefEndStandard As Long, _
                    DefPressureRating As Long, DefFlowDirection As DistribFlow, Npd As Double, _
                    NPDUnitType As String, EndPreparation As Long, ScheduleThickness As Long, _
                    EndStandard As Long, PressureRating As Long, FlowDirection As DistribFlow, _
                    Id As String, objOutputColl As Object, m_oCodeListMetadata As IJDCodeListMetaData, _
                    CreateFlag As Boolean, ByRef lPipeDiam() As Double, ByRef lFlangeThick() As Double, _
                    ByRef lFlangeDiam() As Double, ByRef lsptoffset() As Double, ByRef lDepth() As Double, _
                    Optional lBoltedPartDimIncludesFlgFaceProj = True) _
                    As GSCADNozzleEntities.IJDNozzle
    
    Const METHOD = "CreateRetrieveDynamicNozzleForEquipment:"
    Const E_FAIL = -2147467259
    On Error GoTo ErrorHandler
    Const MAX_NO_NOZ = 30
    Static nozzleCount As Integer

    Dim iCount As Integer
    Dim TerminationClass    As Long
    Dim TerminationSubClass As Long
    Dim SchedulePractice    As Long
    Dim EndPractice         As Long
    Dim RatingPractice      As Long
    
    Dim PortStatus          As DistribPortStatus
    Dim oLogicalDistPort    As GSCADNozzleEntities.IJLogicalDistPort
    Dim oDistribPort        As GSCADNozzleEntities.IJDistribPort
    
    Dim oPipePort           As GSCADNozzleEntities.IJDPipePort
    Dim oNozzleFactory      As GSCADNozzleEntities.NozzleFactory
    Dim oNozzle             As GSCADNozzleEntities.IJDNozzle
    
    'For the time being supporting a maximum of 30 no of nozzles
    Static pipeDia(1 To MAX_NO_NOZ) As Double
    Static FlangeTk(1 To MAX_NO_NOZ) As Double
    Static flangeDia(1 To MAX_NO_NOZ) As Double
    Static sptOffset(1 To MAX_NO_NOZ) As Double
    Static depth(1 To MAX_NO_NOZ) As Double
    'If the function is called from Physical the create the Nozzle and store the Values viz., Pipe dia, flange dia
    ' in a buffer
    If CreateFlag = True Then
        PortStatus = DistribPortStatus_BASE
        Set oNozzleFactory = New GSCADNozzleEntities.NozzleFactory

        Dim oResourceManager As IUnknown
        Set oResourceManager = GetCatalogDBResourceManager
        Dim oDir As New AutoMath.DVector
        Dim oPlacePoint As New AutoMath.DPosition
        
        If Npd <= 0 Then
            Npd = DefNPD
        End If
    
        If EndPreparation <= 0 Then
            EndPreparation = DefEndPreparation
        End If
    
        If ScheduleThickness <= 0 Then
            ScheduleThickness = DefScheduleThickness
        End If
    
        If EndStandard <= 0 Then
            EndStandard = DefEndStandard
        End If
    
        If PressureRating <= 0 Then
            PressureRating = DefPressureRating
        End If
    
        If FlowDirection <= 0 Then
            FlowDirection = DefFlowDirection
        End If
           
        If NPDUnitType = "" Then
            NPDUnitType = DefNPDUnitType
        End If
                  
        TerminationSubClass = m_oCodeListMetadata.ParentValueID("EndPreparation", EndPreparation)
        TerminationClass = m_oCodeListMetadata.ParentValueID("TerminationSubClass", TerminationSubClass)
        SchedulePractice = m_oCodeListMetadata.ParentValueID("ScheduleThickness", ScheduleThickness)
        EndPractice = m_oCodeListMetadata.ParentValueID("EndStandard", EndStandard)
        RatingPractice = m_oCodeListMetadata.ParentValueID("PressureRating", PressureRating)
        
    ''uncomment to make test with text inputs
        Set oNozzle = oNozzleFactory.CreatePipeNozzle(PortIndex, Npd, NPDUnitType, _
                                                EndPreparation, ScheduleThickness, EndStandard, _
                                                PressureRating, FlowDirection, PortStatus, Id, _
                                                TerminationClass, TerminationSubClass, SchedulePractice, _
                                                5, EndPractice, RatingPractice, False, objOutputColl.ResourceManager, oResourceManager)
        Set oPipePort = oNozzle

        'Return nozzle and other values
        If nozzleCount = UBound(lPipeDiam) Then
            nozzleCount = 0
        End If
        nozzleCount = nozzleCount + 1
        If nozzleCount > MAX_NO_NOZ Then Err.Raise E_FAIL
        pipeDia(nozzleCount) = oPipePort.PipingOutsideDiameter      'Static
        FlangeTk(nozzleCount) = oPipePort.FlangeOrHubThickness
        flangeDia(nozzleCount) = oPipePort.FlangeOrHubOutsideDiameter
        depth(nozzleCount) = oPipePort.SeatingOrGrooveOrSocketDepth
        sptOffset(nozzleCount) = oPipePort.FlangeProjectionOrSocketOffset
        ' Case when part dimension includes face projection and port termination class is 'Bolted'
        If lBoltedPartDimIncludesFlgFaceProj = True And _
                                        oPipePort.EndPreparation = 5 Then
            sptOffset(nozzleCount) = 0
        End If
        'Return Values
        lPipeDiam(nozzleCount) = pipeDia(nozzleCount)
        lFlangeThick(nozzleCount) = FlangeTk(nozzleCount)
        lFlangeDiam(nozzleCount) = flangeDia(nozzleCount)
        lDepth(nozzleCount) = depth(nozzleCount)
        lsptoffset(nozzleCount) = sptOffset(nozzleCount)
        Set CreateRetrieveDynamicNozzleForEquipment = oNozzle
        Set oNozzle = Nothing
    Else
    ' if the function is called from places other than Physical, only return the values.
        For iCount = 1 To UBound(lPipeDiam)
            lPipeDiam(iCount) = pipeDia(iCount)
            lFlangeThick(iCount) = FlangeTk(iCount)
            lFlangeDiam(iCount) = flangeDia(iCount)
            lsptoffset(iCount) = sptOffset(iCount)
            lDepth(iCount) = depth(iCount)
        Next iCount
    End If
    Exit Function

ErrorHandler:
  ReportUnanticipatedError3 MODULE, METHOD
  
End Function

'Used to report truly unexpected errors - a last resort response
'As errors actually occur and are reported the calling code should then
'be modified to in anticipate and handle them and not call this sub
Public Sub ReportUnanticipatedError3(InModule As String, InMethod As String, Optional errnumber As Long, Optional Context As String, Optional ErrDescription As String)
    Const E_FAIL = -2147467259
    Err.Raise E_FAIL
End Sub

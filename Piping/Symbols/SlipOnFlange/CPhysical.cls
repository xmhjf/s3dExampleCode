VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPhysical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003-2007, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          NN
'   Creation Date:  Thursday, Nov 23 2000
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   09.Aug.2004     MS     CR 60092 The Set back distance for Slipon flange is
'   the sum of the thickness of the weld inside the flange and the company practice gap.
'   The port2 location is positioned at Company practice gap distance from port1.
'   And the movement of the pipe by the weld size will be taken care by Route.
'   09.Aug.2004     MS     CR 62284 .As the dimensions and placement of these
'   cylinders are governed by the generic data ( i.e. socket offset, socket depth, flange or hub
'   thickness) for both the nozzles that compose the slip-on flange, we would not be able to align
'   the two faces of the cylinders when we use the company practice gap value to position nozzle 2.
'   This because the computation of the nozzle geometry is internal and uses all the end data that
'   is persisted on the nozzle, while the company practice gap that is used to position nozzle 2 is
'   external and cannot be a defining property of nozzle 2. Thus this misalignment occurs.
'   To solve this in the symbol we have created the nozzles without the graphics using the CreatePipeNozzle function
'   (passing the parameter bLtWtGraphics value as True). And added 2 additional outputs in the symbol for creating the
'   slip-on flange geometry (one cylinder for the flange portion and another for the hub portion). Now the composite flange looks correct.
'   We position the second nozzle at out company practice gap value from the nozzle 1. Route needs to move the port2 location by the inner weld size.
'   31.Aug.2004     MS     DI-62953 changed declaration of pipeport from IJDPipePort to IJCatalogPipePort as it has been changed for V5 and above.
'   03.Nov.2004     MS     TR-66125    For the hub cylinder, the socket depth (length through hub) dimension is used to get the thickness of the hub ,
'                                                       whether the user has defined a value for hub thickness or not.
'                                                       If the user does not define a value for the hub OD, the symbol goes by the formula
'                                                       hub OD = pipe OD + arbitrary factor * ( flange OD - pipe OD).
'                                                       Hence the symbol 'should' expect 1) the flange thickness and flange OD to be defined in the bolted end generic data;
'                                                       and 2) the socket depth should be defined in the female end generic data.
'   28.Nov.2004     MS     TR-67584    The symbol needs to use Hubthickness from the female generic data and not the Socketdepth to construct the hub cylinder.
'                                                       In only V6 through TR66125, the sliponflange symbol got updated to use the socketdepth for the hub cylinder.
'                                                       This needs to be reverted back.
'   25.Mar.2005     MS     TR-73996 Provided validation code to check that the cylinders never fail.
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'  1.NOV.2007      RRK  CR-123952 Updated the symbol to support part data basis values of 15 and 20
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit


Private Const MODULE = "Physical:" 'Used for error messages

Private Sub Class_Initialize()

'''

End Sub


Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim cptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    Dim CenterPos       As New AutoMath.DPosition

    CenterPos.Set 0, 0, 0

    Dim iOutput     As Double
    Dim ObjInsulatedBody As Object
    
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parInsulationThickness = arrayOfInputs(2)
    
    iOutput = 0
    
    Dim opipecomp As IJDPipeComponent
    Set opipecomp = oPartFclt
    
    Dim lPartDataBasis As Long
    lPartDataBasis = opipecomp.PartDataBasis
    Dim blnLengthThrHubInclFaceProj As Boolean
    
    'Assigning the boolean parameter(which will be passed to RetrieveParameters function)
    'an appropriate value based on part data basis
    If lPartDataBasis <= 1 Or lPartDataBasis = 20 Then 'Length-through-hub dimension includes flange face projection
        blnLengthThrHubInclFaceProj = True
    ElseIf lPartDataBasis = 15 Then 'Length-through-hub dimension excludes flange face projection
        blnLengthThrHubInclFaceProj = False
    End If

' Place Nozzle 1

    'Dim oPipePort       As GSCADNozzleEntities.IJDPipePort
    Dim oPipePort  As GSCADNozzleEntities.IJCatalogPipePort
    Dim oCollection As IJDCollection
    Dim pPortIndex As Integer
    Dim Npd As Double
    Dim NPDUnitType As String
    Dim EndPreparation As Long
    Dim ScheduleThickness As Long
    Dim EndStandard As Long
    Dim PressureRating As Long
    Dim FlowDirection As DistribFlow
    Dim PortStatus          As DistribPortStatus
    Dim Id As String
    Dim TerminationClass As Long
    Dim TerminationSubClass As Long
    Dim SchedulePractice As Long
    Dim EndPractice As Long
    Dim RatingPractice As Long
    
    Set oCollection = oPartFclt.GetNozzles()
    
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector


    Dim oResourceManager As IUnknown
    Set oResourceManager = GetCatalogDBResourceManager
    
    For pPortIndex = 1 To oCollection.Size
        Set oPipePort = oCollection.Item(pPortIndex)
            If oPipePort.PortIndex = 1 Then
                    Npd = oPipePort.Npd
                    NPDUnitType = oPipePort.NPDUnitType
                    EndPreparation = oPipePort.EndPreparation
                    ScheduleThickness = oPipePort.ScheduleThickness
                    EndStandard = oPipePort.EndStandard
                    PressureRating = oPipePort.PressureRating
                    FlowDirection = oPipePort.FlowDirection
                    PortStatus = DistribPortStatus_BASE
                    Id = oPipePort.Id
                    TerminationClass = oPipePort.TerminationClass
                    TerminationSubClass = oPipePort.TerminationSubClass
                    SchedulePractice = oPipePort.SchedulePractice
                    EndPractice = oPipePort.EndPractice
                    RatingPractice = oPipePort.RatingPractice
                
                Exit For
            End If
    Next pPortIndex
       
    Dim oNozzle             As GSCADNozzleEntities.IJDNozzle
    Dim oLogicalDistPort    As GSCADNozzleEntities.IJLogicalDistPort
    Dim oDistribPort        As GSCADNozzleEntities.IJDistribPort
    
  
    Dim oNozzleFactory      As GSCADNozzleEntities.NozzleFactory
    Set oNozzleFactory = New GSCADNozzleEntities.NozzleFactory
    
    Set oNozzle = oNozzleFactory.CreatePipeNozzle(1, Npd, NPDUnitType, _
                                                EndPreparation, ScheduleThickness, EndStandard, _
                                                PressureRating, FlowDirection, PortStatus, Id, _
                                                TerminationClass, TerminationSubClass, SchedulePractice, _
                                                5, EndPractice, RatingPractice, True, m_OutputColl.ResourceManager, oResourceManager)
       
       
    Set oLogicalDistPort = oNozzle
    Set oDistribPort = oNozzle
    oNozzle.Length = 0
    
' Direction of the Nozzle
    oDir.Set -1, 0, 0
    oDistribPort.SetDirectionVector oDir

' Position of the nozzle should be the connect point of the nozzle
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, _
                                            cptOffset, depth, blnLengthThrHubInclFaceProj
    oPlacePoint.Set -cptOffset + depth, 0, 0
    oLogicalDistPort.SetCenterLocation oPlacePoint

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oNozzle
    Set oNozzle = Nothing
    
' Place Nozzle 2
    For pPortIndex = 1 To oCollection.Size
        Set oPipePort = oCollection.Item(pPortIndex)
            If oPipePort.PortIndex = 2 Then
                Npd = oPipePort.Npd
                    NPDUnitType = oPipePort.NPDUnitType
                    EndPreparation = oPipePort.EndPreparation
                    ScheduleThickness = oPipePort.ScheduleThickness
                    EndStandard = oPipePort.EndStandard
                    PressureRating = oPipePort.PressureRating
                    FlowDirection = oPipePort.FlowDirection
                    PortStatus = DistribPortStatus_BASE
                    Id = oPipePort.Id
                    TerminationClass = oPipePort.TerminationClass
                    TerminationSubClass = oPipePort.TerminationSubClass
                    SchedulePractice = oPipePort.SchedulePractice
                    EndPractice = oPipePort.EndPractice
                    RatingPractice = oPipePort.RatingPractice
                
                Exit For
            End If
    Next pPortIndex
        
    Set oNozzle = oNozzleFactory.CreatePipeNozzle(2, Npd, NPDUnitType, _
                                                EndPreparation, ScheduleThickness, EndStandard, _
                                                PressureRating, FlowDirection, PortStatus, Id, _
                                                TerminationClass, TerminationSubClass, SchedulePractice, _
                                                5, EndPractice, RatingPractice, True, m_OutputColl.ResourceManager, oResourceManager)
       
    
    Set oLogicalDistPort = oNozzle
    Set oDistribPort = oNozzle
    oNozzle.Length = 0
    
' Direction of the Nozzle
    oDir.Set 1, 0, 0
    oDistribPort.SetDirectionVector oDir

    
''   The Set back distance for Slipon flange is the sum of the thickness of the weld inside the flange and the company practice gap.
''   The port2 location is positioned at Company practice gap distance from port1.
''   And the movement of the pipe by the weld size will be taken care by Route.
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, _
                                            cptOffset, depth, blnLengthThrHubInclFaceProj
    Dim dCompanyPracticeGap As Double
    Dim dPort2location As Double
    
    opipecomp.GetCompanyPracticeGap dCompanyPracticeGap
    
    dPort2location = dCompanyPracticeGap
    
' If company practice gap value is not defined then the port2 would be located at cptoffset.
    If dCompanyPracticeGap = 0 Then dPort2location = cptOffset
    
    oPlacePoint.Set dPort2location, 0, 0
    oLogicalDistPort.SetCenterLocation oPlacePoint

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oNozzle
    Set oNozzle = Nothing
    Set oPipePort = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    
' Place graphics for the cylinders
'Place Cylinder 1
    Dim stpoint As New AutoMath.DPosition
    Dim enpoint As New AutoMath.DPosition
    Dim cyl1 As Object
    Dim cyl2 As Object
    Dim dHubOD As Double
    Dim dHubCylDiameter As Double
    Dim dHubCylLength As Double
    Dim dFlangeProjection As Double
    
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, _
                            cptOffset, depth, blnLengthThrHubInclFaceProj
    dFlangeProjection = cptOffset
    'hub OD = pipe OD + arbitrary factor * ( flange OD - pipe OD).
    dHubOD = pipeDiam + 0.3 * (flangeDiam - pipeDiam)
    
    'Made a check such that the cylinder never fails.
    If flangeThick = 0 Then flangeThick = 0.0001
    If flangeDiam = 0 Then flangeDiam = pipeDiam
    
    stpoint.Set 0, 0, 0
    enpoint.Set flangeThick, 0, 0
    
    Set cyl1 = PlaceCylinder(m_OutputColl, stpoint, enpoint, flangeDiam, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), cyl1
    Set cyl1 = Nothing
      
'Place Cylinder 2
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, _
                                        cptOffset, depth, blnLengthThrHubInclFaceProj
    dHubCylLength = flangeThick
    If flangeThick < 0.0001 Then
        dHubCylLength = depth
            If depth < 0.0001 Then
                dHubCylLength = 0.0001
            End If
    End If
    
    stpoint.Set 0, 0, 0
    enpoint.Set dHubCylLength + dFlangeProjection, 0, 0
    
    
    If flangeDiam = 0 Then
        dHubCylDiameter = dHubOD
    Else
        dHubCylDiameter = flangeDiam
    End If
   
    
    'Made a check such that the cylinder never fails.
    If enpoint.x = 0 Then enpoint.x = 0.0001
    If dHubCylDiameter = 0 Then dHubCylDiameter = pipeDiam
    
    Set cyl2 = PlaceCylinder(m_OutputColl, stpoint, enpoint, dHubCylDiameter, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), cyl2
    Set cyl2 = Nothing
    
    Set stpoint = Nothing
    Set enpoint = Nothing
    
  
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
  
    
End Sub


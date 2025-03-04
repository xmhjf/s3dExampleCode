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
'   Copyright (c) 2004-08, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          MS
'   Creation Date:  Friday, Aug 06 2004
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   The lap joint flange symbol (SP3DLapJointFlange.CLapJointFlange) supports both with stub end and flared pipe.
'   For a Non-flared the stub geometry also exists and the StubLength and LapThickness are obtained from the lapjointflange partclass.
'   The nozzle locations are at either ends of the lap and stub.For a Flared pipe the LapThickness value would be provided
'   from the FlaredPipeRule by Route to the symbol. For Flared pipe, based on the isFlaredPipe symbol attribute value which
'   route sends to the symbol and the lapthickness provided the locations of both ports for the lap joint flange symbol
'   are coincident with the end of the flared pipe and only the flange portion of geometry is generated.
'   The lap joint flanges that are referenced by the PDS piping specs are defined with the stub end as a reportable piping commodity.
'   By definition, the lap joint flanges should appear in a part class named LapJointFlange, and the stub ends should be included in a part class named StubEnd.
'
'   26.DEC.2004    MS     TR68880 Removed msgboxes
'   24.Mar.2006    kkk    TR-95468  lap joint flange symbol needs to be modified for socket welded end cases
'  08.SEP.2006     KKC    DI-95670  Replace names with initials in all revision history sheets and symbols
'  1.NOV.2007      RRK    CR-123952 Updated the symbol to support part data basis values of 15 and 20
'  26.Mar.2008     RRK    CR-83045:Implemented part data basis value-'Lap joint flange, backing flange only'(931)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:"    'Used for error messages

Private Sub Class_Initialize()

'''

End Sub


Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)

    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt As PartFacelets.IJDPart
    Dim pipeDiam1 As Double
    Dim flangeThick1 As Double
    Dim flangeDiam1 As Double
    Dim cptOffset As Double
    Dim depth As Double
    Dim cptOffset2 As Double
    Dim depth2 As Double

    Dim pipeDiam2 As Double
    Dim flangeThick2 As Double
    Dim flangeDiam2 As Double

    Dim iOutput As Double
    Dim objFlange As Object
    Dim objStubEnd As Object


    Dim parStubLength As Double
    Dim parLapThickness As Double
    Dim parIsPipeFlared As Double
    Dim parInsulationThickness As Double
    Dim parFacetoFace As Double
    Dim parHubDiameter As Double

    Dim dLapThickness As Double
    Dim dPort1Location As Double
    Dim dPort2Location As Double
    
    ' Inputs
    Set oPartFclt = arrayOfInputs(1)
    '    parStubLength = arrayOfInputs(2)
    '    parLapThickness = arrayOfInputs(3)
    parIsPipeFlared = arrayOfInputs(4)
    parInsulationThickness = arrayOfInputs(5)
    '    parFacetoFace = arrayOfInputs(6)
    '    parHubDiameter = arrayOfInputs(7)

    Dim oPipeComponent As IJDPipeComponent
    Set oPipeComponent = oPartFclt
    Dim lPartDataBasis As Integer
    lPartDataBasis = oPipeComponent.PartDataBasis

    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam1, flangeThick1, flangeDiam1, cptOffset, depth, False
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick2, flangeDiam2, cptOffset2, depth2

    Select Case lPartDataBasis
    Case Is <= 1, 20, 930
        Set oPartFclt = arrayOfInputs(1)
        parStubLength = arrayOfInputs(2)
        parLapThickness = arrayOfInputs(3)
        dLapThickness = parLapThickness

        dPort1Location = depth
        dPort2Location = parStubLength - depth2 + cptOffset2

    Case 15
        Set oPartFclt = arrayOfInputs(1)
        parStubLength = arrayOfInputs(2)
        parLapThickness = arrayOfInputs(3)
        dLapThickness = parLapThickness

        dPort1Location = -cptOffset + depth
        dPort2Location = parStubLength - depth2 + cptOffset2

    Case 931
        parFacetoFace = arrayOfInputs(6)
        parHubDiameter = arrayOfInputs(7)

        'Retrieving the lap thickness value from bolted end generic data
        Call oPipeComponent.GetLapThickness(1, dLapThickness)
        dPort1Location = 0
        dPort2Location = 0
    End Select

    iOutput = 0

    ' Create Output 1 for Lap Flange Cylinder
    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition

    If lPartDataBasis = 931 Then
        stPoint.Set dLapThickness, 0, 0
    Else
        stPoint.Set 0, 0, 0
    End If

    enPoint.Set flangeThick1 + dLapThickness, 0, 0

    Set objFlange = PlaceCylinder(m_OutputColl, stPoint, enPoint, flangeDiam1, True)

    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objFlange
    Set objFlange = Nothing

    ' Create Output 2 for Stub End Cylinder
    If lPartDataBasis = 931 Then
        Dim ObjConicalHub As Object

        Dim dHubDiameter As Double
        dHubDiameter = parHubDiameter

        'Considering hub diameter as optional parameter, assigning a value equal to 1.2 times the pipe diameter of non-bolted end
        If CmpDblEqual(dHubDiameter, 0) Then dHubDiameter = 1.2 * pipeDiam2

        'Check to see that hub diameter with the above assumption  doesn't exceed flange diameter
        If CmpDblGreaterthanOrEqualTo(dHubDiameter, flangeDiam1) Then dHubDiameter = (flangeDiam1 + pipeDiam2) / 2

        stPoint.Set dLapThickness + flangeThick1, 0, 0
        enPoint.Set dLapThickness + parFacetoFace, 0, 0

        Set ObjConicalHub = PlaceCone(m_OutputColl, stPoint, enPoint, dHubDiameter / 2, pipeDiam2 / 2, True)

        ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjConicalHub
        Set ObjConicalHub = Nothing
    Else
        
        If parIsPipeFlared <> 0 Then parStubLength = 0.000001    ''For a Flared pipe the stublength would be zero.
        
        stPoint.Set 0, 0, 0
        enPoint.Set parStubLength, 0, 0

        Set objStubEnd = PlaceCylinder(m_OutputColl, stPoint, enPoint, pipeDiam2, True)

        ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objStubEnd
        Set objStubEnd = Nothing
    End If

    Set stPoint = Nothing
    Set enPoint = Nothing

    ' Place Nozzle 1(With Light weight Graphics, hence no geometry part of the nozzle would be generated)
    Dim oDir As AutoMath.DVector
    Dim objNozzle As GSCADNozzleEntities.IJDNozzle
    Dim oPlacePoint As New AutoMath.DPosition

    Set oDir = New AutoMath.DVector
    oDir.Set -1, 0, 0
    oPlacePoint.Set dPort1Location, 0, 0
    Set objNozzle = CreateNozzleJustaCircle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)

    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing

    ' Place Nozzle 2(With Light weight Graphics, hence no geometry part of the nozzle would be generated)
    'Check for flared pipe where port 2 coincides with port 1
    If parIsPipeFlared <> 0 Then dPort2Location = dPort1Location

    oPlacePoint.Set dPort2Location, 0, 0
    oDir.Set 1, 0, 0

    Set objNozzle = CreateNozzleJustaCircle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)

    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    Set oDir = Nothing
    Set oPlacePoint = Nothing
    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
End Sub

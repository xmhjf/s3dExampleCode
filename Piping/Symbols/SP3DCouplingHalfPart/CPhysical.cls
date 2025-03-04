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
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:         RRK
'   Creation Date:  Thursday, Sept 27 2007
'   Description: This symbol is prepared for male and female adapters(or couplings).
'   Source:www.dixonvalve.com/fgal/lit_pdf/2006/outside_air_applications_2006.pdf
'          www.ivesequipment.com/customer/iveqco/principals/alfalaval/Product%20Data%20Leaflets/Fluid%20Handling/Sanitary%20&%20Biopharm%20fittings.pdf
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   09-Oct-07               CR-126633-Provide male & female coupling symbols
'                                     for sanitary piping & hose couplings
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "CPhysical:" 'Used for error messages

Private Sub Class_Initialize()
'''
End Sub


Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam1        As Double
    Dim flangeThick1     As Double
    Dim sptOffset1       As Double
    Dim flangeDiam1      As Double
    Dim depth1           As Double

    Dim pipeDiam2        As Double
    Dim flangeThick2     As Double
    Dim sptOffset2       As Double
    Dim flangeDiam2      As Double
    Dim depth2           As Double
    
    Dim iOutput     As Double
    
    Dim parFacetoFace As Double
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoFace = arrayOfInputs(2)
    parInsulationThickness = arrayOfInputs(3)
    
    iOutput = 0
    
    Dim oPipeComponent As IJDPipeComponent
    Set oPipeComponent = oPartFclt
    Dim lPartDataBasis As Integer
    lPartDataBasis = oPipeComponent.PartDataBasis
    
    Select Case lPartDataBasis
    Case Is <= 1, 422, 423 '422: Male Coupling, specified by face to face
                           '423: Female Coupling, specified by face to face
        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam1, flangeThick1, flangeDiam1, sptOffset1, depth1
        RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick2, flangeDiam2, sptOffset2, depth2
    
    ' Place Nozzle 1(Output 1)
        Dim oPlacePoint As AutoMath.DPosition
        Dim oDir        As AutoMath.DVector
        Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
        
        Set oPlacePoint = New AutoMath.DPosition
        Set oDir = New AutoMath.DVector
        
        Dim dNozzleLength1 As Double
        
        Dim oPipePort1  As IJCatalogPipePort
        Dim oPipePort2  As IJCatalogPipePort
        Dim oCollection As IJDCollection
        
        Set oCollection = oPartFclt.GetNozzles()
        Set oPipePort1 = oCollection.Item(1)
        Set oPipePort2 = oCollection.Item(2)
        
        'Setting the nozzle lengths based on whether the end preparation belongs two male class or not
        If oPipePort1.TerminationClass = 10 Or oPipePort2.TerminationClass = 10 Then 'Termination Class 10 corresponds to 'Male'
            dNozzleLength1 = 0.7 * parFacetoFace
        Else
            dNozzleLength1 = parFacetoFace - flangeThick2
        End If
        
        If CmpDblLessThan(dNozzleLength1, flangeThick1) Then dNozzleLength1 = flangeThick1
        
        oPlacePoint.Set -parFacetoFace / 2 - sptOffset1 + depth1, 0, 0
        oDir.Set -1, 0, 0
        
        Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, _
                                oPlacePoint, dNozzleLength1)
    ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing
    
    ' Place Nozzle 2 (Output 2)
    
        oPlacePoint.Set (parFacetoFace / 2) + sptOffset2 - depth2, 0, 0
        oDir.Set 1, 0, 0
        
        Dim dNozzleLength2 As Double
        dNozzleLength2 = parFacetoFace - dNozzleLength1
        If CmpDblLessThanOrEqualTo(dNozzleLength2, flangeThick2) Then dNozzleLength2 = flangeThick2 + 0.0001
        Set objNozzle = CreateNozzleWithLength(2, oPartFclt, m_OutputColl, oDir, _
                                oPlacePoint, dNozzleLength2)
                                
    ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing
        Set oPlacePoint = Nothing
        Set oDir = Nothing
        Set oPipeComponent = Nothing
        Set oPipePort1 = Nothing
        Set oPipePort2 = Nothing
    End Select
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
End Sub

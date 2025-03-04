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
'   Creation Date:  Monday 29, Jan 2007
'   Description:
'   This implements the physical aspect of the symbol created for Temporary Cone Strainer
'   to support multiple dimensional bases as per CR-113397
'   The symbol supports the following dimensional bases
'   1.)Part Data Basis Value 110:-Temporary Cone Strainer, specified by basket length inclusive of plate thickness
'   2.)Part Data Basis Value 115:-Temporary Cone Strainer, specified by basket length exclusive of plate thickness
'   3.)Part Data Basis Value 120:-Temporary Cone Strainer, specified by basket length exclusive of plate thickness
'           with identifier tag handle height measured from centerline

'   Change History:

'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)

    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double

    Dim iOutput     As Double
    Dim ObjPlate As Object
    Dim objCone As Object
    Dim ObjIdentifierTag As Object
    
    Dim parFacetoFace As Double
    Dim parConeLength As Double
    Dim parConeDiameter As Double
    Dim parIdentifierTagHeight As Double
    Dim parIdentifierTagHtFrmCenter As Double
    Dim parIdentifierTagWidth As Double
    Dim parInsulationThickness As Double

    Dim dOutsideDiam As Double
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoFace = arrayOfInputs(2)
    parConeLength = arrayOfInputs(3)
    parConeDiameter = arrayOfInputs(4)
'    parIdentifierTagHeight = arrayOfInputs(5)
'    parIdentifierTagHtFrmCenter = arrayOfInputs(6)
    parIdentifierTagWidth = arrayOfInputs(7)

    'Checking for the PartDataBasis
    Dim oPipeComponent As IJDPipeComponent
    Dim lPartDataBasis As Long
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis
    
    If lPartDataBasis <= 1 Or lPartDataBasis = 110 Or lPartDataBasis = 115 Then
        parIdentifierTagHeight = arrayOfInputs(5)
    ElseIf lPartDataBasis = 120 Then
        parIdentifierTagHtFrmCenter = arrayOfInputs(6)
    Else: GoTo ErrorLabel
    End If
    
    Dim oStPoint   As Object
    Dim oEnPoint   As Object
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition

'   Place Output 1 (Plate)
    
    oStPoint.Set -parFacetoFace / 2, 0, 0
    oEnPoint.Set parFacetoFace / 2, 0, 0
    
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    
    dOutsideDiam = flangeDiam
    Set ObjPlate = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dOutsideDiam, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjPlate
    Set ObjPlate = Nothing

'   Place Output 2(Cone)
    
    If (lPartDataBasis = 115 Or lPartDataBasis = 120) Then
    
        oStPoint.Set -parFacetoFace / 2, 0, 0
        oEnPoint.Set parFacetoFace / 2 + parConeLength, 0, 0
        
    ElseIf (lPartDataBasis <= 1 Or lPartDataBasis = 110) Then
    
        oStPoint.Set -parFacetoFace / 2, 0, 0
        oEnPoint.Set parConeLength - (parFacetoFace / 2), 0, 0
    End If
    
    Set objCone = PlaceCone(m_OutputColl, oStPoint, oEnPoint, parConeDiameter / 2, 0, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCone
    Set objCone = Nothing
    
'   Place Output 3(Identifier Tag)
    Dim dTagStPointY As Double
    dTagStPointY = Sqr((dOutsideDiam / 2) ^ 2 - (parIdentifierTagWidth / 2) ^ 2)
    
    If (lPartDataBasis <= 1 Or lPartDataBasis = 110 Or lPartDataBasis = 115) Then
                
        oStPoint.Set -parFacetoFace / 2, dTagStPointY, parIdentifierTagWidth / 2
        oEnPoint.Set parFacetoFace / 2, _
                    dTagStPointY + parIdentifierTagHeight, _
                    -parIdentifierTagWidth / 2
        
    ElseIf (lPartDataBasis = 120) Then
    
        oStPoint.Set -parFacetoFace / 2, dTagStPointY, parIdentifierTagWidth / 2
        oEnPoint.Set parFacetoFace / 2, parIdentifierTagHtFrmCenter, -parIdentifierTagWidth / 2
        
    End If
    
    Set ObjIdentifierTag = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjIdentifierTag
    Set objCone = Nothing
    
' Place Nozzle 1
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    Dim dNozzleLength As Double
    
   
    oPlacePoint.Set -parFacetoFace / 2 - sptOffset + depth, 0, 0
    oDir.Set -1, 0, 0
    
    Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing

' Place Nozzle 2
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth

    oPlacePoint.Set parFacetoFace / 2 + sptOffset - depth, 0, 0
    oDir.Set 1, 0, 0
    
    Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    Set oStPoint = Nothing
    Set oEnPoint = Nothing

    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub


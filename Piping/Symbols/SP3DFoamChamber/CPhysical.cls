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
'   Copyright (c) 2008, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:         VRK
'   Creation Date:  Monday, Feb 25 2008
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This symbol implements PartDataBasis -780 and 781
'   Foam Chamber, with default-780
'   Foam Chamber, with offset-781
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
' 25.Feb.2008     VRK  CR-136578  Enhance foam chamber symbol to provide more realistic dimensional basis
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
    Dim pipeDiam2        As Double
    Dim flangeThick2     As Double
    Dim sptOffset2       As Double
    Dim depth2           As Double
    
    Dim iOutput     As Double
    Dim ObjCylinBody As Object
    
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parBodyDiameter As Double
    Dim parHeight As Double
    Dim parOffset As Double
    
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
'    parFace1toCenter = arrayOfInputs(2)
'    parFace2toCenter = arrayOfInputs(3)
'    parBodyDiameter = arrayOfInputs(4)
'    parHeight = arrayOfInputs(5)
'    parOffset = arrayOfInputs(6)
    
    iOutput = 0
    
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, _
                                            flangeDiam, sptOffset, depth
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick2, _
                                            flangeDiam, sptOffset2, depth2

    Dim oStartPt As AutoMath.DPosition
    Dim oEndPt As AutoMath.DPosition
    Set oStartPt = New AutoMath.DPosition
    Set oEndPt = New AutoMath.DPosition
            
  'Checking for the PartDataBasis
    Dim oPipeComponent As IJDPipeComponent
    Dim lPartDataBasis As Long
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis
    Set oPipeComponent = Nothing
    
    Select Case lPartDataBasis

        Case Is <= 1, 780
        
            parFace1toCenter = arrayOfInputs(2)
            parFace2toCenter = arrayOfInputs(3)
            parBodyDiameter = arrayOfInputs(4)
            parHeight = arrayOfInputs(5)
            
        'Cylindrical Body
            oStartPt.Set -pipeDiam2 / 2, 0, 0
            oEndPt.Set (parHeight - parFace1toCenter), 0, 0
            Set ObjCylinBody = PlaceCylinder(m_OutputColl, oStartPt, oEndPt, parBodyDiameter, True)
        ' Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCylinBody
            Set ObjCylinBody = Nothing
            
        Case 781
            
            parFace1toCenter = arrayOfInputs(2)
            parFace2toCenter = arrayOfInputs(3)
            parBodyDiameter = arrayOfInputs(4)
            parHeight = arrayOfInputs(5)
            parOffset = arrayOfInputs(6)
            
        'Cylindrical Body
            oStartPt.Set -pipeDiam2, parOffset - (parBodyDiameter / 2), 0
            oEndPt.Set (parHeight - parFace1toCenter), parOffset - (parBodyDiameter / 2), 0
            Set ObjCylinBody = PlaceCylinder(m_OutputColl, oStartPt, oEndPt, parBodyDiameter, True)
        ' Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCylinBody
            Set ObjCylinBody = Nothing

        Case Else
            GoTo ErrorLabel: 'Invalid specification
            
        End Select
    
        Dim oPlacePoint As AutoMath.DPosition
    ' Place Nozzle 1
        Dim oDir        As AutoMath.DVector
        Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
        Set oPlacePoint = New AutoMath.DPosition
        Set oDir = New AutoMath.DVector
        
        oPlacePoint.Set -parFace1toCenter - sptOffset + depth, 0, 0
        oDir.Set -1, 0, 0
        If CmpDblLessThan(parFace1toCenter, flangeThick) Then parFace1toCenter = flangeThick
        Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, oPlacePoint, parFace1toCenter)
    ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing
        
    ' Place Nozzle 2
        Dim NozzleLength As Double
        oPlacePoint.Set 0, -parFace2toCenter - sptOffset2 + depth2, 0
        oDir.Set 0, -1, 0
        NozzleLength = parFace2toCenter
        If CmpDblLessThan(NozzleLength, flangeThick2) Then NozzleLength = flangeThick2
        Set objNozzle = CreateNozzleWithLength(2, oPartFclt, m_OutputColl, oDir, oPlacePoint, NozzleLength)
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

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
'   Creation Date:  Tuesday, April 1 2008
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   01.04.2008      VRK     CR-136575:Enhance rupture disc symbol to provide more realistic dimensional basis
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private PI As Double
Private Const MODULE = "Physical:"    'Used for error messages
Private Sub Class_Initialize()
Const METHOD = "Class_Initialize:"
    On Error GoTo Errx
   
    PI = Atn(1) * 4
     
    Exit Sub
Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.description, _
       Err.HelpFile, Err.HelpContext
End Sub
Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    Dim pipeDiam2       As Double
    Dim flangeThick2    As Double
    Dim sptOffset2      As Double
    Dim flangeDiam2     As Double
    Dim depth2          As Double
    Dim pipeDiam3       As Double
    Dim flangeThick3    As Double
    Dim sptOffset3      As Double
    Dim flangeDiam3     As Double
    Dim depth3          As Double
    Dim CenterPos       As New AutoMath.DPosition
    Dim iOutput         As Integer

    CenterPos.Set 0, 0, 0
    
    Dim parHeight As Double
    Dim parBodyOutsideDiameter As Double
    Dim parHubLength As Double
    Dim parNutHeight As Double
    Dim parNutWidth As Double
    Dim parInsulationThickness As Double
    
    Dim oStPoint   As New AutoMath.DPosition
    Dim oEnPoint   As New AutoMath.DPosition

' Inputs
    Set oPartFclt = arrayOfInputs(1)
'    parHeight = arrayOfInputs(2)
'    parBodyOutsideDiameter = arrayOfInputs(3)
'    parHubLength = arrayOfInputs(4)
'    parNutHeight = arrayOfInputs(5)
'    parNutWidth = arrayOfInputs(6)
'    parInsulationThickness = arrayOfInputs(7)
    
    iOutput = 0
    Dim ObjCylinder As Object
    Dim objDiscTag As Object
    Dim objCone As Object
    Dim objUnion As Object
    
    Dim oPipeComponent As IJDPipeComponent
    Set oPipeComponent = oPartFclt
    Dim lPartDataBasis As Integer
    lPartDataBasis = oPipeComponent.PartDataBasis
    Set oPipeComponent = Nothing
    
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, _
                                    flangeDiam, sptOffset, depth
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick2, _
                                    flangeDiam2, sptOffset2, depth2

    Select Case lPartDataBasis
        Case Is <= 1, 785
            parHeight = arrayOfInputs(2)
        Case 786
            parHeight = arrayOfInputs(2)
        Case 787
            parHeight = arrayOfInputs(2)
        Case 788
            parHeight = arrayOfInputs(2)
            parBodyOutsideDiameter = arrayOfInputs(3)
        Case 789
            parHeight = arrayOfInputs(2)
            parBodyOutsideDiameter = arrayOfInputs(3)
            parHubLength = arrayOfInputs(4)
        Case 790
            parHeight = arrayOfInputs(2)
            parHubLength = arrayOfInputs(4)
        Case 791
            parHeight = arrayOfInputs(2)
            parBodyOutsideDiameter = arrayOfInputs(3)
            parHubLength = arrayOfInputs(4)
        Case 792
            parHeight = arrayOfInputs(2)
            parBodyOutsideDiameter = arrayOfInputs(3)
            parHubLength = arrayOfInputs(4)
        Case 793
            parHeight = arrayOfInputs(2)
        Case 794
            parHeight = arrayOfInputs(2)
            parHubLength = arrayOfInputs(4)
        Case 795
            parHeight = arrayOfInputs(2)
            parNutHeight = arrayOfInputs(5)
            parNutWidth = arrayOfInputs(6)
        Case Else
            GoTo ErrorLabel: 'Invalid specification
        
    End Select
    
    If lPartDataBasis = 787 Or lPartDataBasis = 790 Then
        flangeDiam = flangeDiam2
    ElseIf lPartDataBasis = 788 Or lPartDataBasis = 789 Or lPartDataBasis = 791 _
                    Or lPartDataBasis = 792 Then
        flangeDiam = parBodyOutsideDiameter
    ElseIf lPartDataBasis = 795 Then
        flangeDiam = pipeDiam
    End If
    
    Dim dhalfHeight As Double
    Dim dRuptureDiscOffset As Double
    Dim dRuptureDiscTagOffset As Double
    Dim dRuptureDiscTagWidth As Double
    dhalfHeight = 0.5 * parHeight
    dRuptureDiscOffset = 0.3 * parHeight
    dRuptureDiscTagOffset = 0.8 * flangeDiam
    dRuptureDiscTagWidth = 0.6 * flangeDiam
    
    'Create Cylinder
        oStPoint.Set -dhalfHeight, 0, 0
        oEnPoint.Set dhalfHeight, 0, 0
        Set ObjCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, flangeDiam, True)
    ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCylinder
        Set ObjCylinder = Nothing
        
    If Not lPartDataBasis = 795 Then
       'Create Disc Tag
        oStPoint.Set -(dhalfHeight - dRuptureDiscOffset), _
                           Sqr((0.5 * flangeDiam) ^ 2 - (0.25 * dRuptureDiscTagWidth) ^ 2), -0.25 * dRuptureDiscTagWidth
        oEnPoint.Set oStPoint.x + 0.25 * dRuptureDiscOffset, _
                    oStPoint.y + 0.6 * (dRuptureDiscTagOffset - Sqr((0.5 * flangeDiam) ^ 2 - (0.25 * dRuptureDiscTagWidth) ^ 2)), _
                    oStPoint.z + 0.5 * dRuptureDiscTagWidth
        Set objDiscTag = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
    ' Set the output
        m_OutputColl.AddOutput "DiscTag_", objDiscTag
        Set objDiscTag = Nothing

        oStPoint.Set -(dhalfHeight - dRuptureDiscOffset), _
                           oEnPoint.y, -0.5 * dRuptureDiscTagWidth
        oEnPoint.Set oStPoint.x + 0.25 * dRuptureDiscOffset, _
                    oStPoint.y + 0.4 * (dRuptureDiscTagOffset - Sqr((0.5 * flangeDiam) ^ 2 - (0.25 * dRuptureDiscTagWidth) ^ 2)), _
                    oStPoint.z + dRuptureDiscTagWidth
        Set objDiscTag = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
    ' Set the output
        m_OutputColl.AddOutput "DiscTag_", objDiscTag
        Set objDiscTag = Nothing
    End If
    
    ' Place Nozzle 1
        Dim oPlacePoint As AutoMath.DPosition
        Dim oDir        As AutoMath.DVector
        Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
        Set oPlacePoint = New AutoMath.DPosition
        Set oDir = New AutoMath.DVector
        oPlacePoint.Set CenterPos.x - dhalfHeight - sptOffset + depth, _
                        CenterPos.y, CenterPos.z
        If lPartDataBasis = 790 Or lPartDataBasis = 791 Or lPartDataBasis = 792 Then
            oPlacePoint.Set CenterPos.x - dhalfHeight - sptOffset + depth - parHubLength, _
                        CenterPos.y, CenterPos.z
        ElseIf lPartDataBasis = 787 Or lPartDataBasis = 788 Or lPartDataBasis = 789 Then
            oPlacePoint.Set CenterPos.x - dhalfHeight - sptOffset + depth - flangeThick, _
                        CenterPos.y, CenterPos.z
        End If
        oDir.Set -1, 0, 0
        Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, _
                            oPlacePoint)
    '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing
        
    ' Place Nozzle 2
        oPlacePoint.Set CenterPos.x + dhalfHeight + sptOffset2 - depth2, _
                        CenterPos.y, CenterPos.z
        If lPartDataBasis = 789 Or lPartDataBasis = 792 Or lPartDataBasis = 794 Then
            oPlacePoint.Set CenterPos.x + dhalfHeight + sptOffset2 - depth2 + parHubLength, _
                        CenterPos.y, CenterPos.z
        ElseIf lPartDataBasis = 788 Or lPartDataBasis = 791 Or lPartDataBasis = 793 Then
            oPlacePoint.Set CenterPos.x + dhalfHeight + sptOffset2 - depth2 + flangeThick2, _
                        CenterPos.y, CenterPos.z
        End If
        oDir.Set 1, 0, 0
        Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, _
                             oPlacePoint)
    '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing
        
     If lPartDataBasis = 786 Then
        ' Place Nozzle 3-Gauge Tap
            RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam3, flangeThick3, _
                                        flangeDiam3, sptOffset3, depth3
    
            Dim dGaugeTapOffset As Double
            dGaugeTapOffset = 0.3 * parHeight
            oPlacePoint.Set CenterPos.x + dhalfHeight + sptOffset3 - depth3 - dGaugeTapOffset, _
                            CenterPos.y + 0.5 * flangeDiam, CenterPos.z
            oDir.Set 0, 1, 0
            Set objNozzle = CreateNozzleJustaCircle(3, oPartFclt, m_OutputColl, oDir, _
                                  oPlacePoint)
        '   Set the output
            m_OutputColl.AddOutput "GaugeTap", objNozzle
            Set objNozzle = Nothing
    End If
    
    If lPartDataBasis = 789 Or lPartDataBasis = 792 Or lPartDataBasis = 794 Then
        oStPoint.Set dhalfHeight + parHubLength, 0, 0
        oEnPoint.Set dhalfHeight, 0, 0
        Set objCone = PlaceCone(m_OutputColl, oStPoint, oEnPoint, 0.5 * pipeDiam2, 0.25 * (pipeDiam2 + flangeDiam), True)
    '   Set the output
        m_OutputColl.AddOutput "ObjCone1", objCone
        Set objCone = Nothing
    End If
    If lPartDataBasis = 790 Or lPartDataBasis = 791 Or lPartDataBasis = 792 Then
        oStPoint.Set -(dhalfHeight + parHubLength), 0, 0
        oEnPoint.Set -dhalfHeight, 0, 0
        Set objCone = PlaceCone(m_OutputColl, oStPoint, oEnPoint, 0.5 * pipeDiam, 0.25 * (pipeDiam + flangeDiam), True)
    '   Set the output
        m_OutputColl.AddOutput "ObjCone2", objCone
        Set objCone = Nothing
    End If
    If lPartDataBasis = 795 Then
        'Create output for Union Nut
        Dim dAngle As Double
        Dim iCount As Integer
        Dim oAxisVect As AutoMath.DVector
        Dim LineStrPoints(0 To 20) As Double
        Dim oLineString As IngrGeom3D.LineString3d
        Dim oGeomFactory As IngrGeom3D.GeometryFactory
        dAngle = PI / 3
        oStPoint.Set -(parHeight - parNutHeight) / 2, 0, 0
        'Line points for Union Nut
        For iCount = 1 To 7
            LineStrPoints(3 * iCount - 3) = oStPoint.x
            LineStrPoints(3 * iCount - 2) = oStPoint.y + (parNutWidth / 2) * Sin(PI / 6 + iCount * dAngle)
            LineStrPoints(3 * iCount - 1) = oStPoint.z + (parNutWidth / 2) * Cos(PI / 6 + iCount * dAngle)
        Next iCount
        Set oGeomFactory = New IngrGeom3D.GeometryFactory
        Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 7, LineStrPoints)
        Set oAxisVect = New AutoMath.DVector
        oAxisVect.Set 1, 0, 0
        Set objUnion = PlaceProjection(m_OutputColl, oLineString, oAxisVect, parNutHeight, True)
        ' Set the output
        m_OutputColl.AddOutput "Union", objUnion
        Set objUnion = Nothing
        Set oAxisVect = Nothing
        Set oLineString = Nothing
        Set oGeomFactory = Nothing
    End If
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    Set CenterPos = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    
    Exit Sub
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
End Sub

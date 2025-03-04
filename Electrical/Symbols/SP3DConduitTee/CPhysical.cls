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
'   Copyright (c) 2003-08, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          NN
'   Creation Date:  Sunday, March 3 2002
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol implements the following Part Data Bases
'   1)Conduit Tee, with cover at side, specified by face-to-face, height, width, cover width and cover length(9032)
'   2)Conduit Tee, with cover at the bottom, specified by face-to-face, height, width, cover width and cover length(9033)
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   08.SEP.2006     KKC    DI-95670         Replace names with initials in all revision history sheets and symbols
'   22.May.2008     PK     CR-93287         Enhanced the tee symbol implementing partdatabases 9032 and 9033
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Physical:" 'Used for error messages
Private Const COVERHEIGHT = 0.002    'Assumed height of cover
Private Sub Class_Initialize()

'''

End Sub
Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim ConduitDiam     As Double
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    
    Dim iOutput         As Double
    Dim ObjNozzle1      As GSCADNozzleEntities.IJConduitPortOcc
    Dim ObjNozzle2      As GSCADNozzleEntities.IJConduitPortOcc
    Dim ObjNozzle3      As GSCADNozzleEntities.IJConduitPortOcc
    Dim CenterPos       As New AutoMath.DPosition
    Dim parFacetoCenter As Double
    
    Dim parFacetoFace           As Double
    Dim parHeight               As Double
    Dim parWidth                As Double
    Dim parCoverLength          As Double
    Dim parCoverWidth           As Double

    Dim lPartdatabasis          As Long
    Dim oPipeComponent          As IJDPipeComponent
    
    Dim dWidth                  As Double
    Dim dHeight                 As Double
    
    CenterPos.Set 0, 0, 0
    
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    'parFacetoCenter = arrayOfInputs(2)
    'parFacetoFace = arrayOfInputs(3)
    'parHeight = arrayOfInputs(4)
    'parWidth = arrayOfInputs(5)
    'parCoverLength = arrayOfInputs(6)
    'parCoverWidth = arrayOfInputs(7)
    
    Set oPipeComponent = oPartFclt
    lPartdatabasis = oPipeComponent.PartDataBasis
    
    iOutput = 0
    
    If lPartdatabasis <= 1 Then
        parFacetoCenter = arrayOfInputs(2)
    ElseIf lPartdatabasis = 9032 Or lPartdatabasis = 9033 Then
        parFacetoFace = arrayOfInputs(3)
        parHeight = arrayOfInputs(4)
        parWidth = arrayOfInputs(5)
        parCoverLength = arrayOfInputs(6)
        parCoverWidth = arrayOfInputs(7)
    End If
    
    If lPartdatabasis = 9032 Then       'Conduit Tee, with cover at side, specified
        dWidth = 0.7 * parWidth         'by face-to-face, height, width, cover width
        dHeight = parHeight             'and cover length
        
    ElseIf lPartdatabasis = 9033 Then   'Conduit Tee, with cover at the bottom, specified
        dWidth = parWidth               'by face-to-face, height, width, cover
        dHeight = 0.7 * parHeight       'width and cover length
    End If

    If lPartdatabasis <= 1 Then
    
        'Insert your code for output 1(Cylinder Body)
        RetrieveParameters 1, oPartFclt, m_OutputColl, ConduitDiam, flangeThick, flangeDiam, sptOffset, depth
        
        Dim stPoint     As AutoMath.DPosition
        Dim enPoint     As AutoMath.DPosition
        Dim ObjTeeBody1 As Object
        Dim ConduitOD1  As Double
        Set stPoint = New AutoMath.DPosition
        Set enPoint = New AutoMath.DPosition
        
        ConduitOD1 = BodyOD(flangeDiam, ConduitDiam)
        stPoint.Set -parFacetoCenter, 0, 0
        enPoint.Set parFacetoCenter, 0, 0
        Set ObjTeeBody1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, ConduitOD1, True)
        
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjTeeBody1
        Set ObjTeeBody1 = Nothing
        
        'Insert your code for output 2(Branch Tee Body)
        Dim stPoint2    As AutoMath.DPosition
        Dim enPoint2    As AutoMath.DPosition
        Dim ObjTeeBody3 As Object
        Dim ConduitOD3  As Double
        Set stPoint2 = New AutoMath.DPosition
        Set enPoint2 = New AutoMath.DPosition
        
        ConduitOD3 = BodyOD(flangeDiam, ConduitDiam)
        stPoint2.Set 0, 0, 0
        enPoint2.Set 0, parFacetoCenter, 0
        Set ObjTeeBody3 = PlaceCylinder(m_OutputColl, stPoint2, enPoint2, ConduitOD3, True)
        
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjTeeBody3
        Set ObjTeeBody3 = Nothing
            
    Else
        'Insert your code for Body
        RetrieveParameters 1, oPartFclt, m_OutputColl, ConduitDiam, flangeThick, flangeDiam, sptOffset, depth
      
        Dim oGeomFactory    As IngrGeom3D.GeometryFactory
        Dim oLine           As Object
        Dim oArc            As IngrGeom3D.Arc3d
        Dim lines           As Collection
        Dim oCString        As IngrGeom3D.ComplexString3d
        Dim oStPoint        As AutoMath.DPosition
        Dim iCount          As Integer
        Dim oAxis           As AutoMath.DVector
        Dim oBody           As Object

        Set oGeomFactory = New IngrGeom3D.GeometryFactory
        Set oCString = New IngrGeom3D.ComplexString3d
        Set oArc = New IngrGeom3D.Arc3d
        Set lines = New Collection
        Set oStPoint = New AutoMath.DPosition
        Set oAxis = New AutoMath.DVector
        
        'Insert your code for Complex String
        Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
            -0.35 * parFacetoFace + 0.5 * dWidth, 0.5 * dWidth, 0.5 * dHeight, _
            0.35 * parFacetoFace - 0.5 * dWidth, 0.5 * dWidth, 0.5 * dHeight)
        lines.Add oLine

        Set oArc = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, _
            0.35 * parFacetoFace - 0.5 * dWidth, 0.5 * dWidth, 0.5 * dHeight, _
            0.35 * parFacetoFace, 0, 0.5 * dHeight, _
            0.35 * parFacetoFace - 0.5 * dWidth, -0.5 * dWidth, 0.5 * dHeight)
        lines.Add oArc

        Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
            0.35 * parFacetoFace - 0.5 * dWidth, -0.5 * dWidth, 0.5 * dHeight, _
            -0.35 * parFacetoFace + 0.5 * dWidth, -0.5 * dWidth, 0.5 * dHeight)
        lines.Add oLine

        Set oArc = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, _
            -0.35 * parFacetoFace + 0.5 * dWidth, -0.5 * dWidth, 0.5 * dHeight, _
            -0.35 * parFacetoFace, 0, 0.5 * dHeight, _
            -0.35 * parFacetoFace + 0.5 * dWidth, 0.5 * dWidth, 0.5 * dHeight)
        lines.Add oArc

        oStPoint.Set -0.35 * parFacetoFace + 0.5 * dWidth, 0.5 * dWidth, 0.5 * dHeight
        Set oCString = PlaceTrCString(oStPoint, lines)
        
        oAxis.Set 0, 0, -1
        Set oBody = PlaceProjection(m_OutputColl, oCString, oAxis, dHeight, True)
        
        iOutput = iOutput + 1
         m_OutputColl.AddOutput arrayOfOutputs(iOutput), oBody
        
        For iCount = 1 To lines.Count
            lines.Remove 1
        Next iCount
        Set lines = Nothing
        Set oStPoint = Nothing
        Set oLine = Nothing
        Set oArc = Nothing
        Set oBody = Nothing
        Set oAxis = Nothing
        
        'Insert your code for Cover
        Dim oCoverString    As IngrGeom3D.ComplexString3d
        Dim oCover          As Object
        Set oCoverString = New IngrGeom3D.ComplexString3d
        
        Set lines = New Collection
        Set oArc = New IngrGeom3D.Arc3d
        Set oStPoint = New AutoMath.DPosition
        Set oAxis = New AutoMath.DVector
              
        Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
           -parCoverLength / 2 + parCoverWidth / 2, parCoverWidth / 2, 0.5 * dHeight, _
           parCoverLength / 2 - parCoverWidth / 2, parCoverWidth / 2, 0.5 * dHeight)
        lines.Add oLine

        Set oArc = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, _
           parCoverLength / 2 - parCoverWidth / 2, parCoverWidth / 2, 0.5 * dHeight, _
           parCoverLength / 2, 0, 0.5 * dHeight, _
           parCoverLength / 2 - parCoverWidth / 2, -parCoverWidth / 2, 0.5 * dHeight)
        lines.Add oArc
        
        Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
           parCoverLength / 2 - parCoverWidth / 2, -parCoverWidth / 2, 0.5 * dHeight, _
           -parCoverLength / 2 + parCoverWidth / 2, -parCoverWidth / 2, 0.5 * dHeight)
        lines.Add oLine
        
        Set oArc = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, _
           -parCoverLength / 2 + parCoverWidth / 2, -parCoverWidth / 2, 0.5 * dHeight, _
           -parCoverLength / 2, 0, 0.5 * dHeight, _
           -parCoverLength / 2 + parCoverWidth / 2, parCoverWidth / 2, 0.5 * dHeight)
        lines.Add oArc
        
        oStPoint.Set -parCoverLength / 2 + parCoverWidth / 2, _
            parCoverWidth / 2, 0.5 * dHeight
        Set oCoverString = PlaceTrCString(oStPoint, lines)
        
        oAxis.Set 0, 0, 1
        Set oCover = PlaceProjection(m_OutputColl, oCoverString, oAxis, COVERHEIGHT, True)

        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), oCover
    
        Set oGeomFactory = Nothing
        Set oArc = Nothing
        Set oLine = Nothing
        For iCount = 1 To lines.Count
            lines.Remove 1
        Next iCount
        Set lines = Nothing
        Set oStPoint = Nothing
        Set oCoverString = Nothing
            
        'Insert your code for Cylinder at Port1
        Dim oCylinder1  As Object
        Dim oEnPoint    As AutoMath.DPosition
        Set oEnPoint = New AutoMath.DPosition
        Set oStPoint = New AutoMath.DPosition
           
        oStPoint.Set -0.5 * parFacetoFace, 0, 0
        oEnPoint.Set -0.2 * parFacetoFace, 0, 0
        
        Set oCylinder1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 1.1 * ConduitDiam, True)

        'Set the Output
        m_OutputColl.AddOutput "Cylinder1", oCylinder1
        Set oCylinder1 = Nothing
        Set oStPoint = Nothing
        Set oEnPoint = Nothing
        
        RetrieveParameters 2, oPartFclt, m_OutputColl, ConduitDiam, flangeThick, flangeDiam, sptOffset, depth
        
        'Insert your code for Cylinder at Port2
        Dim oCylinder2 As Object
        Set oEnPoint = New AutoMath.DPosition
        Set oStPoint = New AutoMath.DPosition
           
        oStPoint.Set 0.2 * parFacetoFace, 0, 0
        oEnPoint.Set 0.5 * parFacetoFace, 0, 0
        
        Set oCylinder2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 1.1 * ConduitDiam, True)

        'Set the Output
        m_OutputColl.AddOutput "Cylinder2", oCylinder2
        Set oCylinder2 = Nothing
        Set oStPoint = Nothing
        Set oEnPoint = Nothing
        
        RetrieveParameters 3, oPartFclt, m_OutputColl, ConduitDiam, flangeThick, flangeDiam, sptOffset, depth
        
        'Insert your code for Cylinder at Port3
        Dim oCylinder3 As Object
        Set oEnPoint = New AutoMath.DPosition
        Set oStPoint = New AutoMath.DPosition
        
        If lPartdatabasis = 9032 Then
            oStPoint.Set 0, 0.35 * parWidth, 0
            oEnPoint.Set 0, 0.65 * parWidth, 0
        ElseIf lPartdatabasis = 9033 Then
            oStPoint.Set 0, 0, -0.35 * parHeight
            oEnPoint.Set 0, 0, -0.65 * parHeight
        End If
        
        Set oCylinder3 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 1.1 * ConduitDiam, True)
    
        'Set the Output
        m_OutputColl.AddOutput "Cylinder3", oCylinder3
        Set oCylinder3 = Nothing
        Set oStPoint = Nothing
        Set oEnPoint = Nothing
        
    End If
        
    'Insert your code for output 3(Nozzle1)
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    RetrieveParameters 1, oPartFclt, m_OutputColl, ConduitDiam, flangeThick, flangeDiam, sptOffset, depth
    
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    If lPartdatabasis <= 1 Then
        oPlacePoint.Set -parFacetoCenter - sptOffset + depth, 0, 0
        oDir.Set -1, 0, 0
    ElseIf lPartdatabasis = 9032 Or lPartdatabasis = 9033 Then
        oPlacePoint.Set -0.5 * parFacetoFace - sptOffset + depth, 0, 0
        oDir.Set -1, 0, 0
    End If
    
    Set ObjNozzle1 = CreateConduitNozzle(oPlacePoint, oDir, m_OutputColl, oPartFclt, 1)
    
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjNozzle1
    Set ObjNozzle1 = Nothing
    
    'Insert your code for output 4(Nozzle2)
    RetrieveParameters 2, oPartFclt, m_OutputColl, ConduitDiam, flangeThick, flangeDiam, sptOffset, depth
    
    If lPartdatabasis <= 1 Then
        oDir.Set 1, 0, 0
        oPlacePoint.Set parFacetoCenter + sptOffset - depth, 0, 0
    ElseIf lPartdatabasis = 9032 Or lPartdatabasis = 9033 Then
        oPlacePoint.Set 0.5 * parFacetoFace + sptOffset - depth, 0, 0
        oDir.Set 1, 0, 0
    End If
        
    Set ObjNozzle2 = CreateConduitNozzle(oPlacePoint, oDir, m_OutputColl, oPartFclt, 2)
    
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjNozzle2
    Set ObjNozzle2 = Nothing

    'Insert your code for output 5(Nozzle3)
    RetrieveParameters 3, oPartFclt, m_OutputColl, ConduitDiam, flangeThick, flangeDiam, sptOffset, depth
        
    If lPartdatabasis <= 1 Then
        oDir.Set 0, 1, 0
        oPlacePoint.Set 0, parFacetoCenter + sptOffset - depth, 0
    ElseIf lPartdatabasis = 9032 Then
        oDir.Set 0, 1, 0
        oPlacePoint.Set 0, 0.65 * parWidth + sptOffset - depth, 0
    ElseIf lPartdatabasis = 9033 Then
        oDir.Set 0, 0, -1
        oPlacePoint.Set 0, 0, -0.65 * parHeight - sptOffset + depth
    End If
  
    Set ObjNozzle3 = CreateConduitNozzle(oPlacePoint, oDir, m_OutputColl, oPartFclt, 3)
    
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjNozzle3
    Set ObjNozzle3 = Nothing
   
    Exit Sub
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub

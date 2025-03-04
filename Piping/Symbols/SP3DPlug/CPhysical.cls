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
'   Creation Date:  Thursday, Jan 25 2001
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'   16.JUL.2008     MP   CR-145604   Implemented part data basis options 1028,1029,1030,1031 and 1032
'   24.Sep.2008     MP     CR-149961    Implemented the part data basis options 1084, 1085 and 1086
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

    Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages
Private PI As Double


Private Sub Class_Initialize()

PI = 4 * Atn(1)

End Sub


Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)

    Const METHOD = "run"
    On Error GoTo ErrorLabel
    Dim oPartFclt As PartFacelets.IJDPart
    Dim pipeDiam As Double
    Dim flangeThick As Double
    Dim cptOffset As Double
    Dim flangeDiam As Double
    Dim depth As Double
    Dim StartPos As New AutoMath.DPosition
    Dim lPartDataBasis As Long
    Dim oPipeComponent As IJDPipeComponent

    StartPos.Set 0, 0, 0


    Dim iOutput As Double
    Dim ObjBody As Object
    Dim ObjBody2 As Object
    Dim ObjBody3 As Object
    Dim ObjInsulatedBody As Object

    Dim parFacetoEnd As Double
    Dim parInsulationThickness As Double
    Dim parLength1 As Double
    Dim parLength2 As Double
    Dim parWidth As Double
    
 

    ' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoEnd = arrayOfInputs(2)
    parInsulationThickness = arrayOfInputs(3)
    
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis

    If lPartDataBasis <= 1 Or lPartDataBasis = PLUG_SQUARE Or lPartDataBasis = PLUG_PENTAGON Or _
                              lPartDataBasis = PLUG_HEXAGON Or lPartDataBasis = PLUG_OCTAGON Then
        parLength1 = arrayOfInputs(4)
        parLength2 = arrayOfInputs(5)
    End If
    parWidth = arrayOfInputs(6)

    iOutput = 0

    ' Insert your code for output 1(Body)
    Dim objPlugEndPosition As AutoMath.DPosition
    Set objPlugEndPosition = New AutoMath.DPosition
 
    If lPartDataBasis <= 1 Or lPartDataBasis = PLUG_DEFAULT Then

        objPlugEndPosition.Set parFacetoEnd, 0, 0
        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth
        Set ObjBody = PlaceCylinder(m_OutputColl, StartPos, objPlugEndPosition, pipeDiam, True)

        ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBody
        Set ObjBody = Nothing
        
    ElseIf lPartDataBasis = PLUG_SQUARE Or lPartDataBasis = PLUG_PLAIN Or _
                        lPartDataBasis = PLUG_ROUND Or lPartDataBasis = PLUG_COUNTERSUNK Then
    
        If lPartDataBasis = PLUG_PLAIN Or lPartDataBasis = PLUG_ROUND Or lPartDataBasis = PLUG_COUNTERSUNK Then
            Dim dLength1 As Double
            dLength1 = 0.2 * parFacetoEnd
            objPlugEndPosition.Set parFacetoEnd - dLength1, 0, 0
        Else
            objPlugEndPosition.Set parLength1, 0, 0
        End If
        
        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth
        Set ObjBody = PlaceCylinder(m_OutputColl, StartPos, objPlugEndPosition, pipeDiam, True)
        
        ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBody
        Set ObjBody = Nothing
        Set StartPos = Nothing
        
        ' create output (body2)
        If parWidth = 0 Then
            parWidth = 0.6 * pipeDiam
        End If
        
        If lPartDataBasis = PLUG_PLAIN Then
            StartPos.Set parFacetoEnd - dLength1, -parWidth / 2, -parWidth / 2
            objPlugEndPosition.Set parFacetoEnd, parWidth / 2, parWidth / 2
            Set ObjBody2 = PlaceBox(m_OutputColl, StartPos, objPlugEndPosition)
            ' Set the output
            m_OutputColl.AddOutput "Body2", ObjBody2
            Set ObjBody2 = Nothing
        
        ElseIf lPartDataBasis = PLUG_ROUND Then
            StartPos.Set parFacetoEnd - dLength1, 0, 0
            objPlugEndPosition.Set parFacetoEnd, 0, 0
            Set ObjBody2 = PlaceCylinder(m_OutputColl, StartPos, objPlugEndPosition, parWidth, True)
            ' Set the output
            m_OutputColl.AddOutput "Body2", ObjBody2
            Set ObjBody2 = Nothing

        ElseIf lPartDataBasis = PLUG_COUNTERSUNK Then
            Dim objCircle As IngrGeom3D.Circle3d
            Dim oTopSurface As IngrGeom3D.Plane3d
            Dim ObjCmplxStr As IngrGeom3D.ComplexString3d
            Dim oGeomFactory    As IngrGeom3D.GeometryFactory
            Dim lines           As Collection
            Dim oLine           As IngrGeom3D.Line3d
            Dim oCSqrStr         As IngrGeom3D.ComplexString3d
            Dim oStpoint As AutoMath.DPosition
            Set lines = New Collection
            
            Set oGeomFactory = New IngrGeom3D.GeometryFactory
            Set objCircle = New IngrGeom3D.Circle3d
            Set oStpoint = New AutoMath.DPosition
            
            Set oTopSurface = oGeomFactory.Planes3d.CreateByPointNormal(Nothing, parFacetoEnd, 0, 0, 1, 0, 0)
            
            Set objCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, parFacetoEnd, 0, 0, 1, 0, 0, parWidth)
            
            Dim ObjCmplx As IngrGeom3D.ComplexString3d
            Set ObjCmplx = New IngrGeom3D.ComplexString3d
            ObjCmplx.AddCurve objCircle, True
            oTopSurface.AddBoundary ObjCmplx
                        
            Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, parFacetoEnd, parWidth / 2, parWidth / 2, parFacetoEnd, -parWidth / 2, parWidth / 2)
            lines.Add oLine
            
            Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, parFacetoEnd, -parWidth / 2, parWidth / 2, parFacetoEnd, -parWidth / 2, -parWidth / 2)
            lines.Add oLine
            
            Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, parFacetoEnd, -parWidth / 2, -parWidth / 2, parFacetoEnd, parWidth / 2, -parWidth / 2)
            lines.Add oLine
            
            Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, parFacetoEnd, parWidth / 2, -parWidth / 2, parFacetoEnd, parWidth / 2, parWidth / 2)
            lines.Add oLine
            
            oStpoint.Set parFacetoEnd, parWidth / 2, parWidth / 2
            
            Set oCSqrStr = PlaceTrCString(oStpoint, lines)
            Set oLine = Nothing
            Dim iCount As Integer
            
            For iCount = 1 To lines.Count
                lines.Remove 1
            Next iCount
            Set lines = Nothing
            
            oTopSurface.AddHole oCSqrStr
'            Dim objSqrProj As Object
            Dim oAxis As AutoMath.DVector
            Dim oSurfset As IngrGeom3D.IJElements
            Set oAxis = New AutoMath.DVector
            oAxis.Set -1, 0, 0
            Set oSurfset = oGeomFactory.GeometryServices.CreateByProjection(m_OutputColl.ResourceManager, oTopSurface, -1, 0, 0, 0.2 * parFacetoEnd, 1)
            Dim oProjection As Object
            For Each oProjection In oSurfset
                m_OutputColl.AddOutput "Body2", oProjection
            Next oProjection
            
        Else
            StartPos.Set parLength1, -parWidth / 2, -parWidth / 2
            objPlugEndPosition.Set parLength1 + parLength2, parWidth / 2, parWidth / 2
            Set ObjBody2 = PlaceBox(m_OutputColl, StartPos, objPlugEndPosition)
        
            ' Set the output
            m_OutputColl.AddOutput "Body2", ObjBody2
            Set ObjBody2 = Nothing
            
            Set oAxis = Nothing
            Set ObjBody2 = Nothing
                        
        End If
                

    ElseIf lPartDataBasis = PLUG_HEXAGON Or lPartDataBasis = PLUG_OCTAGON Or lPartDataBasis = PLUG_PENTAGON Then
        Dim lNumSides As Long
        Dim oTransMat As AutoMath.DT4x4
        Dim oVec As AutoMath.DVector
        Set oVec = New AutoMath.DVector
        Set oTransMat = New AutoMath.DT4x4
        
        If lPartDataBasis = PLUG_HEXAGON Then
            lNumSides = 6
        ElseIf lPartDataBasis = PLUG_OCTAGON Then
            lNumSides = 8
        ElseIf lPartDataBasis = PLUG_PENTAGON Then
            lNumSides = 10
        End If
        
        Dim dInculdedAngle As Double
        Dim dSideLength As Double
        
        dInculdedAngle = 360 / lNumSides
        dInculdedAngle = (dInculdedAngle * PI) / 180
        objPlugEndPosition.Set parLength1, 0, 0
        
        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth
        
        If CmpDblEqual(parWidth, 0) Then
            'Assuming that side length of the polygon shaped hub is 0.7 times the Piping Outside Diameter of Port1
            dSideLength = 0.8 * pipeDiam
        Else
            dSideLength = parWidth * Sin(dInculdedAngle / 2)
        End If

        Set ObjBody = PlaceCylinder(m_OutputColl, StartPos, objPlugEndPosition, pipeDiam, True)

        ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBody
        Set ObjBody = Nothing
        Set StartPos = Nothing
        Set objPlugEndPosition = Nothing

        Set ObjBody2 = PlaceNnagon(m_OutputColl, lNumSides, dSideLength, parLength2, True)

        oVec.Set parLength1, 0, 0
        oTransMat.LoadIdentity
        oTransMat.Translate oVec
        ObjBody2.Transform oTransMat
        
        ' Set the output
        m_OutputColl.AddOutput "Body2", ObjBody2
        Set ObjBody2 = Nothing
    End If

    ' Place Nozzle 1
    Dim oDir As AutoMath.DVector
    Dim objNozzle As GSCADNozzleEntities.IJDNozzle
    Set oDir = New AutoMath.DVector
    StartPos.Set 0, 0, 0
    oDir.Set -1, 0, 0
    Set oPartFclt = arrayOfInputs(1)
    Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, StartPos)
    
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing

    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub


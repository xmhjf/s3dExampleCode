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
'   ProgID:         SP3DOPLever.OPLever
'   Author:         RUK
'   Creation Date:  Wednesday, Sugust 27, 2008
'   Description: This symbol is graphical implementation of Lever operator
'   Source: E-141 section of the Design document)
'   This symbol implements following partdatabasis
'   Default
'   Lever, Type 1   (60)
'   Lever, inclined, Type 1 (61)
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   27.08.2008      RUK     CR-148069  Provide a more realistic lever operator symbol (Source - E-141 section of the Design document)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Physical:"    'Used for error messages
Private PI As Double
Private Const Lever_Type1 = 60
Private Const Lever_Inclined_Type1 = 61

Private Sub Class_Initialize()
    PI = 4 * Atn(1)
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)

    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt As PartFacelets.IJDPart
    Dim iOutput As Long
    
    Dim parOffsetFrmValCen As Double
    Dim parLeverLength As Double
    Dim parLeverLength1 As Double
    Dim parLeverHeight As Double
    Dim parLeverHeight1 As Double
    Dim parLeverHeight2 As Double
    
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    'parOffsetFrmValCen = arrayOfInputs(2)
    'parLeverLength = arrayOfInputs(3)
    'parLeverLength1 = arrayOfInputs(4)
    'parLeverHeight = arrayOfInputs(5)
    'parLeverHeight1 = arrayOfInputs(6)
    'parLeverHeight2 = arrayOfInputs(7)

    iOutput = 0
    
    Dim oGeomFact As IngrGeom3D.GeometryFactory
    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    Dim oVector As AutoMath.DVector
    Dim oTransMat As AutoMath.DT4x4
    Dim oRotMat As AutoMath.DT4x4
    Dim oLineStr As IngrGeom3D.LineString3d
    
    Set oGeomFact = New GeometryFactory
    Set oStPoint = New DPosition
    Set oEnPoint = New DPosition
    Set oVector = New DVector
    Set oTransMat = New DT4x4
    Set oRotMat = New DT4x4
    Set oLineStr = New LineString3d
    
    Dim objLeverPart1 As Object
    Dim objLeverPart2() As Object
    Dim oCrossStr() As IngrGeom3D.LineString3d
    Dim objHandle As Object
    
    Dim dInclAngle As Double
    Dim dHandleAngle As Double
    Dim dLeverWidth As Double
    Dim dLeverLength1 As Double
    Dim dLeverHeight1 As Double
    Dim dLeverHeight2 As Double
    Dim dLeverLength As Double
    Dim dPoints() As Double
    Dim dHandleLength As Double
    Dim iCount As Long
     
    'Checking for the PartDataBasis
    Dim oOperator As IJValveOperator
    Dim lPartDataBasis As Long
    Set oOperator = oPartFclt
    lPartDataBasis = oOperator.DimensionalBasis
    Set oOperator = Nothing
    
    Select Case lPartDataBasis
    Case Is <= 1, Lever_Type1
        parOffsetFrmValCen = arrayOfInputs(2)
        parLeverLength = arrayOfInputs(3)
        parLeverLength1 = arrayOfInputs(4)
        parLeverHeight1 = arrayOfInputs(6)
        parLeverHeight2 = arrayOfInputs(7)
        
        dLeverWidth = 0.15 * parLeverLength
        If CmpDblLessThan(dLeverWidth, 0.05) Then dLeverWidth = 0.05
        If CmpDblGreaterthan(dLeverWidth, 0.1) Then dLeverWidth = 0.1
        
        dLeverLength = parLeverLength
        If CmpDblGreaterthan(parLeverLength1, LINEAR_TOLERANCE) Then
            dLeverLength1 = parLeverLength1
        Else
            dLeverLength1 = 0.4 * parLeverLength
        End If
        
        dLeverHeight1 = parLeverHeight1
        If CmpDblGreaterthan(parLeverHeight2, LINEAR_TOLERANCE) Then
            dLeverHeight2 = parLeverHeight2
        Else
            dLeverHeight2 = 1.2 * dLeverHeight1
        End If
        dHandleAngle = Atn((dLeverHeight2 + 0.1 * dLeverHeight1) / _
                        (0.8 * (parLeverLength - dLeverLength1)))
                        
        'Create the Lever part 1
        oStPoint.Set 0, parOffsetFrmValCen, 0
        oEnPoint.Set 0, parOffsetFrmValCen + dLeverHeight1, 0
        Set objLeverPart1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 1.5 * dLeverWidth, True)
        
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objLeverPart1
        Set objLeverPart1 = Nothing
            
        'Create the Lever part 2
        ReDim objLeverPart2(1 To 3) As Object
        ReDim oCrossStr(1 To 4) As IngrGeom3D.LineString3d
        ReDim dPoints(0 To 14) As Double
        
        For iCount = 1 To 4
            Set oCrossStr(iCount) = New LineString3d
        Next iCount
        
        'Cross secton 1
        dPoints(0) = 0
        dPoints(1) = 0
        dPoints(2) = -dLeverWidth / 2
        
        dPoints(3) = dPoints(0)
        dPoints(4) = 0.9 * dLeverHeight1
        dPoints(5) = dPoints(2)
        
        dPoints(6) = dPoints(0)
        dPoints(7) = dPoints(4)
        dPoints(8) = dLeverWidth / 2
        
        dPoints(9) = dPoints(0)
        dPoints(10) = dPoints(1)
        dPoints(11) = dPoints(8)
        
        dPoints(12) = dPoints(0)
        dPoints(13) = dPoints(1)
        dPoints(14) = dPoints(2)
        
        Set oCrossStr(1) = oGeomFact.LineStrings3d.CreateByPoints(Nothing, 5, dPoints)
        
        'Cross section 2
        Set oCrossStr(2) = oGeomFact.LineStrings3d.CreateByPoints(Nothing, 5, dPoints)
        oVector.Set dLeverLength1 + 0.3 * (dLeverLength - dLeverLength1), 0, 0
        oTransMat.LoadIdentity
        oTransMat.Translate oVector
        oCrossStr(2).Transform oTransMat
        
        'Cross section 3
        dPoints(0) = dLeverLength1 + 0.5 * (dLeverLength - dLeverLength1)
        dPoints(1) = 0.5 * dLeverHeight1
        dPoints(2) = -0.25 * dLeverWidth
        
        dPoints(3) = dPoints(0)
        dPoints(4) = 0.9 * dLeverHeight1
        dPoints(5) = dPoints(2)
        
        dPoints(6) = dPoints(0)
        dPoints(7) = dPoints(4)
        dPoints(8) = 0.25 * dLeverWidth
        
        dPoints(9) = dPoints(0)
        dPoints(10) = dPoints(1)
        dPoints(11) = dPoints(8)
        
        dPoints(12) = dPoints(0)
        dPoints(13) = dPoints(1)
        dPoints(14) = dPoints(2)
        
        Set oCrossStr(3) = oGeomFact.LineStrings3d.CreateByPoints(Nothing, 5, dPoints)
        
        'Cross section 4
        Set oCrossStr(4) = oGeomFact.LineStrings3d.CreateByPoints(Nothing, 5, dPoints)
        oVector.Set 0.5 * (dLeverLength - dLeverLength1), 0, 0
        oTransMat.LoadIdentity
        oTransMat.Translate oVector
        oCrossStr(4).Transform oTransMat
        
        For iCount = 1 To 3
            Set objLeverPart2(iCount) = oGeomFact.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
                            oCrossStr(iCount), oCrossStr(iCount + 1), True)
        Next iCount
        
        oVector.Set 0, 0, 1
        oRotMat.LoadIdentity
        oRotMat.Rotate dInclAngle, oVector
        
        oVector.Set 0, parOffsetFrmValCen + 0.1 * dLeverHeight1, 0
        oTransMat.LoadIdentity
        oTransMat.Translate oVector
        
        'Set the Output
        For iCount = 1 To 3
            objLeverPart2(iCount).Transform oRotMat
            objLeverPart2(iCount).Transform oTransMat
            m_OutputColl.AddOutput "Lever_", objLeverPart2(iCount)
            Set objLeverPart2(iCount) = Nothing
        Next iCount
        
        For iCount = 1 To 4
            Set oCrossStr(iCount) = Nothing
        Next iCount
        
        'Create the Handle
        ReDim dPoints(0 To 14) As Double
        
        dHandleLength = (dLeverHeight2 + 0.1 * dLeverHeight1) / Sin(dHandleAngle)
        
        dPoints(0) = 0
        dPoints(1) = 0
        dPoints(2) = -0.2 * dLeverWidth
        
        dPoints(3) = dPoints(0)
        dPoints(4) = 0.2 * dLeverHeight1
        dPoints(5) = dPoints(2)
        
        dPoints(6) = dHandleLength
        dPoints(7) = dPoints(4)
        dPoints(8) = dPoints(2)
        
        dPoints(9) = dPoints(6)
        dPoints(10) = 0
        dPoints(11) = dPoints(2)
        
        dPoints(12) = dPoints(0)
        dPoints(13) = dPoints(1)
        dPoints(14) = dPoints(2)
        
        Set oLineStr = oGeomFact.LineStrings3d.CreateByPoints(Nothing, 5, dPoints)
        oVector.Set 0, 0, 1
        Set objHandle = PlaceProjection(m_OutputColl, oLineStr, oVector, 0.4 * dLeverWidth, True)
        
        oVector.Set 0, 0, -1
        oTransMat.LoadIdentity
        oTransMat.Rotate (dHandleAngle), oVector
        objHandle.Transform oTransMat
        
        oVector.Set dLeverLength1 + 0.1 * (dLeverLength - dLeverLength1), _
                        parOffsetFrmValCen + 0.1 * dLeverHeight1 + _
                        (dLeverLength1 + 0.1 * (dLeverLength - dLeverLength1)) * Sin(dInclAngle), 0
        oTransMat.LoadIdentity
        oTransMat.Translate oVector
        objHandle.Transform oTransMat
        
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objHandle
        Set objHandle = Nothing
    
    Case Lever_Inclined_Type1
        parOffsetFrmValCen = arrayOfInputs(2)
        parLeverLength = arrayOfInputs(3)
        parLeverHeight = arrayOfInputs(5)
        
        dLeverWidth = 0.15 * parLeverLength
        If CmpDblLessThan(dLeverWidth, 0.05) Then dLeverWidth = 0.05
        If CmpDblGreaterthan(dLeverWidth, 0.1) Then dLeverWidth = 0.1
    
        dLeverLength1 = 0.4 * parLeverLength
        
        dInclAngle = Atn(0.8 * parLeverHeight / parLeverLength)
        
        'Compute the LeverLength for the inclined and straight types
        dLeverLength = parLeverLength / Cos(dInclAngle)
        
        dHandleAngle = Atn((0.2 * parLeverHeight + dLeverLength1 * Tan(dInclAngle)) / _
                                    (parLeverLength - dLeverLength1))
                        
        'Create the Lever part 1
        oStPoint.Set 0, parOffsetFrmValCen, 0
        oEnPoint.Set 0, parOffsetFrmValCen + 0.3 * parLeverHeight, 0
        Set objLeverPart1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dLeverWidth, True)
        
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objLeverPart1
        Set objLeverPart1 = Nothing
            
        'Create the Lever part 2
        ReDim objLeverPart2(1) As Object
        ReDim oCrossStr(1 To 2) As IngrGeom3D.LineString3d
        ReDim dPoints(0 To 14) As Double
        
        For iCount = 1 To 2
            Set oCrossStr(iCount) = New LineString3d
        Next iCount
        
        'Cross secton 1
        dPoints(0) = 0
        dPoints(1) = -0.1 * parLeverHeight / Cos(dInclAngle)
        dPoints(2) = -dLeverWidth / 2
        
        dPoints(3) = dPoints(0)
        dPoints(4) = 0.1 * parLeverHeight / Cos(dInclAngle)
        dPoints(5) = dPoints(2)
        
        dPoints(6) = dPoints(0)
        dPoints(7) = dPoints(4)
        dPoints(8) = dLeverWidth / 2
        
        dPoints(9) = dPoints(0)
        dPoints(10) = dPoints(1)
        dPoints(11) = dPoints(8)
        
        dPoints(12) = dPoints(0)
        dPoints(13) = dPoints(1)
        dPoints(14) = dPoints(2)
        
        Set oCrossStr(1) = oGeomFact.LineStrings3d.CreateByPoints(Nothing, 5, dPoints)
        
        'Cross section 2
        dPoints(0) = dLeverLength
        dPoints(1) = -0.05 * parLeverHeight / Cos(dInclAngle)
        dPoints(2) = -0.25 * dLeverWidth
        
        dPoints(3) = dPoints(0)
        dPoints(4) = 0.05 * parLeverHeight / Cos(dInclAngle)
        dPoints(5) = dPoints(2)
        
        dPoints(6) = dPoints(0)
        dPoints(7) = dPoints(4)
        dPoints(8) = 0.25 * dLeverWidth
        
        dPoints(9) = dPoints(0)
        dPoints(10) = dPoints(1)
        dPoints(11) = dPoints(8)
        
        dPoints(12) = dPoints(0)
        dPoints(13) = dPoints(1)
        dPoints(14) = dPoints(2)
        
        Set oCrossStr(2) = oGeomFact.LineStrings3d.CreateByPoints(Nothing, 5, dPoints)
        
        oVector.Set 0, 0, 1
        oRotMat.LoadIdentity
        oRotMat.Rotate dInclAngle, oVector
        
        For iCount = 1 To 2
            oCrossStr(iCount).Transform oRotMat
        Next iCount
        
        Set objLeverPart2(1) = oGeomFact.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
                        oCrossStr(1), oCrossStr(2), True)

        oVector.Set 0, parOffsetFrmValCen + 0.2 * parLeverHeight, 0
        oTransMat.LoadIdentity
        oTransMat.Translate oVector

        'Set the Output
        objLeverPart2(1).Transform oTransMat
        m_OutputColl.AddOutput "Lever_", objLeverPart2(1)
        Set objLeverPart2(1) = Nothing
        
        For iCount = 1 To 2
            Set oCrossStr(iCount) = Nothing
        Next iCount
        
        'Create the Handle
        ReDim dPoints(0 To 14) As Double

        dHandleLength = (parLeverLength - dLeverLength1) / Cos(dHandleAngle)

        dPoints(0) = 0
        dPoints(1) = -0.05 * parLeverHeight
        dPoints(2) = -0.2 * dLeverWidth

        dPoints(3) = dPoints(0)
        dPoints(4) = 0.05 * parLeverHeight
        dPoints(5) = dPoints(2)

        dPoints(6) = dHandleLength
        dPoints(7) = dPoints(4)
        dPoints(8) = dPoints(2)

        dPoints(9) = dPoints(6)
        dPoints(10) = 0
        dPoints(11) = dPoints(2)

        dPoints(12) = dPoints(0)
        dPoints(13) = dPoints(1)
        dPoints(14) = dPoints(2)

        Set oLineStr = oGeomFact.LineStrings3d.CreateByPoints(Nothing, 5, dPoints)
        oVector.Set 0, 0, 1
        Set objHandle = PlaceProjection(m_OutputColl, oLineStr, oVector, 0.4 * dLeverWidth, True)

        oVector.Set 0, 0, -1
        oTransMat.LoadIdentity
        oTransMat.Rotate dHandleAngle, oVector
        objHandle.Transform oTransMat

        oVector.Set dLeverLength1, parOffsetFrmValCen + 0.2 * parLeverHeight + dLeverLength1 * Tan(dInclAngle), 0
        oTransMat.LoadIdentity
        oTransMat.Translate oVector
        objHandle.Transform oTransMat

        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objHandle
        Set objHandle = Nothing
    
    End Select
    'Remove the References
    Set oGeomFact = Nothing
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    Set oVector = Nothing
    Set oTransMat = Nothing
    Set oRotMat = Nothing
    Set oLineStr = Nothing
    Set objLeverPart1 = Nothing
    Set objHandle = Nothing
    
    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub


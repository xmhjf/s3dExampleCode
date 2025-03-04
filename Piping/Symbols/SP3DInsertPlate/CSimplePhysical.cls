VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSimplePhysical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   ProgId        :        SP3DInsertPlate.InsertPlate
'   File          :        CSimplePhysical.cls
'   Author        :        PK
'   Creation Date :        Wednesday 5, May 2007
'   Description   :        Insert Plate Penetration
'   Source        :        B-PP-1B.pdf
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   05.DEC.07       PK     CR-131488 Created the symbol
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages
Private PI As Double
Private Sub Class_Initialize()
Const METHOD = "Class_Initialize:"
On Error GoTo Errx
      PI = 4 * Atn(1)
    Exit Sub
Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)

    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt                            As PartFacelets.IJDPart
    Dim parOutDia                          As Double
    Dim parInDia                             As Double
    Dim parThickness                     As Double
    Dim parThicknessMultiplier      As Double
   
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parInDia = arrayOfInputs(2)
    parOutDia = arrayOfInputs(3)
    parThickness = arrayOfInputs(4)
    parThicknessMultiplier = arrayOfInputs(5)
    'Insert your code for Insert Plate
    Dim dLinePts(0 To 17)   As Double
    Dim oPlate                     As Object
    Dim oLineStr                  As IngrGeom3D.LineString3d
    Dim oAxis                      As AutoMath.DVector
    Dim oCentPos                As AutoMath.DPosition
    
    Set oCentPos = New AutoMath.DPosition
    Set oAxis = New AutoMath.DVector
    Set oLineStr = New IngrGeom3D.LineString3d
    
    dLinePts(0) = 0
    dLinePts(1) = parInDia / 2
    dLinePts(2) = 0
    
    dLinePts(3) = 0
    dLinePts(4) = parOutDia / 2
    dLinePts(5) = 0
    
    dLinePts(6) = parThickness
    dLinePts(7) = parOutDia / 2
    dLinePts(8) = 0
    
    Dim dTaperedLength               As Double
    Dim dTaperAngle                    As Double
    Dim dPlHalfLen                       As Double
    Dim dHorProTaperedLength    As Double
    Dim dVertProTaperedLength   As Double
    
    dPlHalfLen = (parOutDia - parInDia) / 2
    
    'Thickness:Tapered Length = 1:4
    dVertProTaperedLength = parThickness * (parThicknessMultiplier - 1)
    dHorProTaperedLength = 4 * dVertProTaperedLength
        
    If CmpDblGreaterthanOrEqualTo(dHorProTaperedLength, dPlHalfLen) Then
        dHorProTaperedLength = dPlHalfLen - 0.0001
    End If
    
    dLinePts(9) = parThickness * parThicknessMultiplier
    dLinePts(10) = parOutDia / 2 - dHorProTaperedLength
    dLinePts(11) = 0
    
    dLinePts(12) = parThickness * parThicknessMultiplier
    dLinePts(13) = parInDia / 2
    dLinePts(14) = 0
    
    dLinePts(15) = 0
    dLinePts(16) = parInDia / 2
    dLinePts(17) = 0
    
    Set oLineStr = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 6, dLinePts)
    oAxis.Set 1, 0, 0
    oCentPos.Set 0, 0, 0
    Set oPlate = PlaceRevolution(m_OutputColl, oLineStr, oAxis, oCentPos, 2 * PI, True)
    
    Dim oVec                As AutoMath.DVector
    Dim oTransMat       As AutoMath.DT4x4
    
    Set oVec = New AutoMath.DVector
    Set oTransMat = New AutoMath.DT4x4
    
    oVec.Set -parThickness / 2, 0, 0
    oTransMat.LoadIdentity
    oTransMat.Translate oVec
    oPlate.Transform oTransMat
    
    'Set the Output
    m_OutputColl.AddOutput "InsertPlate", oPlate
    Set oGeomFactory = Nothing
    Set oAxis = Nothing
    Set oCentPos = Nothing
    Set oPlate = Nothing
    Set oLineStr = Nothing
    Set oTransMat = Nothing
    Set oVec = Nothing
    
    Exit Sub
        
ErrorLabel:
   Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
End Sub

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
'   ProgId        :        SP3DSMPipeSleeveWSeal.SleeveWSeal
'   File          :        CSimplePhysical.cls
'   Author        :        PK
'   Creation Date :        Wednesday 6, May 2007
'   Description   :        Penetration Pipe Sleeve with Seal
'   Source        :        B-PP-13/13A.pdf
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   06.DEC.2007     PK     CR-131488 Created the symbol
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages
Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)

    Const METHOD = "run"
           
    On Error GoTo ErrorHandler
    
    Dim oPartFclt         As PartFacelets.IJDPart
    Dim parOutDia         As Double
    Dim parInDia          As Double
    Dim parSealWidth      As Double
    Dim parSleeveWidth    As Double
    
    Dim pipeDiam            As Double
    Dim flangeThick         As Double
    Dim sptOffset           As Double
    Dim flangeDiam          As Double
    Dim depth               As Double
    
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    
    'Set Inputs
    Set oPartFclt = arrayOfInputs(1)
    parInDia = arrayOfInputs(2)
    parOutDia = arrayOfInputs(3)
    parSleeveWidth = arrayOfInputs(4)
    parSealWidth = arrayOfInputs(5)
        
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    
    Dim dLinePts(0 To 20)   As Double
    Dim oPipeSleeveWtseal   As Object
    Dim oLineStr            As IngrGeom3D.LineString3d
    Dim oAxis               As AutoMath.DVector
    Dim oCentPos            As AutoMath.DPosition
    
    Set oCentPos = New AutoMath.DPosition
    Set oAxis = New AutoMath.DVector
    Set oLineStr = New IngrGeom3D.LineString3d
    
    dLinePts(0) = -parSleeveWidth / 2
    dLinePts(1) = pipeDiam / 2
    dLinePts(2) = 0
    
    dLinePts(3) = -parSleeveWidth / 2
    dLinePts(4) = parOutDia / 2
    dLinePts(5) = 0
    
    dLinePts(6) = parSleeveWidth / 2
    dLinePts(7) = parOutDia / 2
    dLinePts(8) = 0
    
    dLinePts(9) = parSleeveWidth / 2
    dLinePts(10) = parInDia / 2
    dLinePts(11) = 0
    
    dLinePts(12) = parSealWidth - parSleeveWidth / 2
    dLinePts(13) = parInDia / 2
    dLinePts(14) = 0
    
    dLinePts(15) = parSealWidth - parSleeveWidth / 2
    dLinePts(16) = pipeDiam / 2
    dLinePts(17) = 0
    
    dLinePts(18) = -parSleeveWidth / 2
    dLinePts(19) = pipeDiam / 2
    dLinePts(20) = 0
    
    Set oLineStr = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 7, dLinePts)
    oAxis.Set 1, 0, 0
    oCentPos.Set 0, 0, 0
    Set oPipeSleeveWtseal = PlaceRevolution(m_OutputColl, oLineStr, oAxis, oCentPos, 8 * Atn(1), True)
    
    'Set the Output
    m_OutputColl.AddOutput "PipeSleeve", oPipeSleeveWtseal
    Set oGeomFactory = Nothing
    Set oAxis = Nothing
    Set oCentPos = Nothing
    Set oPipeSleeveWtseal = Nothing
    Set oLineStr = Nothing
    Exit Sub
        
ErrorHandler:
   Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
End Sub

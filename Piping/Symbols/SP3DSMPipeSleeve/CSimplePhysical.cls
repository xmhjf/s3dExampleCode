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
'   Copyright (c) 2007-08, Intergraph Corporation. All rights reserved.
'
'   ProgId        :        SP3DSMPipeSleeve.SMPipeSleeve
'   File          :        CSimplePhysical.cls
'   Author        :        PK
'   Creation Date :        Wednesday 6, May 2007
'   Description   :        Penetration Pipe Sleeve
'   Source        :        B-PP-3.pdf
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   06.DEC.2007     PK     CR-131488 Created the symbol
'   11.AUG.2008     BS     CR-147688 Enhanced the symbol to Short Sleeve
'   11.AUG.2008     BS     CR-147689 Enhanced the symbol to  Long Sleeve
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages
Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)

    Const METHOD = "run"
    On Error GoTo ErrorLabel
        
    Dim oPartFclt           As PartFacelets.IJDPart
    Dim parInDia            As Double
    Dim parOutDia           As Double
    Dim parSleeveWidth1     As Double
    Dim parSleeveWidth2     As Double
    Dim parThickness     As Double
    Dim parHoleDia     As Double
    Dim parLength     As Double
    
    Dim oGeomFactory        As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    'parInDia = arrayOfInputs(2)
    parOutDia = arrayOfInputs(3)
    'parSleeveWidth1 = arrayOfInputs(4)
    'parSleeveWidth2 = arrayOfInputs(5)
    'parThickness = arrayOfInputs(6)
    'parHoleDia = arrayOfInputs(7)
    'parLength = arrayOfInputs(8)


    Dim dLinePts(0 To 14)   As Double
    Dim oPipeSleeve         As Object
    Dim oLineStr            As IngrGeom3D.LineString3d
    Dim oAxis               As AutoMath.DVector
    Dim oCentPos            As AutoMath.DPosition
    
    Set oCentPos = New AutoMath.DPosition
    Set oAxis = New AutoMath.DVector
    Set oLineStr = New IngrGeom3D.LineString3d
    
    Dim oPipeComponent As IJDPipeComponent
    Set oPipeComponent = oPartFclt
    Dim lPartDataBasis As Integer
    lPartDataBasis = oPipeComponent.PartDataBasis

    Select Case lPartDataBasis
        Case Is <= 1
            parInDia = arrayOfInputs(2)
            parSleeveWidth1 = arrayOfInputs(4)
            parSleeveWidth2 = arrayOfInputs(5)
            
            dLinePts(0) = parSleeveWidth1
            dLinePts(1) = parInDia / 2
            dLinePts(2) = 0
            
            dLinePts(3) = parSleeveWidth1
            dLinePts(4) = parOutDia / 2
            dLinePts(5) = 0
            
            dLinePts(6) = -parSleeveWidth2
            dLinePts(7) = parOutDia / 2
            dLinePts(8) = 0
            
            dLinePts(9) = -parSleeveWidth2
            dLinePts(10) = parInDia / 2
            dLinePts(11) = 0
            
            dLinePts(12) = parSleeveWidth1
            dLinePts(13) = parInDia / 2
            dLinePts(14) = 0
            
        Case 1065           'long sleeve
            parThickness = arrayOfInputs(6)
            parHoleDia = arrayOfInputs(7)
            parLength = arrayOfInputs(8)
            
            dLinePts(0) = parLength / 2
            dLinePts(1) = parOutDia / 2 - parThickness
            dLinePts(2) = 0
            
            dLinePts(3) = parLength / 2
            dLinePts(4) = parOutDia / 2
            dLinePts(5) = 0
            
            dLinePts(6) = -parLength / 2
            dLinePts(7) = parOutDia / 2
            dLinePts(8) = 0
            
            dLinePts(9) = -parLength / 2
            dLinePts(10) = parOutDia / 2 - parThickness
            dLinePts(11) = 0
            
            dLinePts(12) = parLength / 2
            dLinePts(13) = parOutDia / 2 - parThickness
            dLinePts(14) = 0

        
        Case 1066           'short sleeve
            parThickness = arrayOfInputs(6)
            parLength = arrayOfInputs(8)
            
            dLinePts(0) = parLength / 2
            dLinePts(1) = parOutDia / 2 - parThickness
            dLinePts(2) = 0
            
            dLinePts(3) = parLength / 2
            dLinePts(4) = parOutDia / 2
            dLinePts(5) = 0
            
            dLinePts(6) = -parLength / 2
            dLinePts(7) = parOutDia / 2
            dLinePts(8) = 0
            
            dLinePts(9) = -parLength / 2
            dLinePts(10) = parOutDia / 2 - parThickness
            dLinePts(11) = 0
            
            dLinePts(12) = parLength / 2
            dLinePts(13) = parOutDia / 2 - parThickness
            dLinePts(14) = 0
        
        Case Else
            GoTo ErrorLabel:      ' Invalid Part data basis.
        
    End Select

    Set oLineStr = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dLinePts)
    oAxis.Set 1, 0, 0
    oCentPos.Set 0, 0, 0
    Set oPipeSleeve = PlaceRevolution(m_OutputColl, oLineStr, oAxis, oCentPos, 8 * Atn(1), True)
    
    'Set the Output
    m_OutputColl.AddOutput "PipeSleeve", oPipeSleeve
    Set oGeomFactory = Nothing
    Set oAxis = Nothing
    Set oCentPos = Nothing
    Set oPipeSleeve = Nothing
    Set oLineStr = Nothing
    Set oPipeComponent = Nothing

    Exit Sub
        
ErrorLabel:
   Err.Raise Err.Number, Err.Source & " " & METHOD, Err.description, _
       Err.HelpFile, Err.HelpContext
End Sub




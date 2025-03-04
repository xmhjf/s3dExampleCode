VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CInsulation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2008, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   ProgID:         SP3DMoistureSeparator.MoistureSeparator
'   Author:         GL
'   Creation Date:  Tuesday, Oct 07 2008
'   Description:    The symbol is SP3DMoistureSeparator developed as per the CR-148509
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages



Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim negligible_thickness        As Double
    Dim parA                        As Double
    Dim parB                        As Double
    Dim parC                        As Double
    Dim parD                        As Double
    Dim parG                        As Double
    Dim parInsulationThickness      As Double
    Dim dSeperatorLength            As Double
    Dim pipeDiam                    As Double
    Dim flangeThick                 As Double
    Dim flangeDiam                  As Double
    Dim cptOffset                   As Double
    Dim depth                       As Double
    Dim iOutput                     As Double
    Dim dPlugPosition               As Double
    Dim dPlugCylDia                 As Double
    Dim dCircCenterY                As Double
    Dim dExtendedLength             As Double
    Dim dOvalPointY                 As Double
    Dim iCount                      As Integer

    Dim oPartFclt                   As PartFacelets.IJDPart
    Dim oStPoint                    As AutoMath.DPosition
    Dim oEnPoint                    As AutoMath.DPosition
    Dim geomFactory                 As IngrGeom3D.GeometryFactory
    Dim objCircle                   As IngrGeom3D.ComplexString3d
    Dim oCString                    As IngrGeom3D.ComplexString3d
    Dim ObjRuledBody                As IngrGeom3D.IJElements
    Dim objCrossSections            As IngrGeom3D.IJElements
    Dim objTraceCurves              As IngrGeom3D.IJElements
    Dim oArc                        As Object
    Dim oLine                       As Object
    Dim insCyl1                     As Object
    Dim insCyl2                     As Object
    Dim insCyl3                     As Object
    Dim insFlange1                  As Object
    Dim insFlange2                  As Object
    Dim insFlange3                  As Object
    Dim insPlugCyl                  As Object
    Dim insInsulatedBody            As Object
    Dim oArc1                       As Object
    Dim oLine1                      As Object
    Dim oLine_Trc1                  As Object
    Dim oLine_Trc2                  As Object
    Dim oLine_Trc3                  As Object
    Dim oLine_Trc4                  As Object
    Dim arcs                        As Collection
    Dim arcs1                       As Collection

    Set arcs = New Collection
    Set arcs1 = New Collection
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    Set geomFactory = New IngrGeom3D.GeometryFactory
    Set objCrossSections = New JObjectCollection
    Set objTraceCurves = New JObjectCollection

    ' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parA = arrayOfInputs(2)  'Face3 to Center
    parB = arrayOfInputs(3)  'Face3 to Strainer Top
    parC = arrayOfInputs(4)  'Body Width
    parD = arrayOfInputs(5)  'Face to Face
    parG = arrayOfInputs(6)  'Nozzle Offset 3
    parInsulationThickness = arrayOfInputs(7)

    dExtendedLength = 0.1 * parA
    dCircCenterY = parB - parA - (parC / 2)
    dOvalPointY = parA - dExtendedLength - (parC / 2)
    dSeperatorLength = 0.8 * parD
    dPlugPosition = 0.2 * dSeperatorLength
    dPlugCylDia = 1.6 * (parC / 2)
    negligible_thickness = 0.001
    iOutput = 0

    'Code for Insulation Cylinder1(At the port1)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, _
                       cptOffset, depth
    'Create insulation for flange-to-moisture separator body on left side
    'only if parD / 2 > (dSeperatorLength / 2) + parInsulationThickness
    If CmpDblGreaterthan((parD / 2), (dSeperatorLength / 2) + parInsulationThickness) Then
        oStPoint.Set -parD / 2, 0, 0
        oEnPoint.Set -(dSeperatorLength / 2) - parInsulationThickness, 0, 0
        Set insCyl1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                  (parC / 2) + (2 * parInsulationThickness), True)
        m_OutputColl.AddOutput "insCyl1", insCyl1
    End If
    
    'Flange
    If CmpDblGreaterthan(flangeDiam, pipeDiam) Then
        oStPoint.Set -parD / 2, 0, 0
        oEnPoint.Set -parD / 2 + flangeThick + parInsulationThickness, 0, 0
        Set insFlange1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                        flangeDiam + (2 * parInsulationThickness), True)
    Else
        oStPoint.Set -parD / 2, 0, 0
        oEnPoint.Set -parD / 2 + negligible_thickness, 0, 0
        Set insFlange1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                        (parC / 2) + (2 * parInsulationThickness), True)
    End If
    
    m_OutputColl.AddOutput "insFlange1", insFlange1
    Set insCyl1 = Nothing
    Set insFlange1 = Nothing
    
    'Code for Insulation Cylinder2(At the port2)
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, _
                       cptOffset, depth
    'Create insulation for flange-to-moisture separator body on Right side
    'only if parD / 2 > (dSeperatorLength / 2) + parInsulationThickness
    If CmpDblGreaterthan((parD / 2), (dSeperatorLength / 2) + parInsulationThickness) Then
        oStPoint.Set parD / 2, 0, 0
        oEnPoint.Set (dSeperatorLength / 2) + parInsulationThickness, 0, 0
        
        Set insCyl2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                    (parC / 2) + (2 * parInsulationThickness), True)
        m_OutputColl.AddOutput "insCyl2", insCyl2
    End If
    
    'Flange
    If CmpDblGreaterthan(flangeDiam, pipeDiam) Then
        oStPoint.Set parD / 2, 0, 0
        oEnPoint.Set parD / 2 - flangeThick - parInsulationThickness, 0, 0
        Set insFlange2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                        flangeDiam + (2 * parInsulationThickness), True)
    Else
        oStPoint.Set parD / 2, 0, 0
        oEnPoint.Set parD / 2 - negligible_thickness, 0, 0
        Set insFlange2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                         (parC / 2) + 2 * parInsulationThickness, True)
    End If
    
    m_OutputColl.AddOutput "insFlange2", insFlange2
    Set insCyl2 = Nothing
    Set insFlange2 = Nothing
    
    'Code for Insulation Cylinder3(At the port3)
    RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, _
                       cptOffset, depth
                       
    oStPoint.Set (parD / 2) - parG, -parA, 0
    oEnPoint.Set (parD / 2) - parG, 0, 0
    
    Set insCyl3 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                  (0.8 * (parC / 2)) + (2 * parInsulationThickness), True)
    m_OutputColl.AddOutput "insCyl3", insCyl3
    
    'Flange
    If CmpDblGreaterthan(flangeDiam, pipeDiam) Then
        oStPoint.Set (parD / 2) - parG, -parA, 0
        oEnPoint.Set (parD / 2) - parG, -parA + flangeThick + parInsulationThickness, 0
        Set insFlange3 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                         flangeDiam + 2 * parInsulationThickness, True)
    Else
        oStPoint.Set (parD / 2) - parG, -parA, 0
        oEnPoint.Set (parD / 2) - parG, -parA + negligible_thickness, 0
        Set insFlange3 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                        (0.8 * (parC / 2)) + (2 * parInsulationThickness), True)
    End If
    
    m_OutputColl.AddOutput "insFlange3", insFlange3
    Set insCyl3 = Nothing
    Set insFlange3 = Nothing
    
    Dim dSeparatorXCoor As Double
    If CmpDblGreaterthan((dSeperatorLength / 2) + parInsulationThickness, (parD / 2)) Then
        dSeparatorXCoor = parD / 2
    Else
        dSeparatorXCoor = (dSeperatorLength / 2) + parInsulationThickness
    End If
    
    'Code for Output3(Insulation Seperator Body)
    'Circle Creation
    Set oArc = geomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, _
              -dSeparatorXCoor, dCircCenterY, 0, _
              -dSeparatorXCoor, dCircCenterY, _
              -(parC / 2) - parInsulationThickness, -dSeparatorXCoor, _
              dCircCenterY + (parC / 2) + parInsulationThickness, 0)
    arcs.Add oArc
    
    Set oArc = geomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, _
               -dSeparatorXCoor, _
              dCircCenterY, 0, -dSeparatorXCoor, _
              dCircCenterY + (parC / 2) + parInsulationThickness, 0, _
              -dSeparatorXCoor, dCircCenterY, _
              (parC / 2) + parInsulationThickness)
    arcs.Add oArc
    
    Set oArc = geomFactory.Arcs3d.CreateByCtrNormStartEnd(Nothing, -dSeparatorXCoor, _
               dCircCenterY, 0, 1, 0, 0, -dSeparatorXCoor, _
               dCircCenterY, (parC / 2) + parInsulationThickness, -dSeparatorXCoor, _
               dCircCenterY - (parC / 2) - parInsulationThickness, 0)
    arcs.Add oArc
    
    Set oArc = geomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, -dSeparatorXCoor, _
               dCircCenterY, 0, -dSeparatorXCoor, _
               dCircCenterY - (parC / 2) - parInsulationThickness, 0, _
               -dSeparatorXCoor, dCircCenterY, -(parC / 2) - parInsulationThickness)
    arcs.Add oArc
    
    oStPoint.Set -dSeparatorXCoor, dCircCenterY, -(parC / 2) - parInsulationThickness
    Set objCircle = PlaceTrCString(oStPoint, arcs)
    objCrossSections.Add objCircle
    
    'Oval Creation
    Set oArc1 = geomFactory.Arcs3d.CreateBy3Points(Nothing, dSeparatorXCoor, _
                dCircCenterY, -parC / 2 - parInsulationThickness, dSeparatorXCoor, _
                dCircCenterY + (parC / 2) + parInsulationThickness, 0, dSeparatorXCoor, _
                dCircCenterY, (parC / 2) + parInsulationThickness)
    arcs1.Add oArc1
    
    Set oLine1 = geomFactory.Lines3d.CreateBy2Points(Nothing, dSeparatorXCoor, _
                dCircCenterY, (parC / 2) + parInsulationThickness, dSeparatorXCoor, _
                -dOvalPointY, (parC / 2) + parInsulationThickness)
    arcs1.Add oLine1
    
    Set oArc1 = geomFactory.Arcs3d.CreateBy3Points(Nothing, dSeparatorXCoor, _
                -dOvalPointY, (parC / 2) + parInsulationThickness, dSeparatorXCoor, _
                -dOvalPointY - (parC / 2) - parInsulationThickness, 0, dSeparatorXCoor, -dOvalPointY, _
                -parC / 2 - parInsulationThickness)
    arcs1.Add oArc1
    
    Set oLine1 = geomFactory.Lines3d.CreateBy2Points(Nothing, dSeparatorXCoor, _
                -dOvalPointY, -(parC / 2) - parInsulationThickness, dSeparatorXCoor, _
                dCircCenterY, -(parC / 2) - parInsulationThickness)
    arcs1.Add oLine1
    
    oStPoint.Set dSeparatorXCoor, dCircCenterY, _
                 -(parC / 2) - parInsulationThickness
    Set oCString = PlaceTrCString(oStPoint, arcs1)
    objCrossSections.Add oCString
    
    'Trace Curves

    Set oLine_Trc1 = geomFactory.Lines3d.CreateBy2Points(Nothing, -dSeparatorXCoor, _
                    dCircCenterY, (-parC / 2) - parInsulationThickness, dSeparatorXCoor, _
                    dCircCenterY, (-parC / 2) - parInsulationThickness)
    objTraceCurves.Add oLine_Trc1
    
    Set oLine_Trc2 = geomFactory.Lines3d.CreateBy2Points(Nothing, -dSeparatorXCoor, _
                    dCircCenterY + (parC / 2) + parInsulationThickness, 0, _
                    dSeparatorXCoor, dCircCenterY + _
                    (parC / 2) + parInsulationThickness, 0)
    objTraceCurves.Add oLine_Trc2
    
    Set oLine_Trc3 = geomFactory.Lines3d.CreateBy2Points(Nothing, -dSeparatorXCoor, _
                    dCircCenterY, (parC / 2) + parInsulationThickness, dSeparatorXCoor, _
                    dCircCenterY, (parC / 2) + parInsulationThickness)
    objTraceCurves.Add oLine_Trc3
    
    Set oLine_Trc4 = geomFactory.Lines3d.CreateBy2Points(Nothing, -dSeparatorXCoor, _
                     dCircCenterY - (parC / 2) - parInsulationThickness, 0, _
                     dSeparatorXCoor, -dOvalPointY - (parC / 2) _
                     - parInsulationThickness, 0)
    objTraceCurves.Add oLine_Trc4

    Set ObjRuledBody = geomFactory.GeometryServices.CreateBySkinning(m_OutputColl.ResourceManager, _
                        objTraceCurves, objCrossSections, True)
    Set insInsulatedBody = ObjRuledBody
    
    iOutput = iOutput + 1
    For Each insInsulatedBody In ObjRuledBody
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), insInsulatedBody
    Next insInsulatedBody

    Set insInsulatedBody = Nothing
    Set objCircle = Nothing
    Set oCString = Nothing
    Set ObjRuledBody = Nothing
    Set geomFactory = Nothing
    Set oArc = Nothing
    Set oLine = Nothing
    Set oArc1 = Nothing
    Set oLine1 = Nothing
    Set oLine_Trc1 = Nothing
    Set oLine_Trc2 = Nothing
    Set oLine_Trc3 = Nothing
    Set oLine_Trc4 = Nothing
    
    For iCount = 1 To objTraceCurves.Count
        objTraceCurves.Remove 1
    Next iCount
    Set objTraceCurves = Nothing
    
    For iCount = 1 To objCrossSections.Count
        objCrossSections.Remove 1
    Next iCount
    Set objCrossSections = Nothing
    
    For iCount = 1 To arcs.Count
        arcs.Remove 1
    Next iCount
    Set arcs = Nothing
    
    For iCount = 1 To arcs1.Count
        arcs1.Remove 1
    Next iCount
    Set arcs1 = Nothing

    'Code for Output4(Plug)
    oStPoint.Set -dPlugPosition, 0, 0
    oEnPoint.Set -dPlugPosition, (1.1 * (parB - parA)) + _
                (0.25 * parC / 2) + parInsulationThickness, 0
    Set insPlugCyl = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                    dPlugCylDia + 2 * parInsulationThickness, True)
    
    m_OutputColl.AddOutput "insPlugCyl", insPlugCyl
    Set insPlugCyl = Nothing
    Set oStPoint = Nothing
    Set oEnPoint = Nothing

Exit Sub

ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.description, _
              Err.HelpFile, Err.HelpContext

End Sub

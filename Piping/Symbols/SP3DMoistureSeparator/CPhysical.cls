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
'   ProgID:         SP3DMoistureSeparator.MoistureSeparator
'   Author:         GL
'   Creation Date:  Tuesday, Oct 07 2008
'   Description:    The symbol is SP3DMoistureSeparator developed as per the CR-148509
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Physical:" 'Used for error messages
Private PI       As Double

Private Sub Class_Initialize()
      PI = 4 * Atn(1)
End Sub
Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim parA                    As Double
    Dim parB                    As Double
    Dim parC                    As Double
    Dim parD                    As Double
    Dim parG                    As Double
    Dim parInsulationThickness  As Double
    Dim dSeperatorLength        As Double
    Dim pipeDiam                As Double
    Dim flangeThick             As Double
    Dim flangeDiam              As Double
    Dim cptOffset               As Double
    Dim depth                   As Double
    Dim iOutput                 As Double
    Dim dPlugPosition           As Double
    Dim dPlugCylDia             As Double
    Dim dPlugGasketDia          As Double
    Dim dplugSideLength         As Double
    Dim dCircCenterY            As Double
    Dim dExtendedLength         As Double
    Dim dOvalPointY             As Double
    Dim iCount                  As Integer
       
    Dim oPartFclt               As PartFacelets.IJDPart
    Dim oStPoint                As AutoMath.DPosition
    Dim oEnPoint                As AutoMath.DPosition
    Dim geomFactory             As IngrGeom3D.GeometryFactory
    Dim objCircle               As IngrGeom3D.ComplexString3d
    Dim oCString                As IngrGeom3D.ComplexString3d
    Dim ObjRuledBody            As IngrGeom3D.IJElements
    Dim objCrossSections        As IngrGeom3D.IJElements
    Dim objTraceCurves          As IngrGeom3D.IJElements
    Dim oVec                    As AutoMath.DVector
    Dim oTransMat               As AutoMath.DT4x4
    Dim objCyl1                 As Object
    Dim objCyl2                 As Object
    Dim objSeperatorBody        As Object
    Dim objPlug                 As Object
    Dim objNozzle1              As Object
    Dim objNozzle2              As Object
    Dim objNozzle3              As Object
    Dim objInsulatedBody        As Object
    Dim oArc                    As Object
    Dim oLine                   As Object
    Dim oArc1                   As Object
    Dim oLine1                  As Object
    Dim oLine_Trc1              As Object
    Dim oLine_Trc2              As Object
    Dim oLine_Trc3              As Object
    Dim oLine_Trc4              As Object
    Dim plugCyl                 As Object
    Dim plugGasket              As Object
    Dim arcs                    As Collection
    Dim arcs1                   As Collection
    
    Set arcs = New Collection
    Set arcs1 = New Collection
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    Set geomFactory = New IngrGeom3D.GeometryFactory
    Set objCrossSections = New JObjectCollection
    Set objTraceCurves = New JObjectCollection
    Set oTransMat = New DT4x4
    Set oVec = New AutoMath.DVector
    
    ' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parA = arrayOfInputs(2)  'Face3 to Center
    parB = arrayOfInputs(3)  'Face3 to Strainer Top
    parC = arrayOfInputs(4)  'Body Width
    parD = arrayOfInputs(5)  'Face to Face
    parG = arrayOfInputs(6)  'Nozzle Offset 3
    parInsulationThickness = arrayOfInputs(7)
    
    iOutput = 0
    
    dExtendedLength = 0.1 * parA '(Assuming dExtendedLength to be 0.1 times parA)
    dCircCenterY = parB - parA - (parC / 2)
    dOvalPointY = parA - dExtendedLength - (parC / 2)
    dSeperatorLength = 0.8 * parD
    dPlugPosition = 0.2 * dSeperatorLength
    dPlugCylDia = 1.6 * (parC / 2)
    dPlugGasketDia = 0.8 * dPlugCylDia
    dplugSideLength = 0.5 * dPlugGasketDia
      
    'Code for Output1(Cylinder1 at the inline ports)
    oStPoint.Set -parD / 2, 0, 0
    oEnPoint.Set parD / 2, 0, 0
    
    Set objCyl1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, (parC / 2), True)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCyl1
    Set objCyl1 = Nothing

    'Code for Output2(Cylinder2 at the third port)
    RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, _
                       cptOffset, depth
    oStPoint.Set (parD / 2) - parG, -parA, 0
    oEnPoint.Set (parD / 2) - parG, 0, 0
    
    Set objCyl2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 0.8 * (parC / 2), True)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCyl2
    Set objCyl2 = Nothing
    
    'Code for Output3(Seperator Body)

    'Circle Creation

    Set oArc = geomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, -dSeperatorLength / 2, _
               dCircCenterY, 0, -dSeperatorLength / 2, dCircCenterY, -parC / 2, _
               -dSeperatorLength / 2, dCircCenterY + (parC / 2), 0)
    arcs.Add oArc
    
    Set oArc = geomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, -dSeperatorLength / 2, _
               dCircCenterY, 0, -dSeperatorLength / 2, dCircCenterY + (parC / 2), 0, _
               -dSeperatorLength / 2, dCircCenterY, parC / 2)
    arcs.Add oArc
    
    Set oArc = geomFactory.Arcs3d.CreateByCtrNormStartEnd(Nothing, -dSeperatorLength / 2, _
               dCircCenterY, 0, 1, 0, 0, -dSeperatorLength / 2, dCircCenterY, parC / 2, _
               -dSeperatorLength / 2, dCircCenterY - (parC / 2), 0)
    arcs.Add oArc
    
    Set oArc = geomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, -dSeperatorLength / 2, _
               dCircCenterY, 0, -dSeperatorLength / 2, dCircCenterY - (parC / 2), 0, _
               -dSeperatorLength / 2, dCircCenterY, -parC / 2)
    arcs.Add oArc
    
    oStPoint.Set -dSeperatorLength / 2, dCircCenterY, -parC / 2
    Set objCircle = PlaceTrCString(oStPoint, arcs)
    objCrossSections.Add objCircle
   
    'Oval Creation
    Set oArc1 = geomFactory.Arcs3d.CreateBy3Points(Nothing, dSeperatorLength / 2, _
                dCircCenterY, -parC / 2, dSeperatorLength / 2, dCircCenterY + (parC / 2), 0, _
                dSeperatorLength / 2, dCircCenterY, parC / 2)
    arcs1.Add oArc1
    
    Set oLine1 = geomFactory.Lines3d.CreateBy2Points(Nothing, dSeperatorLength / 2, _
                 dCircCenterY, parC / 2, dSeperatorLength / 2, -dOvalPointY, parC / 2)
    arcs1.Add oLine1
    
    Set oArc1 = geomFactory.Arcs3d.CreateBy3Points(Nothing, dSeperatorLength / 2, _
                -dOvalPointY, parC / 2, dSeperatorLength / 2, -dOvalPointY - (parC / 2), 0, _
                dSeperatorLength / 2, -dOvalPointY, -parC / 2)
    arcs1.Add oArc1
    
    Set oLine1 = geomFactory.Lines3d.CreateBy2Points(Nothing, dSeperatorLength / 2, _
                -dOvalPointY, -parC / 2, dSeperatorLength / 2, dCircCenterY, -parC / 2)
    arcs1.Add oLine1
    
    oStPoint.Set dSeperatorLength / 2, dCircCenterY, -parC / 2
    Set oCString = PlaceTrCString(oStPoint, arcs1)
    objCrossSections.Add oCString
    
    'Trace Curves
    Set oLine_Trc1 = geomFactory.Lines3d.CreateBy2Points(Nothing, -dSeperatorLength / 2, _
                     dCircCenterY, (-parC / 2), dSeperatorLength / 2, dCircCenterY, _
                     (-parC / 2))
    objTraceCurves.Add oLine_Trc1
    
    Set oLine_Trc2 = geomFactory.Lines3d.CreateBy2Points(Nothing, -dSeperatorLength / 2, _
                     dCircCenterY + (parC / 2), 0, dSeperatorLength / 2, dCircCenterY + (parC / 2), 0)
    objTraceCurves.Add oLine_Trc2
    
    Set oLine_Trc3 = geomFactory.Lines3d.CreateBy2Points(Nothing, -dSeperatorLength / 2, _
                     dCircCenterY, parC / 2, dSeperatorLength / 2, dCircCenterY, parC / 2)
    objTraceCurves.Add oLine_Trc3
    
    Set oLine_Trc4 = geomFactory.Lines3d.CreateBy2Points(Nothing, -dSeperatorLength / 2, _
                    dCircCenterY - (parC / 2), 0, dSeperatorLength / 2, -dOvalPointY - (parC / 2), 0)
    objTraceCurves.Add oLine_Trc4

    Set ObjRuledBody = geomFactory.GeometryServices.CreateBySkinning(m_OutputColl.ResourceManager, _
                       objTraceCurves, objCrossSections, True)
    Set objSeperatorBody = ObjRuledBody
    
    iOutput = iOutput + 1
    For Each objSeperatorBody In ObjRuledBody
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objSeperatorBody
    Next objSeperatorBody

    Set objSeperatorBody = Nothing
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
    
    'Plug Cylinder
    oStPoint.Set -dPlugPosition, 0, 0
    oEnPoint.Set -dPlugPosition, parB - parA, 0
    
    Set plugCyl = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dPlugCylDia, True)
    m_OutputColl.AddOutput "plugCyl", plugCyl
    Set plugCyl = Nothing
    
    'Plug Gasket
    oStPoint.Set -dPlugPosition, parB - parA, 0
    oEnPoint.Set -dPlugPosition, 1.1 * (parB - parA), 0
    
    Set plugGasket = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dPlugGasketDia, True)
    m_OutputColl.AddOutput "plugGasket", plugGasket
    Set plugGasket = Nothing
    
    'Hexagonal Plug
    Set objPlug = PlaceNnagon(m_OutputColl, 6, dplugSideLength, 0.25 * (parC / 2), True)
    
    oTransMat.LoadIdentity
    oVec.Set 0, 0, 1
    oTransMat.Rotate PI / 2, oVec
    objPlug.Transform oTransMat
    
    oTransMat.LoadIdentity
    oVec.Set -dPlugPosition, 1.1 * (parB - parA), 0
    oTransMat.Translate oVec
    objPlug.Transform oTransMat
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objPlug
    Set objPlug = Nothing
    
    Set oVec = Nothing
    Set oTransMat = Nothing
      
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'NOZZLE
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Dim oDirVec As AutoMath.DVector
    Set oDirVec = New AutoMath.DVector

    'Nozzle1
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, _
                       cptOffset, depth
    
    oStPoint.Set (-parD / 2) + depth - cptOffset, 0, 0
    oDirVec.Set -1, 0, 0
    
    Set objNozzle1 = CreateNozzle(1, oPartFclt, m_OutputColl, oDirVec, oStPoint)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle1
    Set objNozzle1 = Nothing

    'Nozzle2
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, _
                       cptOffset, depth
    
    oStPoint.Set (parD / 2) - depth + cptOffset, 0, 0
    oDirVec.Set 1, 0, 0
    
    Set objNozzle2 = CreateNozzle(2, oPartFclt, m_OutputColl, oDirVec, oStPoint)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle2
    Set objNozzle2 = Nothing

    'Nozzle 3
    RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, _
                       cptOffset, depth
    
    oDirVec.Set 0, -1, 0
    oStPoint.Set (parD / 2) - parG, -parA + depth - cptOffset, 0
    
    Set objNozzle3 = CreateNozzle(3, oPartFclt, m_OutputColl, oDirVec, oStPoint)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle3
    Set objNozzle3 = Nothing
    
    Set oDirVec = Nothing
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    
Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub

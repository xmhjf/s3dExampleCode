Attribute VB_Name = "EqpFndCommon"
Option Explicit

'******************************************************************
' Copyright (C) 2006, Intergraph Corporation. All rights reserved.
'
'File
'    EqpFndCommon.cls
'
'Author
'       1-Mar-03        SS
'
'Description
'
'Notes
'
'History:
'   09-May-06   JMS DI#97751 - All lines are constructed through the
'                   geometry factory instead of being "new"ed
' 04-Aug-06 AS  TR#99968 Added new function to handle EqpFnd Migration
'
'   09-Aug-06   SS  DI#103385 - modified getproxy call in AddMaterialRelationShip to use
'                   PreferProxyByAliasName = true to get named proxy for material rather than
'                   the proxy itself and use it to establish StructEntityMaterial relation.
'
'   31-Aug-07   SS  TR#121865 - Migrate() react to mirror copy of pump with eqp fndn
'
'  02-Nov-07  SS    TR#128057 - A delta of .127m and .25m was being added to fndn port
'                   calculations based on standard pump foundation port. This will lead eqp fndns with diff dimensions
'                   when fndn ports are being added manually in eqp environment. Removed them.
'
'  10-Jan-08   SS   TR#122653 - user should not be able to select bounding surfaces
'                   that are above the equipment foundation port plane
'
'  24-Jun-09    WR  TR-145118 - Added new method CreateSolidAsPlanes.
'  15-Apr-14    RRK TR-CP-250882 Made change to CalculateHeightFromOriginToSupporingPlane method to solve Recorded Exception Minidump
'*******************************************************************

Private Const MODULE = "EqpFndCommon"
Public Const IJStructMaterial = "{93F52983-9504-11D4-9D40-00105AA5BAEB}"
Public Const IJWeightCG = "{DC284B37-5A00-11D2-BE2B-0800364AA003}"
Public Const IJDAttributes = "{B25FD387-CFEB-11D1-850B-080036DE8E03}"
Private Const MODELDATABASE = "Model"
Private Const CATALOGDATABASE = "Catalog"
Private oLocalizer As IJLocalizer

Private Const STRUCTMEMPARTPROGID = "SPSMembers.SPSMemberPartPrismatic"

Public Const IJFullObject = "{bcbfb3c0-98c2-11d1-93de-08003670a902}"
Public Const IJEqpFoundationPort = "{11120746-CEAB-11D3-B680-00A0C9AC60E9}"
Public Const IJGeometry = "{96eb9676-6530-11d1-977f-080036754203}"
Public Const IJPlane = "{4317C6B3-D265-11D1-9558-0060973D4824}"
Public Const AssemblyMembers1RelationshipCLSID = "{45E4020F-F8D8-47A1-9B00-C9570C1E0B17}"
Public Const ISPSLogicalAxis = "{37D0751F-3260-4457-A55C-56CE788CB093}"
Public Const ISPSAxisRotation = "{56CCD3A5-B756-4ab0-9C68-1F586D1E7C66}"

Private Const CROSSSEC_IFACE = "ISTRUCTCrossSectionDimensions"
Private Const CROSSSEC_DEPTH_ATTRNAME = "Depth"
Private Const CROSSSEC_WIDTH_ATTRNAME = "Width"

Public Const E_FAIL = -2147467259
Public Const PI = 3.14159265358979
Public Const INFINITY = 1000000

Private Const BLOCKANDSLABFNDASM_IFACE = "IJUASPSBlockAndSlabFndnAsm"
Private Const BLOCKFND_IFACE = "IJUASPSBlockFndn"
Private Const BLOCK_COMP_NAME = "BlockComponent"
Private Const BLOCKLENGTH_ATTRNAME = "BlockLength"
Private Const BLOCKWIDTH_ATTRNAME = "BlockWidth"
Private Const BLOCKHT_ATTRNAME = "BlockHeight"
Private Const BLOCKSIZEBYRULE_ATTRNAME = "IsBlockSizeDrivenByRule"
Private Const BLOCKEDGECLEAR_ATTRNAME = "BlockEdgeClearance"

Private Const SLABFND_IFACE = "IJUASPSSlabFndn"
Private Const SLABLENGTH_ATTRNAME = "SlabLength"
Private Const SLABWIDTH_ATTRNAME = "SlabWidth"
Private Const SLABHT_ATTRNAME = "SlabHeight"
Private Const SLABSIZEBYRULE_ATTRNAME = "IsSlabSizeDrivenByRule"
Private Const SLABEDGECLEAR_ATTRNAME = "SlabEdgeClearance"
Public Const SYMBOL_REP_NAME = "DetailPhysical"
Public Const IJOUTPUTCOLLECTION = "{15916CAF-6CB5-11D1-A655-00A0C98D7F13}"
Public Const TO_OUTPUTS = "toOutputs"


Public Function GetResourceMgr() As IUnknown
Const MT = "GetResourceMgr"
On Error GoTo ErrorHandler

     Dim jContext As IJContext
     Dim oDBTypeConfig As IJDBTypeConfiguration
     Dim oConnectMiddle As IJDAccessMiddle
     Dim strModelDBID As String
     
     'Get the middle context
     Set jContext = GetJContext()
     
     Set oDBTypeConfig = jContext.GetService("DBTypeConfiguration")
     Set oConnectMiddle = jContext.GetService("ConnectMiddle")
     
     strModelDBID = oDBTypeConfig.get_DataBaseFromDBType(MODELDATABASE)
     Set GetResourceMgr = oConnectMiddle.GetResourceManager(strModelDBID)
 
    Exit Function
ErrorHandler: HandleError MODULE, MT
End Function

Public Function GetDefinition(ByVal name As String) As Object
Const MT = "GetDefinition"
On Error GoTo ErrorHandler

    Dim iPom As IJDPOM
    Set iPom = GetCatalogResourceManager()
    Set GetDefinition = iPom.GetObject(name)
    Exit Function
    
ErrorHandler:
    HandleError MODULE, MT
End Function


Public Function GetRefCollection(pSO As IJSmartOccurrence) As IJDReferencesCollection
Const METHOD = "GetRefCollection"
On Error GoTo ErrorHandler

'IJSmartOccurence GUID
'{A2A655C0-E2F5-11D4-9825-00104BD1CC25}
  
     Dim pRelationHelper As IMSRelation.DRelationHelper
     Dim pCollectionHelper As IMSRelation.DCollectionHelper
     Set pRelationHelper = pSO
     On Error Resume Next
     Set pCollectionHelper = pRelationHelper.CollectionRelations("{A2A655C0-E2F5-11D4-9825-00104BD1CC25}", "toArgs_O")
     On Error Resume Next
     Set GetRefCollection = pCollectionHelper.Item("RC")
      
    Exit Function
Errx:
ErrorHandler:
    HandleError MODULE, METHOD
End Function

Public Sub ConnectSmartOccurrence(pSO As IJSmartOccurrence, pRefColl As IJDReferencesCollection)
Const MT = "ConnectSmartOccurrence"
On Error GoTo ErrorHandler
  
   'connect the reference collection to the smart occurrence
    Dim pRelationHelper As IMSRelation.DRelationHelper
    Dim pCollectionHelper As IMSRelation.DCollectionHelper
    Dim pRelationshipHelper As DRelationshipHelper
    Dim pRevision As IJRevision
    
    Set pRelationHelper = pSO
    Set pCollectionHelper = pRelationHelper.CollectionRelations("{A2A655C0-E2F5-11D4-9825-00104BD1CC25}", "toArgs_O")
    pCollectionHelper.Add pRefColl, "RC", pRelationshipHelper
    Set pRevision = New JRevision
    pRevision.AddRelationship pRelationshipHelper
  
    Exit Sub
ErrorHandler: HandleError MODULE, MT
End Sub

Public Function GetCAODefAttribute(ByVal pMemberDescription As IJDMemberDescription, _
                              InterfaceName As String, AttributeName As String) As String
Const METHOD = "GetCAODefAttribute"
On Error GoTo ErrorHandler
     
    Dim oSmartOcc As IJSmartOccurrence
    Dim oSmartItem As IJSmartItem
    Dim Attrs As IJDAttributes
    Dim SlabComp As String
    Set oSmartOcc = pMemberDescription.CAO
    Set oSmartItem = oSmartOcc.ItemObject
    Set Attrs = oSmartItem
    GetCAODefAttribute = Attrs.CollectionOfAttributes(InterfaceName).Item(AttributeName).Value
    Exit Function
    
ErrorHandler:
    HandleError MODULE, METHOD
End Function

Public Sub SetTransform(ByRef tmx As Automath.DT4x4, _
                        origin As Automath.DPosition, _
                        xVec As Automath.DVector, _
                        yVec As Automath.DVector)
Const METHOD = "SetTransform"
On Error GoTo ErrHandler

    Dim xxVec As Automath.DVector, yyVec As Automath.DVector, zzVec As Automath.DVector
            
    Set xxVec = xVec.Clone
    Set yyVec = yVec.Clone
    xxVec.Length = 1#
    yyVec.Length = 1#
    Set zzVec = xxVec.Cross(yyVec) 'New DVector
    'zzVec.Set 0, 0, 1
    zzVec.Length = 1#
    Set yyVec = Nothing
    Set yyVec = zzVec.Cross(xxVec)  ' in case x and y are not perp

    tmx.IndexValue(0) = xxVec.x
    tmx.IndexValue(1) = xxVec.y
    tmx.IndexValue(2) = xxVec.z
    tmx.IndexValue(3) = 0#
    
    tmx.IndexValue(4) = yyVec.x
    tmx.IndexValue(5) = yyVec.y
    tmx.IndexValue(6) = yyVec.z
    tmx.IndexValue(7) = 0#

    tmx.IndexValue(8) = zzVec.x
    tmx.IndexValue(9) = zzVec.y
    tmx.IndexValue(10) = zzVec.z
    tmx.IndexValue(11) = 0#
    
    
    tmx.IndexValue(12) = origin.x
    tmx.IndexValue(13) = origin.y
    tmx.IndexValue(14) = origin.z
    tmx.IndexValue(15) = 1#
    
    Set xxVec = Nothing
    Set yyVec = Nothing
    Set zzVec = Nothing
    
    Exit Sub
    
ErrHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub GetTransformAndHoles(holes() As DPosition, _
                                          NumberOfHoles As Integer, _
                                          FoundationPorts As IJElements, _
                                          FirstTrans As IJDT4x4)
Const METHOD = "GetTransformAndHoles"
On Error GoTo ErrorHandler

'    The foundation Ports can come in different orders.  So basis will be different.
    Dim oPort As IJEqpFoundationPort
    Dim oPointOrPort As Object
    Dim Trans As IJDT4x4
    Dim tempHoles() As Variant
    Dim FirstPort As Boolean
    Dim a As Integer, CurrentHole As Integer, PortNumber As Integer
    Dim xAxis As DVector, yAxis As DVector, zaxis As DVector
    Dim oX#, oY#, oZ#, xax#, xay#, xaz#, zax#, zay#, zaz#, yax#, yay#, yaz#
    Dim bPointOption As Boolean
    Set xAxis = New DVector
    Set yAxis = New DVector
    Set zaxis = New DVector
    
    Set FirstTrans = New DT4x4
    FirstTrans.LoadIdentity
    Set Trans = New DT4x4
    Trans.LoadIdentity
    FirstPort = True
    CurrentHole = 1
    
    ReDim holes(0) As DPosition
    'Set holes(FoundationPorts.count) = New DPosition
    For Each oPointOrPort In FoundationPorts
        If TypeOf oPointOrPort Is IJEqpFoundationPort Then
            Set oPort = oPointOrPort
            PortNumber = PortNumber + 1
            If oPort.NumberOfHoles > 0 Then
                
                ReDim Preserve holes(UBound(holes) + oPort.NumberOfHoles) As DPosition
                Call oPort.GetHoles(tempHoles())
                
                oPort.GetCS oX, oY, oZ, xax, xay, xaz, zax, zay, zaz
                xAxis.Set xax, xay, xaz
                zaxis.Set zax, zay, zaz
                Set yAxis = zaxis.Cross(xAxis)
                
                If FirstPort Then
                    'Load holes into Holes()
                    FirstPort = False
                    For a = LBound(tempHoles) To UBound(tempHoles)
                        Set holes(CurrentHole) = New DPosition
                        holes(CurrentHole).Set tempHoles(a, 1), tempHoles(a, 2), 0
                        CurrentHole = CurrentHole + 1
                    Next a
        
                    FirstTrans.IndexValue(0) = xax
                    FirstTrans.IndexValue(1) = xay
                    FirstTrans.IndexValue(2) = xaz
                    FirstTrans.IndexValue(3) = 0
                    FirstTrans.IndexValue(4) = yAxis.x
                    FirstTrans.IndexValue(5) = yAxis.y
                    FirstTrans.IndexValue(6) = yAxis.z
                    FirstTrans.IndexValue(7) = 0
                    FirstTrans.IndexValue(8) = zax
                    FirstTrans.IndexValue(9) = zay
                    FirstTrans.IndexValue(10) = zaz
                    FirstTrans.IndexValue(11) = 0
                    FirstTrans.IndexValue(12) = oX
                    FirstTrans.IndexValue(13) = oY
                    FirstTrans.IndexValue(14) = oZ
                    FirstTrans.IndexValue(15) = 1
                    FirstTrans.Invert
                Else
        
                    'convert to be in the same coordinate system as the first port
                    Trans.IndexValue(0) = xax
                    Trans.IndexValue(1) = xay
                    Trans.IndexValue(2) = xaz
                    Trans.IndexValue(3) = 0
                    Trans.IndexValue(4) = yAxis.x
                    Trans.IndexValue(5) = yAxis.y
                    Trans.IndexValue(6) = yAxis.z
                    Trans.IndexValue(7) = 0
                    Trans.IndexValue(8) = zax
                    Trans.IndexValue(9) = zay
                    Trans.IndexValue(10) = zaz
                    Trans.IndexValue(11) = 0
                    Trans.IndexValue(12) = oX
                    Trans.IndexValue(13) = oY
                    Trans.IndexValue(14) = oZ
                    Trans.IndexValue(15) = 1
                    
                    Dim tempPoint As DPosition
                    Set tempPoint = New DPosition
                    For a = LBound(tempHoles) To UBound(tempHoles)
                    
                        tempPoint.Set tempHoles(a, 1), tempHoles(a, 2), 0
                        ''Firsttrans is the inverted transformation matrix from the first port.  (destination of second port)
                        ''Trans is the transformation matrix from the second port. (convert points from this to first)
                        ''tempPoint is a point in the coordinate system of second port.  Transform to be in first port's cs
                        Set tempPoint = FirstTrans.TransformPosition(Trans.TransformPosition(tempPoint))
                        Set holes(CurrentHole) = New DPosition
                        holes(CurrentHole).Set tempPoint.x, tempPoint.y, tempPoint.z
                        CurrentHole = CurrentHole + 1
                    
                    Next a
                End If
            End If
        Else 'point case. This method should be called only in case of multiple points.
             ' Commenting it for now.
'''            Dim oPoint As IJPoint
'''            Dim tempPoint1 As New DPosition
'''            Set oPoint = oPointOrPort
'''            oPoint.GetPoint oX, oY, oZ
'''            tempPoint1.Set oX, oY, oZ
'''            ReDim Preserve holes(UBound(holes) + 1) As DPosition
'''            Set holes(CurrentHole) = New DPosition
'''            If CurrentHole = 1 Then
'''                FirstTrans.IndexValue(5) = -1
'''                FirstTrans.IndexValue(12) = oX
'''                FirstTrans.IndexValue(13) = oY
'''                FirstTrans.IndexValue(14) = oZ
'''                FirstTrans.IndexValue(15) = 1
'''            End If
'''            Set tempPoint1 = FirstTrans.TransformPosition(Trans.TransformPosition(tempPoint1))
'''            tempPoint1.Get oX, oY, oZ
'''
'''            holes(CurrentHole).Set oX, oY, oZ
'''            CurrentHole = CurrentHole + 1
        End If
    Next
    
    NumberOfHoles = CurrentHole - 1
    FirstTrans.Invert  'Returns the matrix for transforming the points into global space.
    
    Set xAxis = Nothing
    Set yAxis = Nothing
    Set zaxis = Nothing
    Set oPort = Nothing
    Set Trans = Nothing
    
    Exit Sub
        
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub SetOrigin(pPropertyDescriptions As Object, ByRef eqPos As Automath.DPosition)
Const METHOD = "SetOrigin"
On Error GoTo ErrorHandler

    Dim oObj As IJLocalCoordinateSystem
    Set oObj = pPropertyDescriptions.CAO
    
    Dim oPosition As DPosition
    Set oPosition = oObj.Position
    
    Dim dx, dy, dz As Double
    dx = oPosition.x
    dy = oPosition.y
    dz = oPosition.z
    eqPos.Set dx, dy, dz
    
    Set oObj = Nothing
    Set oPosition = Nothing
    
Exit Sub
    
ErrorHandler:
    HandleError MODULE, METHOD
End Sub
Public Sub GetOriginAndDirections(pPropertyDescriptions As Object, ByRef eqPos As Automath.DPosition, _
                                  ByRef xVec As DVector, ByRef yVec As DVector, ByRef zVec As DVector)
Const METHOD = "GetOriginAndDirections"
On Error GoTo ErrorHandler
    
    Dim oFndnMatrix As IJDMatrixAccess
    Set oFndnMatrix = pPropertyDescriptions.CAO
    Dim newTmx As Automath.DT4x4
    Set newTmx = oFndnMatrix.Matrix
    Dim x As Double, y As Double, z As Double
    
    'Set the origin
    x = newTmx.IndexValue(12)
    y = newTmx.IndexValue(13)
    z = newTmx.IndexValue(14)
    eqPos.Set x, y, z
    
    'Set the X axis vector
    x = newTmx.IndexValue(0)
    y = newTmx.IndexValue(1)
    z = newTmx.IndexValue(2)
    xVec.Set x, y, z
    
    'Set the Y axis vector
    x = newTmx.IndexValue(4)
    y = newTmx.IndexValue(5)
    z = newTmx.IndexValue(6)
    yVec.Set x, y, z
    
    
    'Set the Z axis vector
    x = newTmx.IndexValue(8)
    y = newTmx.IndexValue(9)
    z = newTmx.IndexValue(10)
    zVec.Set x, y, z

Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

'Finds the 'Range' that these holes encompass.
Public Sub CalculateRectangleOrientation(pts() As DPosition, _
                                         ByRef xVec As DVector, _
                                         ByRef yVec As DVector)
Const METHOD = "CalculateRectangleOrientation"
     
    Dim index As Long, ii As Long
    Dim x As Double, y As Double, z As Double
    Dim a As Integer, Min As Integer, First As Integer
    Dim MaxX As Double, MaxY As Double, MinX As Double, MinY As Double
    Dim m_BoltOffsetLength As Double, m_BoltOffsetWidth As Double
    Set oLocalizer = New IMSLocalizer.Localizer
    oLocalizer.Initialize App.Path & "\" & "SPSEqpFndMacros"
    
    m_BoltOffsetLength = 0#  ' 0.25
    m_BoltOffsetWidth = 0#  ' 0.127
    
    z = pts(1).z
    
    First = 1
    MaxX = pts(First).x
    MinX = pts(First).x
    MaxY = pts(First).y
    MinY = pts(First).y
    
    If UBound(pts) > 1 Then
        For a = First + 1 To UBound(pts)
            If pts(a).x > MaxX Then
                MaxX = pts(a).x
            End If
            If pts(a).x < MinX Then
                MinX = pts(a).x
            End If
            If pts(a).y > MaxY Then
                MaxY = pts(a).y
            End If
            If pts(a).y < MinY Then
                MinY = pts(a).y
            End If
        Next
        
        If Abs(MaxY - MinY) > 0.000001 Then
        Else
            On Error GoTo 0
            Err.Raise E_FAIL
        End If
          
        If Abs(MaxX - MinX) > 0.000001 Then
        Else
            On Error GoTo 0
            Err.Raise E_FAIL
        End If
    End If
    
    'Sets the length to be the direction with the longest vector.
    If MaxY - MinY > MaxX - MinX Then
        xVec.Set (MaxX - MinX), 0, 0
        xVec.Length = xVec.Length + 2 * m_BoltOffsetWidth
        yVec.Set 0, (MinY - MaxY), 0
        yVec.Length = yVec.Length + 2 * m_BoltOffsetLength
    Else
        xVec.Set 0, (MaxY - MinY), 0
        xVec.Length = xVec.Length + 2 * m_BoltOffsetWidth
        yVec.Set MaxX - MinX, 0, 0     'Length
        yVec.Length = yVec.Length + 2 * m_BoltOffsetLength
    End If
    Set oLocalizer = Nothing
    Exit Sub
    
End Sub

Public Sub SetRectangleOrientationByClearance(pts() As DPosition, _
                                              dXClear As Double, _
                                              dYClear As Double, _
                                              ByRef xVec As DVector, _
                                              ByRef yVec As DVector)
Const METHOD = "SetRectangleOrientationByClearance"
    Dim index As Long, ii As Long
    Dim x As Double, y As Double, z As Double
    Dim a As Integer, Min As Integer, First As Integer
    Dim MaxX As Double, MaxY As Double, MinX As Double, MinY As Double
    
    z = pts(1).z
    
    First = 1
    MaxX = pts(First).x
    MinX = pts(First).x
    MaxY = pts(First).y
    MinY = pts(First).y
    If UBound(pts) > 1 Then
        For a = First + 1 To UBound(pts)
            If pts(a).x > MaxX Then
                MaxX = pts(a).x
            End If
            If pts(a).x < MinX Then
                MinX = pts(a).x
            End If
            If pts(a).y > MaxY Then
                MaxY = pts(a).y
            End If
            If pts(a).y < MinY Then
                MinY = pts(a).y
            End If
        Next
        
        If Abs(MaxY - MinY) > 0.000001 Then
        Else
    
            Err.Raise E_FAIL
        End If
          
        If Abs(MaxX - MinX) > 0.000001 Then
        Else
            Err.Raise E_FAIL
        End If
    End If
    
    'Sets the length to be the direction with the longest vector.
    If MaxY - MinY > MaxX - MinX Then
        xVec.Set (MaxX - MinX), 0, 0
        xVec.Length = xVec.Length + dXClear
        yVec.Set 0, (MinY - MaxY), 0
        yVec.Length = yVec.Length + dYClear
    Else
        xVec.Set 0, (MaxY - MinY), 0
        xVec.Length = xVec.Length + dXClear
        yVec.Set MaxX - MinX, 0, 0     'Length
        yVec.Length = yVec.Length + dYClear
    End If
        
    Exit Sub
    
End Sub

'Manish TR#65124- This should return center of range generated by selected equipments
' that is (maxX-minX)/2 & (*maxY-minY)/2 and not  centroid as it was doing earlier.
Public Sub GetCentroidOfPositions(pts() As DPosition, _
                             NumberOfHoles As Integer, _
                             centroidX As Double, _
                             centroidY As Double, _
                             centroidZ As Double)
Const METHOD = "GetCentroidOfPositions "
On Error GoTo ErrorHandler
    
    Dim xmin As Double, xmax As Double
    Dim ymin As Double, ymax As Double
    Dim zmax As Double, zmin As Double
    
    Dim i As Long

    xmin = pts(1).x
    ymin = pts(1).y
    zmin = pts(1).z
    xmax = pts(1).x
    ymax = pts(1).y
    zmax = pts(1).z
    

    For i = 1 To NumberOfHoles
        If xmin > pts(i).x Then xmin = pts(i).x
        If ymin > pts(i).y Then ymin = pts(i).y
        If zmin > pts(i).z Then zmin = pts(i).z
        If xmax < pts(i).x Then xmax = pts(i).x
        If ymax < pts(i).y Then ymax = pts(i).y
        If zmax < pts(i).z Then zmax = pts(i).z
    Next i
    
    centroidX = (xmax + xmin) / 2
    centroidY = (ymax + ymin) / 2
    centroidZ = (zmax + zmin) / 2

Exit Sub
    
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CalculateCircularOrientation(pts() As DPosition, _
                                         NumberOfHoles As Integer, _
                                         ByRef origin As DPosition, _
                                         ByRef xVec As DVector, _
                                         ByRef yVec As DVector)
Const METHOD = "CalculateCircularOrientation"
On Error GoTo ErrorHandler
    
    Dim CenterX As Double, CenterY As Double, CenterZ As Double
    Dim CircleCenter As Automath.DPosition
    Set CircleCenter = New DPosition
    'this given algorithm takes care of 2 or more holes,
    'no steps are taken if there is only one hole
    If NumberOfHoles > 2 Then
        
        GetCentroidOfPositions pts(), NumberOfHoles, CenterX, CenterY, CenterZ   ' TR#65124
        CircleCenter.Set CenterX, CenterY, CenterZ
      
        Dim i As Integer
       
        For i = 1 To (NumberOfHoles)
            If pts(i).DistPt(CircleCenter) > xVec.Length Then  'TR#65124
                xVec.x = pts(i).x - CircleCenter.x
                xVec.y = pts(i).y - CircleCenter.y
            End If ' TR#65124
        Next i
        xVec.z = 0 'pts(1).z
        
    ElseIf NumberOfHoles = 2 Then
        'assuming the holes to be 180 degrees apart
        Dim x1 As Double, x2 As Double, x3 As Double
        Dim y1 As Double, y2 As Double, y3 As Double
        Dim z1 As Double, z2 As Double, z3 As Double
        
        x1 = pts(1).x
        y1 = pts(1).y
        z1 = pts(1).z
        
        x2 = pts(2).x
        y2 = pts(2).y
        z2 = pts(2).z
        
        CircleCenter.x = (x1 + x2) / 2#
        CircleCenter.y = (y1 + y2) / 2#
        CircleCenter.z = pts(1).z '(z1 + z2) / 2#
        
        xVec.x = pts(1).x - CircleCenter.x
        xVec.y = pts(1).y - CircleCenter.y
        xVec.z = 0 'pts(1).z

    End If
    origin.Set CircleCenter.x, CircleCenter.y, CircleCenter.z
    
    Dim zVec As Automath.DVector
    Set zVec = New DVector
    zVec.Set 0#, 0#, -1#
    
    Set yVec = zVec.Cross(xVec)
    
    Set CircleCenter = Nothing
    
    Exit Sub
    
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CalculateDistanceFromPtToSurface(dPt As DPosition, _
                                            oLineAxis As DVector, _
                                            dNPt As DPosition, _
                                            dRPt As DPosition, _
                                            ByRef dHt As Double)
Const METHOD = "CalculateDistanceFromPtToSurface"
On Error GoTo ErrorHandler
    
    Dim oLine As Line3d
    Dim oGeomFactory As New IngrGeom3D.GeometryFactory
    Dim DummyFace As New Plane3d
    Dim oNewBottomSurf As IJSurface
    Dim code As Geom3dIntersectConstants
    Dim oTempEles As IJElements
    Set oTempEles = New JObjectCollection 'Elements
    
    Set oLine = oGeomFactory.Lines3d.CreateByPtVectLength(Nothing, dPt.x, dPt.y, dPt.z, oLineAxis.x, oLineAxis.y, oLineAxis.z, INFINITY)
   
    oLine.Infinite = True
    Set DummyFace = oGeomFactory.Planes3d.CreateByPointNormal(Nothing, _
                                                              dRPt.x, dRPt.y, dRPt.z, _
                                                              dNPt.x, dNPt.y, dNPt.z)
        
    Set oNewBottomSurf = DummyFace

    oNewBottomSurf.Intersect oLine, oTempEles, code
  
    If oTempEles.count <> 0 Then
    
        Dim pt1 As Double, pt2 As Double, pt3 As Double
        Dim oPoint As IJPoint
        Dim dist As Double
        
        Set oPoint = New Point3d
        Set oPoint = oTempEles.Item(1)
        oPoint.GetPoint pt1, pt2, pt3
        
        dHt = Sqr((pt1 - dPt.x) * (pt1 - dPt.x) + (pt2 - dPt.y) * (pt2 - dPt.y) + (pt3 - dPt.z) * (pt3 - dPt.z))
        If dHt > 0.00001 Then Else dHt = 0#     ' tr 55357
    End If

    Set oLine = Nothing
    Set DummyFace = Nothing
    Set oGeomFactory = Nothing
    Set oNewBottomSurf = Nothing
    
    If Not oTempEles Is Nothing Then
        oTempEles.Clear
        Set oTempEles = Nothing
    End If
  
    Exit Sub
    
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub SetInputs_Supported_Supporting(ByVal FndObject As Object, _
                                          ByVal FndDefnObject As Object, _
                                          ByVal Supported As IJElements, _
                                          ByVal Supporting As IJElements)
Const METHOD = "SetInputs_Supported_Supporting"

    Dim strUserType As String
    Dim strSCName As String
    Dim oSmartItem As IJSmartItem
    Dim oSmartClass As IJSmartClass
    Dim oUserType As IJDUserType
    Dim oSmartOcc As IJSmartOccurrence
    Dim oEquipFactory As SPSEquipFoundationFactory
    Dim ijdObj As iJDObject
    Dim NewSO As Boolean
    Dim i As Integer
    NewSO = False
    
    ' Create the reference collection
    Dim oSymbolEntitiesFactory As New IMSSymbolEntities.DSymbolEntitiesFactory
    Dim oReferencesCollection As IMSSymbolEntities.IJDReferencesCollection
    Dim oReferencesCollection1 As IMSSymbolEntities.IJDReferencesCollection
    Dim oReferencesCollection2 As IMSSymbolEntities.IJDReferencesCollection
    
   
    Set oSmartOcc = FndObject
    On Error Resume Next
    Set oReferencesCollection = GetRefCollection(oSmartOcc)
    If Not oReferencesCollection Is Nothing Then
        Set oReferencesCollection1 = oReferencesCollection.IJDEditJDArgument.GetEntityByIndex(1)
        Set oReferencesCollection2 = oReferencesCollection.IJDEditJDArgument.GetEntityByIndex(2)
        If Not oReferencesCollection1 Is Nothing Then
            oReferencesCollection1.IJDEditJDArgument.RemoveAll
        Else
            Set oReferencesCollection1 = oSymbolEntitiesFactory.CreateEntity(referencesCollection, _
                                                                             GetResourceMgr())
        End If
        
        If Not oReferencesCollection2 Is Nothing Then
            oReferencesCollection2.IJDEditJDArgument.RemoveAll
        Else
            If Supporting.count >= 1 Then   ' tr 76833
                Set oReferencesCollection2 = oSymbolEntitiesFactory.CreateEntity(referencesCollection, _
                                                                                 GetResourceMgr())
            End If
            
        End If
        oReferencesCollection.IJDEditJDArgument.RemoveAll
        NewSO = False
    Else
        Set oReferencesCollection = oSymbolEntitiesFactory.CreateEntity(referencesCollection, _
                                                                        GetResourceMgr())
        Set oReferencesCollection1 = oSymbolEntitiesFactory.CreateEntity(referencesCollection, _
                                                                         GetResourceMgr())
        If Supporting.count >= 1 Then   ' tr 76833
            Set oReferencesCollection2 = oSymbolEntitiesFactory.CreateEntity(referencesCollection, _
                                                                         GetResourceMgr())
        End If
        NewSO = True
    End If
    
      
    Dim OldItem As IJSmartItem
    Dim strOldItemName As String
    On Error Resume Next
    Set OldItem = oSmartOcc.ItemObject
    Err.Clear

    
    If Not OldItem Is Nothing Then
        strOldItemName = OldItem.name
    End If
    
    Set oSmartItem = FndDefnObject
    If strOldItemName <> oSmartItem.name Then
        
        Set oSmartClass = oSmartItem.Parent
        strUserType = oSmartClass.SOUserType
        Set oUserType = oSmartOcc
        oUserType.UserType = strUserType
        strSCName = oSmartClass.SCName
        oSmartOcc.RootSelectorClass = strSCName
        oSmartOcc.ROOTSELECTION = oSmartItem.name
    End If
    
    For i = 1 To Supported.count
        If TypeOf Supported.Item(i) Is IJEqpFoundationPort Then
            oReferencesCollection1.IJDEditJDArgument.SetEntity i, Supported.Item(i), IJEqpFoundationPort, "SPSFndnPortToRC_DEST"
        Else
            oReferencesCollection1.IJDEditJDArgument.SetEntity i, Supported.Item(i), IJFullObject, "SPSRCtoRC_DEST"
        End If
        
    Next i

    oReferencesCollection.IJDEditJDArgument.SetEntity 1, oReferencesCollection1, IJFullObject, "SPSRCToRC_1_DEST"
    
    If Supporting.count >= 1 Then
        Dim oTempObj As Object, oSupportingSurface As Object
        Set oTempObj = Supporting.Item(1)
        Set oSupportingSurface = GetStablePortIfApplicable(oTempObj)
        Set oTempObj = Nothing
        oReferencesCollection2.IJDEditJDArgument.SetEntity 1, oSupportingSurface, IJPlane, "SPSSuppPlaneToRC_DEST"
        
        oReferencesCollection.IJDEditJDArgument.SetEntity 2, oReferencesCollection2, IJFullObject, "SPSRCToRC_2_DEST"
    Else
        If Not oReferencesCollection2 Is Nothing Then
            Set ijdObj = oReferencesCollection2
            ijdObj.Remove
        End If
    End If
    
    If NewSO Then
        ConnectSmartOccurrence oSmartOcc, oReferencesCollection
    End If
        
    Set oSymbolEntitiesFactory = Nothing
    Set oReferencesCollection = Nothing
    Set oReferencesCollection1 = Nothing
    Set oReferencesCollection2 = Nothing
    Set oSmartOcc = Nothing
    Set oUserType = Nothing
    Set oSmartItem = Nothing
    Set oSmartClass = Nothing

Exit Sub
    
End Sub

Public Sub GetInputs_Supported_Supporting(ByVal FndObject As Object, _
                                          ByVal Supported As IJElements, _
                                          ByVal Supporting As IJElements)
Const METHOD = "GetInputs_Supported_Supporting"
On Error GoTo ErrorHandler

    Dim oSmartOcc As IJSmartOccurrence
    Set oSmartOcc = FndObject
    
    Dim oRefColl As IJDReferencesCollection
    Set oRefColl = GetRefCollection(oSmartOcc)
    
    Dim oRefColl1 As IJDReferencesCollection
    Dim oRefColl2 As IJDReferencesCollection
    
    Set oRefColl1 = oRefColl.IJDEditJDArgument.GetEntityByIndex(1)
    On Error Resume Next
    Set oRefColl2 = oRefColl.IJDEditJDArgument.GetEntityByIndex(2)
    
'    Dim i As Integer
'    Dim cnt As Integer
    
'    cnt = oRefColl1.IJDEditJDArgument.GetCount
'    For i = 1 To cnt
'        Supported.Add oRefColl1.IJDEditJDArgument.GetEntityByIndex(i)
'    Next i
    
    Dim pEnumJDArgument As IEnumJDArgument
    Dim arg1 As IJDArgument
    Dim found As Long
    Dim iloop As Long
     
    Set pEnumJDArgument = oRefColl1
    If Not pEnumJDArgument Is Nothing Then
        pEnumJDArgument.Reset
        iloop = 1
        Do
           pEnumJDArgument.Next 1, arg1, found
           If found <> 0 Then
                Supported.Add arg1.Entity
                Set arg1 = Nothing
'                iloop = iloop + 1 ' will cause memory overwrite shouldn't be here
           Else: Exit Do
           End If
        Loop
    End If
    
    If oRefColl2 Is Nothing Then
    Else
        Dim cnt2 As Integer
        cnt2 = oRefColl2.IJDEditJDArgument.GetCount
        If cnt2 >= 1 Then Supporting.Add oRefColl2.IJDEditJDArgument.GetEntityByIndex(1)
    End If
    
    Set oRefColl = Nothing
    Set oRefColl1 = Nothing
    Set oRefColl2 = Nothing
    Set pEnumJDArgument = Nothing
    Set oSmartOcc = Nothing
    
    Exit Sub
    
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Function IsSOOverridden(attrCol As CollectionProxy) As Boolean
Const MT = "IsSOOverridden"
On Error GoTo ErrorHandler
Dim i As Integer
Dim oAttr As IJDAttribute
  IsSOOverridden = False
  For i = 1 To attrCol.count
    Set oAttr = attrCol.Item(i)
    If Not IsEmpty(oAttr.Value) Then
        IsSOOverridden = True
        Set oAttr = Nothing
        Exit For
    End If
    Set oAttr = Nothing
  Next i
 
 Exit Function
ErrorHandler: HandleError MODULE, MT
End Function

Public Sub CopyValuesToSOFromItem(soCol As CollectionProxy, itmCol As CollectionProxy)
Const MT = "CopyValuesToSOFromItem"
On Error GoTo ErrorHandler
    Dim i As Integer

    If soCol.count <> itmCol.count Then
        Exit Sub
    End If
    For i = 1 To soCol.count
       
        If soCol.Item(i).AttributeInfo.name = itmCol.Item(i).AttributeInfo.name Then
            soCol.Item(i).Value = itmCol.Item(i).Value
        End If
    Next i
     
    Exit Sub

ErrorHandler: HandleError MODULE, MT

End Sub

Public Sub AddMaterialRelationShip(FoundObj As Object, MatlObj As Object, RelName As String)
Const METHOD = "AddMaterialRelationShip"
On Error GoTo ErrorHandler
    
    Dim oIJDAssocRelation As IJDAssocRelation
    Dim oTargetObjCol As IJDTargetObjectCol
    Dim oRelship As IMSRelation.DRelationshipHelper
    Dim iProxy As IJDProxy
    Dim oModelPOM As IJDPOM
    Dim iJDObject As iJDObject
    
    Set iJDObject = FoundObj
    Set oModelPOM = iJDObject.ResourceManager
    Set iProxy = oModelPOM.GetProxy(MatlObj, True)
    
    Set oIJDAssocRelation = FoundObj
    
    Set oTargetObjCol = oIJDAssocRelation.CollectionRelations("IJStructMaterial", "StructEntityMaterial_ORIG")
     
        Dim pUnk As Object
        On Error Resume Next
        Set pUnk = oTargetObjCol.Item(RelName)
        If pUnk Is Nothing Then
            Call oTargetObjCol.Add(iProxy, RelName, oRelship)
        ElseIf (Not pUnk Is Nothing) And pUnk Is MatlObj Then
            Exit Sub
        ElseIf (Not pUnk Is Nothing) And Not pUnk Is MatlObj Then
            Call oTargetObjCol.Remove(RelName)
            Call oTargetObjCol.Add(iProxy, RelName, oRelship)
        End If
        
Exit Sub
ErrorHandler: HandleError MODULE, METHOD
End Sub

Public Function GetCrossSection(ByVal SectionStandard As String, _
                                  ByVal SectionName As String) As Object
Const METHOD = "GetCrossSection"
On Error GoTo ErrorHandler

    Dim oCrossSec As Object
    Dim oStructServices As New RefDataMiddleServices.StructCrossSectionServices
    oStructServices.GetStructureCrossSectionDefinition GetCatalogResourceManager, SectionStandard, "", SectionName, oCrossSec
    
    Set GetCrossSection = oCrossSec
    
    Set oCrossSec = Nothing
    Set oStructServices = Nothing
    
    Exit Function
ErrorHandler: HandleError MODULE, METHOD
End Function

Public Sub GenerateNameForMember(Obj As Object)
Const METHOD = "GenerateNameForMember"
On Error GoTo ErrorHandler

    Dim NameRule As String
    Dim found As Boolean
    found = False
    On Error Resume Next
      
    Dim NamingRules As IJElements
    Dim oNameRuleHolder As GSCADGenericNamingRulesFacelets.IJDNameRuleHolder
    Dim oActiveNRHolder As GSCADGenericNamingRulesFacelets.IJDNameRuleHolder
    Dim oNameRuleHlpr As GSCADNameRuleSemantics.IJDNamingRulesHelper
    Set oNameRuleHlpr = New GSCADNameRuleHlpr.NamingRulesHelper
    Call oNameRuleHlpr.GetEntityNamingRulesGivenProgID(STRUCTMEMPARTPROGID, NamingRules)
    Dim ncount As Integer
    Dim oNameRuleAE As GSCADGenNameRuleAE.IJNameRuleAE
      
    For ncount = 1 To NamingRules.count
        Set oNameRuleHolder = NamingRules.Item(1)
    Next ncount

    Call oNameRuleHlpr.AddNamingRelations(Obj, oNameRuleHolder, oNameRuleAE)
    Set oNameRuleHolder = Nothing
    
    Set oActiveNRHolder = Nothing
    Set oNameRuleHolder = Nothing
    Set oNameRuleAE = Nothing
 
    Exit Sub

ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub GetCrossSecData(SectionName As String, _
                            SectionStandard As String, _
                            ByRef sectionWidth As Double, _
                            ByRef sectionDepth As Double)
Const METHOD = "GetCrossSecData"
On Error GoTo ErrorHandler
                     
    Dim Attrs As IJDAttributes
    
    Set Attrs = GetCrossSection(SectionStandard, SectionName)
    sectionDepth = Attrs.CollectionOfAttributes(CROSSSEC_IFACE).Item(CROSSSEC_DEPTH_ATTRNAME).Value
    sectionWidth = Attrs.CollectionOfAttributes(CROSSSEC_IFACE).Item(CROSSSEC_WIDTH_ATTRNAME).Value

    Set Attrs = Nothing
    
    Exit Sub
    
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

' get symbol outputs
Public Function GetSymbolOutputs(pObject As Object, representationName As String) As IJElements
Const METHOD = "GetSymbolOutputs"
On Error GoTo ErrHandler

    Dim iIJDSymbol As IJDSymbol
    Set iIJDSymbol = pObject
    
    Dim iIJDSymbolDefinition As IJDSymbolDefinition
    Dim iIJDRepresentations As IJDRepresentations
    Dim iIJDRepresentation As IJDRepresentation
    
    Set iIJDSymbolDefinition = iIJDSymbol.IJDSymbolDefinition(0)
    Set iIJDRepresentations = iIJDSymbolDefinition.IJDRepresentations
    Set iIJDRepresentation = iIJDRepresentations.GetRepresentationByName(representationName)
    
    Dim iIJDOutputs As IJDOutputs
    'Get the IJDOutputs from the representation
    Set iIJDOutputs = iIJDRepresentation
    
    Dim count As Long, ii As Long
    count = iIJDOutputs.OutputCount
    
    Dim iIJElements As IJElements
    Set iIJElements = New JObjectCollection
    
    Dim iIJDOutput As IJDOutput
    Dim strOutputName As String
    Dim iIDispatch As Object
    On Error Resume Next
    For ii = 1 To count
        Set iIJDOutput = iIJDOutputs.GetOutputAtIndex(ii)
        strOutputName = iIJDOutput.name
        Set iIDispatch = iIJDSymbol.BindToOutput(representationName, strOutputName)
        iIJElements.Add iIDispatch
    Next ii
    Set GetSymbolOutputs = iIJElements
    
    Set iIJDSymbol = Nothing
    Set iIJDSymbolDefinition = Nothing
    Set iIJDRepresentations = Nothing
    Set iIJDRepresentation = Nothing
    Set iIJDOutputs = Nothing
    Set iIJElements = Nothing
    Set iIJDOutput = Nothing
    Exit Function
    
ErrHandler:
    HandleError MODULE, METHOD
End Function

'compute weight and cg from symbol outputs
Public Sub CalculateVolumeCGSurfaceArea(iIJElements As IJElements, _
                                      ByRef outVolume As Double, _
                                      ByRef outX As Double, _
                                      ByRef outY As Double, _
                                      ByRef outZ As Double, _
                                      ByRef outPaintArea As Double)
Const METHOD = "CalculateVolumeCGSurfaceArea"
On Error GoTo ErrHandler

    Dim eleCount As Long
    eleCount = iIJElements.count
    Dim ii As Long
    Dim iProj As IngrGeom3D.Projection3d
    Dim projVecx As Double, projVecy As Double, projVecz As Double
    Dim iCurve As IJCurve
    Dim cvArea As Double, cvCGx As Double, cvCGy As Double, cvCGz As Double
    Dim projCGx As Double, projCGy As Double, projCGz As Double
    Dim projVolume As Double, accumVolume As Double
    Dim projVecLength As Double
    Dim accumCGx As Double, accumCGy As Double, accumCGz As Double
    Dim paintArea As Double
    Dim iPlane As IJPlane
    Dim oBoundary As ComplexString3d
    accumVolume = 0#
    accumCGx = 0#
    accumCGy = 0#
    accumCGz = 0#
    paintArea = 0#

    For ii = 1 To eleCount
    
        If TypeOf iIJElements.Item(ii) Is Projection3d Then
            
            Set iProj = iIJElements.Item(ii)
            'get the projection vector and the base curve
            iProj.GetProjection projVecx, projVecy, projVecz
            Set iCurve = iProj.Curve
            cvArea = iCurve.Area
            iCurve.centroid cvCGx, cvCGy, cvCGz
            projCGx = cvCGx + 0.5 * projVecx
            projCGy = cvCGy + 0.5 * projVecy
            projCGz = cvCGz + 0.5 * projVecz
            projVecLength = Sqr((projVecx * projVecx) + (projVecy * projVecy) + (projVecz * projVecz))
            projVolume = cvArea * projVecLength
            
            'accumCG gets moved by the ratio of volumes
            accumVolume = accumVolume + projVolume
            If accumVolume > 0 Then
                accumCGx = accumCGx + (projCGx - accumCGx) * projVolume / accumVolume
                accumCGy = accumCGy + (projCGy - accumCGy) * projVolume / accumVolume
                accumCGz = accumCGz + (projCGz - accumCGz) * projVolume / accumVolume
            End If
            
            ' paintArea is surface area of the surfaceOfProjection
            paintArea = paintArea + 2 * cvArea + iCurve.Length * projVecLength
            
        ElseIf TypeOf iIJElements.Item(ii) Is IJPlane Then
            
            Set iPlane = iIJElements.Item(ii)
            
            iPlane.GetBoundary 1, oBoundary
            Set iPlane = Nothing
            
            Set iCurve = oBoundary
            iCurve.centroid cvCGx, cvCGy, cvCGz
            accumCGx = accumCGx + cvCGx
            accumCGy = accumCGy + cvCGy
            accumCGz = accumCGz + cvCGz
            paintArea = paintArea + iCurve.Area
            
            If ii = eleCount Then
                accumCGx = accumCGx / eleCount
                accumCGy = accumCGy / eleCount
                accumCGz = accumCGz / eleCount
            End If
            Set oBoundary = Nothing
            Set iPlane = Nothing
        End If
        Set iCurve = Nothing
        
    Next ii 'for each element
    
    outVolume = accumVolume
    outX = accumCGx
    outY = accumCGy
    outZ = accumCGz
    outPaintArea = paintArea

    Set iProj = Nothing
    Set iCurve = Nothing
    Exit Sub
    
ErrHandler:
    HandleError MODULE, METHOD
End Sub


Public Sub InitNewOutput(pOC As IJDOutputCollection, name As String)
Const METHOD = "InitNewOutput"
On Error GoTo ErrorHandler
    
    
    Dim oRep As IMSSymbolEntities.IJDRepresentation
    Dim oOutputs As IMSSymbolEntities.IJDOutputs
    Dim oOutput As IMSSymbolEntities.IJDOutput
    
    Set oOutput = New DOutput
    Set oRep = pOC.definition.IJDRepresentations.GetRepresentationByName("Physical")
    Set oOutputs = oRep

    oOutput.name = name
    oOutput.Description = name
    oOutputs.SetOutput oOutput
    oOutput.Reset
    Exit Sub
ErrorHandler:
      HandleError MODULE, METHOD

End Sub


Public Sub InitOctCurvePoints(pts() As Double, OctOverAllDim As Double, ZDirec)
Const METHOD = "InitOctCurvePoints"
On Error GoTo ErrorHandler

Dim OctFaceLength As Double

  OctFaceLength = OctOverAllDim / (1 + 2 * Sin(PI / 4))
  
 'Build points in local XY plane at the centroid of the rectangle
    pts(0) = -(OctFaceLength / 2#)
    pts(1) = (OctOverAllDim / 2#)
    pts(2) = ZDirec
    
    pts(3) = (OctFaceLength / 2#)
    pts(4) = (OctOverAllDim / 2#)
    pts(5) = ZDirec
    
    pts(6) = (OctOverAllDim / 2#)
    pts(7) = (OctFaceLength / 2#)
    pts(8) = ZDirec
    
    pts(9) = (OctOverAllDim / 2#)
    pts(10) = -(OctFaceLength / 2#)
    pts(11) = ZDirec
    
    pts(12) = (OctFaceLength / 2#)
    pts(13) = -(OctOverAllDim / 2#)
    pts(14) = ZDirec
    
    pts(15) = -(OctFaceLength / 2#)
    pts(16) = -(OctOverAllDim / 2#)
    pts(17) = ZDirec
    
    pts(18) = -(OctOverAllDim / 2#)
    pts(19) = -(OctFaceLength / 2#)
    pts(20) = ZDirec
    
    pts(21) = -(OctOverAllDim / 2#)
    pts(22) = (OctFaceLength / 2#)
    pts(23) = ZDirec
    
    pts(24) = pts(0)
    pts(25) = pts(1)
    pts(26) = pts(2)
    
Exit Sub
ErrorHandler:   HandleError MODULE, METHOD
End Sub

Public Function CreateSolidAsPlanesUsingSweep(pConnection As IUnknown, _
                                              pCSLine As IJLineString, _
                                              traceVec As IJDVector, _
                                              SolidHeight As Double) As IJDObjectCollection
Const METHOD = "CreateSolidAsPlanesUsingSweep"
 
    On Error GoTo ErrorHandler
    
    Dim oGeomFactory As New IngrGeom3D.GeometryFactory
    Dim oCSCmpxCurve As ComplexString3d, oTraceCmpxCurve As ComplexString3d
    Dim oCScurve As Line3d, oTraceCurve As Line3d
    Dim pX As Double, pY As Double, pZ As Double
    Dim ii As Integer, j As Integer
    Dim numPts As Long
    Dim dCSpts() As Double
    
    
    ' create a complex string and add cross section to it
    Dim iElements As IJElements
    Set iElements = New JObjectCollection
    
    pCSLine.GetPoints numPts, dCSpts
    j = 0
    
    For ii = 1 To numPts - 1
        Set oCScurve = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                                            dCSpts(j * 3), dCSpts(1 + j * 3), dCSpts(2 + j * 3), _
                                                            dCSpts(3 + j * 3), dCSpts(4 + j * 3), dCSpts(5 + j * 3))
        If Not oCScurve Is Nothing Then iElements.Add oCScurve
        Set oCScurve = Nothing
        j = j + 1
    Next ii
    
    Set oCSCmpxCurve = oGeomFactory.ComplexStrings3d.CreateByCurves(Nothing, iElements)
    
    ' create a trace curve
    Set oTraceCurve = oGeomFactory.Lines3d.CreateByPtVectLength(Nothing, _
                                                            dCSpts(0), dCSpts(1), dCSpts(2), _
                                                            traceVec.x, traceVec.y, traceVec.z, _
                                                            SolidHeight)
    ' create solid using singlesweep
    ' use break both trace and cross section option and with caps
    Dim stNorm() As Double, endNorm() As Double
    Dim numCaps As Long
    Set CreateSolidAsPlanesUsingSweep = oGeomFactory.GeometryServices.CreateBySingleSweepWCaps(pConnection, _
                                                        oTraceCurve, oCSCmpxCurve, _
                                                        CircularCorner, BreakPathAndCrossSection, _
                                                        StartAtTraceBeg, stNorm, endNorm, _
                                                        True, numCaps)
    Set oGeomFactory = Nothing
    Set oTraceCmpxCurve = Nothing
    Set oCSCmpxCurve = Nothing
    Set oCScurve = Nothing
    Set iElements = Nothing
    Set oTraceCurve = Nothing
    
    Exit Function

ErrorHandler:
    HandleError MODULE, METHOD
End Function

'*************************************************************************
'Function
'MigrateEqpFnd
'
'Abstract
'Migrates thr EF to the correct surface if it is split.
'
'Arguments
'IJDAggregatorDescription interface of the member
'
'Return
'
'Exceptions
'
'*************************************************************************

Public Sub MigrateEqpFnd(oAggregatorDesc As IJDAggregatorDescription, oMigrateHelper As IJMigrateHelper)

  Const MT = "MigrateEqpFnd"
  On Error GoTo ErrorHandler
    
    Dim oSmartOcc As IJSmartOccurrence
    Dim oRefCollAsm As IJDReferencesCollection
    Dim oRefColl1 As IJDReferencesCollection
    Dim oRefColl2 As IJDReferencesCollection
    Dim oObjectsReplacing() As Object
    Dim bIsInputMigrated As Boolean
    Dim iRCCount As Integer
    Dim oEqpFndPort As IJEqpFoundationPort
    Dim dOrigX As Double
    Dim dOrigY As Double
    Dim dOrigZ As Double
    Dim dZX  As Double
    Dim dZY As Double
    Dim dZZ As Double
    Dim dXX As Double
    Dim dXY As Double
    Dim dXZ As Double
    Dim oPoint As IJPoint
    Dim oGeom3DFactory As New GeometryFactory
    Dim vArrHoles() As Variant
    Dim lHoles As Long
    
    dOrigX = 0#
    dOrigY = 0#
    dOrigZ = 0#
    
    Set oSmartOcc = oAggregatorDesc.CAO
    Set oRefCollAsm = GetRefCollection(oSmartOcc)
    iRCCount = oRefCollAsm.IJDEditJDArgument.GetCount
    
    Set oRefColl1 = oRefCollAsm.IJDEditJDArgument.GetEntityByIndex(1)
    
    ' get fndn ports on the first RefColl (Fnd ports RC)
    ' for each of the foundation ports check if they were migrated
    '    revision mgr returns the collection of replacing objects
    '    in current design we see a one to one swap (fndn port of the clone with a new fndn port) during pump mirror copy
    '    add the replacing ports to the new migrated ports collection
    ' if there are any migrated ports remove existing fndn ports on first RefColl
    ' and add the newly migrated foundation ports.
    
    Dim FoundationPorts As IJElements
    Set FoundationPorts = New JObjectCollection
         
    Dim pEnumJDArgument As IEnumJDArgument
    Set pEnumJDArgument = oRefColl1
    
    Dim arg1 As IJDArgument
    Dim oPort As Object
    Dim found As Long
    Dim iloop As Long
    
    If Not pEnumJDArgument Is Nothing Then
        pEnumJDArgument.Reset
        iloop = 1
        Do
           pEnumJDArgument.Next 1, arg1, found
           If found <> 0 Then
                Set oPort = arg1.Entity
                FoundationPorts.Add oPort
                Set oPort = Nothing
                Set arg1 = Nothing
                iloop = iloop + 1
           Else: Exit Do
           End If
        Loop
    End If
    
    Dim MigratedFoundationPorts As IJElements
    Set MigratedFoundationPorts = New JObjectCollection
    
    Dim pObjectReplacingColl As IJDObjectCollection
    If iloop >= 1 Or FoundationPorts.count >= 1 Then
    
        ' check if any of the fndn ports are being migrated
        Dim oFndnPort As IJEqpFoundationPort
        Dim bIsDeleted As Boolean
        
        For Each oFndnPort In FoundationPorts
            oMigrateHelper.ObjectsReplacing oFndnPort, pObjectReplacingColl, bIsDeleted
            
            If pObjectReplacingColl.count > 0 Then
    
                Dim pObjectReplacing As Object
                For Each pObjectReplacing In pObjectReplacingColl
                    If TypeOf pObjectReplacing Is IJEqpFoundationPort Then
                        MigratedFoundationPorts.Add pObjectReplacing
                    End If
                Next
                
            End If
            pObjectReplacingColl.Clear
        Next
        
        'If any of the fndn port inputs are migrated, reset them on the ref coll
        If MigratedFoundationPorts.count >= 1 Then
        
            Call oRefColl1.IJDEditJDArgument.RemoveAll
            Dim indx As Integer
            
            For indx = 1 To MigratedFoundationPorts.count
                oRefColl1.IJDEditJDArgument.SetEntity indx, MigratedFoundationPorts.Item(indx), IJEqpFoundationPort, "SPSFndnPortToRC_DEST"
            Next indx
        
        End If
        
        Exit Sub ' need to exit here as fndn ports migrate is done
        
    End If
        
    Set oEqpFndPort = oRefColl1.IJDEditJDArgument.GetEntityByIndex(1)
    
    Set oPoint = oGeom3DFactory.Points3d.CreateByPoint(Nothing, 0#, 0#, 0#)
    Set oGeom3DFactory = Nothing
    
    lHoles = oEqpFndPort.NumberOfHoles
    If lHoles > 0 Then
    ' Get the position of first hole if hole count is positive. Aproximate point location
        Dim oHoleLocations() As DPosition
        Dim oElemsEqpFndPort As IJElements
        Dim Trans As IJDT4x4
        Dim iNumHoles As Integer
        Dim oHolePos As DPosition
                
        Set oElemsEqpFndPort = New JObjectCollection
        oElemsEqpFndPort.Add oEqpFndPort
        Call GetTransformAndHoles(oHoleLocations(), iNumHoles, oElemsEqpFndPort, Trans)
        Set oHolePos = Trans.TransformPosition(oHoleLocations(1))
        dOrigX = oHolePos.x
        dOrigY = oHolePos.y
        dOrigZ = oHolePos.z
    Else
    'Get the coordinate system for the foundation port . maybe close enough
        oEqpFndPort.GetCS dOrigX, dOrigY, dOrigZ, dXX, dXY, dXZ, dZX, dZY, dZZ
    End If
    oPoint.SetPoint dOrigX, dOrigY, dOrigZ
    
    'Only call migrate for the refcoll with support plane which will be second refcoll
    If iRCCount > 1 Then
        Set oRefColl2 = oRefCollAsm.IJDEditJDArgument.GetEntityByIndex(2)
        MigrateRefColl oRefColl2, oMigrateHelper, oObjectsReplacing, bIsInputMigrated, oPoint
     
        If bIsInputMigrated And UBound(oObjectsReplacing) > 0 Then
            'If any of the inputs are indeed migrated, reset them on the ref coll
            Call oRefColl2.IJDEditJDArgument.RemoveAll
            oRefColl2.IJDEditJDArgument.SetEntity 1, oObjectsReplacing(1), IJPlane, "SPSSuppPlaneToRC_DEST"
        End If
    End If
    
    Erase oHoleLocations

  Exit Sub
ErrorHandler:  HandleError MODULE, MT
End Sub

Public Function CreateSolidAsPlanes(pConnection As IUnknown, _
                                                      pCSLine As IJLineString, _
                                                      traceVec As IJDVector, _
                                                      SolidHeight As Double) As IJElements
Const METHOD = "CreateSolidAsPlanes"
 
    On Error GoTo ErrorHandler
    
    Dim oGeomFactory As New IngrGeom3D.GeometryFactory
    Dim oCSCmpxCurve As ComplexString3d, oTraceCmpxCurve As ComplexString3d
    Dim oCScurve As Line3d, oTraceCurve As Line3d
    Dim pX As Double, pY As Double, pZ As Double
    Dim ii As Integer, j As Integer
    Dim numPts As Long
    Dim dCSpts() As Double
        
    ' create a complex string and add cross section to it
    Dim iElements As IJElements
    Set iElements = New JObjectCollection
    
    pCSLine.GetPoints numPts, dCSpts
    j = 0
    
    For ii = 1 To numPts - 1
        Set oCScurve = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                                            dCSpts(j * 3), dCSpts(1 + j * 3), dCSpts(2 + j * 3), _
                                                            dCSpts(3 + j * 3), dCSpts(4 + j * 3), dCSpts(5 + j * 3))
        If Not oCScurve Is Nothing Then iElements.Add oCScurve
        Set oCScurve = Nothing
        j = j + 1
    Next ii
    
    Set oCSCmpxCurve = oGeomFactory.ComplexStrings3d.CreateByCurves(Nothing, iElements)
    
    ' create a trace curve
    Set oTraceCurve = oGeomFactory.Lines3d.CreateByPtVectLength(Nothing, _
                                                            dCSpts(0), dCSpts(1), dCSpts(2), _
                                                            traceVec.x, traceVec.y, traceVec.z, _
                                                            SolidHeight)
    ' create solid using singlesweep
    ' use break both trace and cross section option and with caps
    Dim stNorm() As Double, endNorm() As Double
    Dim numCaps As Long
    Set CreateSolidAsPlanes = oGeomFactory.GeometryServices.CreateBySingleSweepWCapsOpts(pConnection, _
                                                        oTraceCurve, oCSCmpxCurve, _
                                                        CircularCorner, BreakPathAndCrossSection, _
                                                        StartAtTraceBeg, SkinningCrossSectionOrientation.TraditionalOrientation, stNorm, endNorm, _
                                                        True, numCaps)
    Set oGeomFactory = Nothing
    Set oTraceCmpxCurve = Nothing
    Set oCSCmpxCurve = Nothing
    Set oCScurve = Nothing
    Set iElements = Nothing
    Set oTraceCurve = Nothing
    
    Exit Function

ErrorHandler:
    HandleError MODULE, METHOD
End Function

Public Sub CalculateHeightFromSupporingPlane(ByVal SupportPlane As IJPlane, Trans As IJDT4x4, HoleLocations() As DPosition, ByVal zaxis As DVector, ByVal eqPos As IJDPosition, ByRef newHeight As Double)
    ' calculate the block ht from the support plane
        Dim x As Double, y As Double, z As Double
        Dim x1 As Double, y1 As Double, z1 As Double
        SupportPlane.GetNormal x1, y1, z1
        SupportPlane.GetRootPoint x, y, z

        Dim suppN As New DPosition
        Dim suppR As New DPosition

        suppN.Set x1, y1, z1
        suppR.Set x, y, z

        zaxis.Length = 1     ' normalize

        Call CalculateDistanceFromPtToSurface(eqPos, zaxis, suppN, suppR, newHeight)
        
        If UBound(HoleLocations) > 1 Then
            ' check for greatest distance from port coords to support plane (non parallel case)
            Dim jj As Integer
            Dim xloc As Double, yloc As Double, zloc As Double
            Dim holePos As DPosition
            Dim dHt As Double
            Set holePos = New Automath.DPosition

            For jj = 1 To UBound(HoleLocations)
                xloc = HoleLocations(jj).x
                yloc = HoleLocations(jj).y
                zloc = HoleLocations(jj).z

                holePos.Set 0, 0, 0
                holePos.Set xloc, yloc, zloc
                Set holePos = Trans.TransformPosition(holePos)

                dHt = 0#
                Call CalculateDistanceFromPtToSurface(holePos, zaxis, suppN, suppR, dHt)
                If dHt >= newHeight Then newHeight = dHt
            Next jj
            Set holePos = Nothing
        End If
'
        Set suppN = Nothing
        Set suppR = Nothing
Exit Sub
End Sub

Public Sub CalculateHeightFromOriginToSupporingPlane(ByVal SupportPlane As IJPlane, ByVal eqPos As IJDPosition, ByVal zaxis As DVector, ByRef newHeight As Double)
Const METHOD = "CalculateHeightFromOriginToSupporingPlane"
On Error GoTo ErrorHandler
    ' calculate the ht from the support plane
    Dim x As Double, y As Double, z As Double
    Dim x1 As Double, y1 As Double, z1 As Double
    
    If Not SupportPlane Is Nothing Then
        SupportPlane.GetNormal x1, y1, z1
        SupportPlane.GetRootPoint x, y, z
    End If
    
    Dim suppN As New DPosition
    Dim suppR As New DPosition
    suppN.Set x1, y1, z1
    suppR.Set x, y, z

    zaxis.Length = 1     ' normalize
    Call CalculateDistanceFromPtToSurface(eqPos, zaxis, suppN, suppR, newHeight)
    
    Set suppN = Nothing
    Set suppR = Nothing
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub GetCentroidOfAllHoles(ByVal FoundationPorts As IJElements, ByRef centroid As DPosition)
Const METHOD = "GetCentroidOfAllHoles"
On Error GoTo ErrorHandler
    ' calculate the centroid of all the holes.
    Dim port As IJEqpFoundationPort
    Dim allHoles() As DPosition
    Dim tempHoles() As Variant
    Dim CurrentHole As Integer
    CurrentHole = 1
    Dim a As Integer
    ReDim allHoles(0) As DPosition
    For Each port In FoundationPorts
        If port.NumberOfHoles > 0 Then
            ReDim Preserve allHoles(UBound(allHoles) + port.NumberOfHoles) As DPosition
            Call port.GetHoles(tempHoles())
            For a = LBound(tempHoles) To UBound(tempHoles)
                Set allHoles(CurrentHole) = New DPosition
                allHoles(CurrentHole).Set tempHoles(a, 1), tempHoles(a, 2), 0
                CurrentHole = CurrentHole + 1
            Next a
        End If
    Next
    Dim centroidX As Double
    Dim centroidY As Double
    Dim centroidZ As Double

    GetCentroidOfPositions allHoles, UBound(allHoles), centroidX, centroidY, centroidZ

    Set centroid = New DPosition
    centroid.Set centroidX, centroidY, centroidZ
    
    Erase allHoles
Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub GetPortAndHoleLocations(pEnumJDArgument As IEnumJDArgument, ByRef FoundationPorts As IJElements, ByRef HoleLocations() As Automath.DPosition, ByRef bPtOption As Boolean)
Const METHOD = "GetPortAndHoleLocations"
On Error GoTo ErrorHandler
    ReDim HoleLocations(0) As Automath.DPosition
    Dim arg1 As IJDArgument
    Dim found As Long
    Dim iloop As Long
    Dim oPort As Object
    Dim oPt As IJPoint
    If Not pEnumJDArgument Is Nothing Then
        pEnumJDArgument.Reset
        iloop = 1
        Do
           pEnumJDArgument.Next 1, arg1, found
           If found <> 0 Then
                Set oPort = arg1.Entity
                If TypeOf oPort Is IJPoint Then
                    bPtOption = True
                    Set oPt = oPort
                    Dim ptX As Double, ptY As Double, ptZ As Double
                    oPt.GetPoint ptX, ptY, ptZ
                    ReDim Preserve HoleLocations(UBound(HoleLocations) + 1) As DPosition
                    Set HoleLocations(iloop) = New DPosition
                    HoleLocations(iloop).Set ptX, ptY, ptZ
                    Set oPt = Nothing
                End If
                FoundationPorts.Add oPort
                Set oPort = Nothing
                Set arg1 = Nothing
                iloop = iloop + 1
           Else: Exit Do
           End If
        Loop
    End If
    
    If FoundationPorts.count = 0 Then
        bPtOption = True
    End If
    
    ' tr 77250 raise an error when all fnd port inputs are deleted
    If iloop = 0 Then
        SPSToDoErrorNotify EqpToDoMsgCodelist, TDL_EQPFND_NOTASSOC_WITHEQP, pEnumJDArgument, Nothing
        Err.Raise SPS_MACRO_WARNING 'E_FAIL
    End If


    
Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Function GetOriginShiftForUnevenSupportedObjects(HoleLocations() As DPosition, _
                                                        ByVal eqPos As DPosition, _
                                                        ByVal zVec As DVector, _
                                                        transMatrix As IJDT4x4) As Double
Const METHOD = "GetOriginShiftForUnevenSupportedObjects"
On Error GoTo ErrorHandler

    ' In case the supported objects(equipments) are not in same heights then adjust the EF also
    ' so that the EF should start from the bottom most point in its normal dir.
    ' therefore the origin of the EF should be shifted to the lower most point along the normal.
    ' Identify the maximum distance between the origin and all the hole locations
    ' along the direction of EF projection.
    Dim xloc As Double, yloc As Double, zloc As Double
    Dim dMaxDiff As Double
    
    Dim holePos As DPosition
    Set holePos = New Automath.DPosition
    xloc = HoleLocations(1).x
    yloc = HoleLocations(1).y
    zloc = HoleLocations(1).z
    holePos.Set xloc, yloc, zloc
    Set holePos = transMatrix.TransformPosition(holePos)
    
    Dim normalPt As New DPosition
    normalPt.Set zVec.x, zVec.y, zVec.z
    Call CalculateDistanceFromPtToSurface(eqPos, zVec, normalPt, holePos, dMaxDiff)
    
    Dim jj As Integer
    For jj = 2 To UBound(HoleLocations)
       xloc = HoleLocations(jj).x
       yloc = HoleLocations(jj).y
       zloc = HoleLocations(jj).z
       
       holePos.Set xloc, yloc, zloc
       Set holePos = transMatrix.TransformPosition(holePos)
       Dim dHolePosHt As Double
       Call CalculateDistanceFromPtToSurface(eqPos, zVec, normalPt, holePos, dHolePosHt)
       If Not dHolePosHt = 0 And dHolePosHt > dMaxDiff Then dMaxDiff = dHolePosHt
    Next jj
    Set holePos = Nothing
    GetOriginShiftForUnevenSupportedObjects = dMaxDiff
    
Exit Function
ErrorHandler:
    HandleError MODULE, METHOD
End Function


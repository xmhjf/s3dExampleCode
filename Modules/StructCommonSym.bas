Attribute VB_Name = "StructCommonSym"
Option Explicit
'*******************************************************************
'
'Copyright (C) 2006 Intergraph Corporation. All rights reserved.
'
'File : StructCommonSym.bas
'
'Author : A Patil
'
'Description :
'    SmartPlant Structural common symbol functions
'
'History:
'
' 05/21/03   JMS Replaced msgbox with JServErrors
' 09-May-06  JMS DI#97751 - All lines are constructed through the
'                   geometry factory instead of being "new"ed'
' 08-Aug-06 AS  TR# 99969 Added common footing migration method
'
' 09-Aug-06   SS DI#103385 - modified getproxy call in AddRelationShip to use
'                PreferProxyByAliasName = true to get named proxy for material rather than
'                the proxy itself and use it to establish StructEntityMaterial relation.
'
'  19-Sep-06  SS    TR#105983 - add each of the member outputs to IJAssembly for it to have
'                   proper hierarchy in Assembly tab.
'
' 06-Mar-07  RS&SS  CR#41094 - Changes for placing a footing in space and dis/reconnect to members
' 15-Apr-14  RRK    TR-CP-250764-Made change to SetOccurrenceMatrix method to solve Recorded Exception Minidump
'
'********************************************************************
Private Const MODULE = "StructCommonSym"
Public Const MODELDATABASE = "Model"
Private Const CATALOGDATABASE = "Catalog"
Public Const IJGeometry = "{96eb9676-6530-11d1-977f-080036754203}"
Public Const IJPlane = "{4317C6B3-D265-11D1-9558-0060973D4824}"
Public Const IJDAttributes = "{B25FD387-CFEB-11D1-850B-080036DE8E03}"
Public Const IJStructElevationDatum = "{A632CD84-0C10-4607-9910-AA13A1CD3E10}"
Public Const IJDOccurrence = "{274317DB-0F9D-11D2-94AD-080036CD8E03}"
Public Const IJStructMaterial = "{93F52983-9504-11D4-9D40-00105AA5BAEB}"
Public Const IJWeightCG = "{DC284B37-5A00-11D2-BE2B-0800364AA003}"
Public Const IJGraphicEntity = "{4FA5DC20-133E-11D1-9730-080036104103}"
Public Const IJSmartOccurrence = "{A2A655C0-E2F5-11D4-9825-00104BD1CC25}"
Public Const ISPSMemberPartPrismatic = "{53B6E606-78C6-4BCA-A640-43A2258EDED1}"
Public Const ISPSMemberSystemPhysicalAxis = "{A250A0A0-17A1-42c2-A1B2-B74E3B3CF131}"
Public Const ISPSMemberSystemXSectionNotify = "{39314DCB-901B-46c0-9521-500F695F4A08}"
Public Const ISPSAxisRotation = "{56CCD3A5-B756-4ab0-9C68-1F586D1E7C66}"
Public Const ISPSMemberSystemSuppingNotify1 = "{88CCCC19-BB35-431f-BCCE-27030A21D028}"
Public Const ISPSMemberSystemSuppingNotify2 = "{C155EED1-B0D8-41a3-B7A0-9526ACD67E2D}"
Public Const AssemblyMembers1RelationshipCLSID = "{45E4020F-F8D8-47A1-9B00-C9570C1E0B17}"
Public Const IJFullObject = "{bcbfb3c0-98c2-11d1-93de-08003670a902}"
Public Const IJControlPoint = "{F54DCDA8-2C24-47B9-A4ED-ACA4F3DF15D7}"
Public Const IJVolume = "{}"

Public Const FOOTINGCOMPPROGID = "SPSFootings.SPSFootingComponent"

Public Const TOL = 0.001
Public Const PI = 3.14159265358979

Public Const E_FAIL = -2147467259
Public Const DOUBLE_VALUE = 8
Public Const BOOL = -7
Public Const CHAR = 1
Private oLocalizer As IJLocalizer





'*************************************************************************
'Function
'AddRelationShip
'
'Abstract
'Establishes the relation between the input object and input material
'input
'Footing object, Material object, relationship name
'
'Return
'
'Exceptions
'
'***************************************************************************
Public Sub AddRelationShip(FtgObj As Object, MatlObj As Object, RelName As String)
Const METHOD = "AddRelationShip"
On Error GoTo ErrorHandler
    
    Dim oIJDAssocRelation As IJDAssocRelation
    Dim oTargetObjCol As IJDTargetObjectCol
    Dim oRelship As IMSRelation.DRelationshipHelper
    Dim iProxy As IJDProxy
    Dim oModelPOM As IJDPOM
    Dim ijdObject As ijdObject
    
    If FtgObj Is Nothing Or MatlObj Is Nothing Then
        Exit Sub
    End If
    If Not TypeOf FtgObj Is ijdObject Then
        Exit Sub
    End If
   
    Set ijdObject = FtgObj
    If ijdObject Is Nothing Then
        Exit Sub
    End If
    Set oModelPOM = ijdObject.ResourceManager
    Set iProxy = oModelPOM.GetProxy(MatlObj, True)
    
    Set oIJDAssocRelation = FtgObj
        
    Set oTargetObjCol = oIJDAssocRelation.CollectionRelations(STRUCT_MATERIAL, STRUCT_ENTITY_MATERIAL_ORIG)
     
        Dim pUnk As Object
        On Error Resume Next
        Set pUnk = oTargetObjCol.Item(RelName)
        If pUnk Is Nothing Then
            Call oTargetObjCol.Add(iProxy, RelName, oRelship)
        ElseIf Not pUnk Is MatlObj Then
            Call oTargetObjCol.Remove(RelName)
            Call oTargetObjCol.Add(iProxy, RelName, oRelship)
        End If
        
Exit Sub
ErrorHandler: HandleError MODULE, METHOD
End Sub
'*************************************************************************
'Function
'ConnectSmartOccurrence
'
'Abstract
'Establishes the relation between smart occrrence and reference collection
'input
'Smart occurrence object, reference collection
'
'Return
'
'Exceptions
'
'***************************************************************************
Public Sub ConnectSmartOccurrence(pSO As IJSmartOccurrence, pRefColl As IJDReferencesCollection)
Const METHOD = "ConnectSmartOccurrence"
 On Error GoTo ErrorHandler
  
   'connect the reference collection to the smart occurrence
    Dim pRelationHelper As IMSRelation.DRelationHelper
    Dim pCollectionHelper As IMSRelation.DCollectionHelper
    Dim pRelationshipHelper As DRelationshipHelper
    Dim pRevision As IJRevision
    
    Set pRelationHelper = pSO
    Set pCollectionHelper = pRelationHelper.CollectionRelations(IJSmartOccurrence, "toArgs_O")
    pCollectionHelper.Add pRefColl, "RC", pRelationshipHelper
    Set pRevision = New JRevision
    pRevision.AddRelationShip pRelationshipHelper
  
 Exit Sub
ErrorHandler: HandleError MODULE, METHOD
End Sub
'*************************************************************************
'Function
'GetRefCollection
'
'Abstract
'gets the reference collection that the smart occurrence is related to.
'The reference collection contains the input object on which the computation of
'smart occurrence is dependent
'
'input
'Smart occurrence object
'
'Return
'Reference collection
'Exceptions
'
'***************************************************************************
Public Function GetRefCollection(pSO As IJSmartOccurrence) As IJDReferencesCollection
Const METHOD = "GetRefCollection"
On Error GoTo ErrorHandler
'IJSmartOccurence GUID
'{A2A655C0-E2F5-11D4-9825-00104BD1CC25}
  
     Dim pRelationHelper As IMSRelation.DRelationHelper
     Dim pCollectionHelper As IMSRelation.DCollectionHelper
     Set pRelationHelper = pSO
     On Error Resume Next
     Set pCollectionHelper = pRelationHelper.CollectionRelations(IJSmartOccurrence, "toArgs_O")
     On Error Resume Next
     Set GetRefCollection = pCollectionHelper.Item("RC")
      
Exit Function
ErrorHandler: HandleError MODULE, METHOD
End Function
'*************************************************************************
'Function
'IsSOOverridden
'
'Abstract
'checks if the passed in interface is overridden. It is overridden if one
'of the value is not empty.When an object is created associated user attributes are
'initially empty. This method checks for it.
'Arguments:
'input collection corresponding to the interface
'
'Return
'True - overridden
'False- not overridden
'Exceptions
'
'***************************************************************************
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
'*************************************************************************
'Function
'CopyValuesToSOFromItem
'
'Abstract
'copies values form one collection to another
'Arguments:
'output collection, input collection
'
'Return
'
'
'Exceptions
'
'***************************************************************************
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
'*************************************************************************
'Function
'GetResourceMgr
'
'Abstract
'gets the resource manager
'Arguments:
'none
'
'Return
'object
'
'Exceptions
'
'***************************************************************************
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
'*************************************************************************
'Function
'GetDefinition
'
'Abstract
'gets the object from name
'Arguments:
'name of object
'
'Return
'object
'
'Exceptions
'
'***************************************************************************
Public Function GetDefinition(ByVal name As String) As Object
Const MT = "GetDefinition"
On Error GoTo ErrHandler
    Dim iPom As IJDPOM
    Set iPom = GetCatalogResourceManager()
    Set GetDefinition = iPom.GetObject(name)
    Exit Function
ErrHandler:
      HandleError MODULE, MT
End Function
'*************************************************************************
'Function
'InitRectCurvePoints
'
'Abstract
'constructs a corner points of the rectangle from width and height
'Arguments:
'width, height, ZDir
'
'Return
'corner points of the rectangle
'
'Exceptions
'
'***************************************************************************
Public Sub InitRectCurvePoints(pts() As Double, width As Double, Height As Double, ZDirec)
Const METHOD = "InitRectCurvePoints"
 'Build points in local XY plane at the centroid of the rectangle
    pts(0) = -(width / 2#)
    pts(1) = -(Height / 2#)
    pts(2) = ZDirec
    
    pts(3) = width / 2#
    pts(4) = -(Height / 2#)
    pts(5) = ZDirec
    
    pts(6) = width / 2#
    pts(7) = Height / 2#
    pts(8) = ZDirec
    
    pts(9) = -(width / 2#)
    pts(10) = Height / 2#
    pts(11) = ZDirec
    
    pts(12) = pts(0) 'Same as first point for closed shape
    pts(13) = pts(1)
    pts(14) = pts(2)
Exit Sub
 
End Sub
'*************************************************************************
'Function
'GetCAODefAttribute
'
'Abstract
'Gets the attribute value
'Arguments:
'Object, interface, attribute name
'
'Return
'attribute value
'
'Exceptions
'
'***************************************************************************
Public Function GetCAODefAttribute(ByVal pMemberDescription As IJDMemberDescription, _
                              InterfaceName As String, AttributeName As String) As String
Const METHOD = "GetCAODefAttribute"
On Error GoTo ErrHandler
     Dim oSmartOcc As IJSmartOccurrence
     Dim oSmartItem As IJSmartItem
     Dim Attrs As IJDAttributes
     Dim SlabComp As String
     Set oSmartOcc = pMemberDescription.CAO
     Set oSmartItem = oSmartOcc.ItemObject
     Set Attrs = oSmartItem
     GetCAODefAttribute = Attrs.CollectionOfAttributes(InterfaceName).Item(AttributeName).Value
Exit Function
ErrHandler:
         HandleError MODULE, METHOD
End Function
'*************************************************************************
'Function
'GetModelCPpositionPart
'
'Abstract
'Gets the X,Y of the cardinal point, in member coordinates, at the start or end of the member part
'Arguments:
'CP, member part, end
'
'Return
'X,Y
'
'Exceptions
'
'***************************************************************************
Public Sub GetModelCPPositionPart(CP As Long, oMembPart As ISPSMemberPartPrismatic, x As Double, y As Double, bEnd As Boolean)
Const METHOD = "GetModelCPpositionPart"
On Error GoTo ErrorHandler

     If CP <> 5 Then
          Dim offX As Double, offY As Double
          Dim offX1 As Double, offY1 As Double
          Dim Pos As IJDPosition
          
          Dim a As Double, b As Double, c As Double  'Start point of oMembPart
          Dim a1 As Double, b1 As Double, c1 As Double  ' End Point of oMembPart
          
          Set Pos = New DPosition
          Dim NewPos As IJDPosition
          Set NewPos = New DPosition
          Dim MMatrix As IJDT4x4
          Set MMatrix = New DT4x4
          
          Dim MMatrixEnd As IJDT4x4
          Set MMatrixEnd = New DT4x4
          
          Dim MMatrixEndA As IJDT4x4
          Set MMatrixEndA = New DT4x4
          
          oMembPart.CrossSection.GetCardinalPointOffset CP, offX, offY
          oMembPart.CrossSection.GetCardinalPointOffset 5, offX1, offY1
          oMembPart.Rotation.GetTransform MMatrix
          offX1 = offX1 - offX
          offY1 = offY1 - offY
          If bEnd Then
          Pos.x = oMembPart.Axis.length
          Else
          Pos.x = 0#
          End If
          'added to handle reflect- TR#66921
          If oMembPart.Rotation.Mirror = True Then
             offX1 = -offX1
             offY1 = offY1
          End If
          'added to handle reflect
          
          Pos.y = -offX1
          Pos.z = offY1
          Set NewPos = MMatrix.TransformPosition(Pos)
          
          'DI#84020 Added this Check for Curved members
          If TypeOf oMembPart Is ISPSMemberPartCurve And bEnd Then 'curved member
                Pos.x = 0
                oMembPart.Axis.EndPoints a, b, c, a1, b1, c1
                oMembPart.Rotation.GetTransformAtPosition a1, b1, c1, MMatrixEnd, MMatrixEndA
                Set NewPos = MMatrixEnd.TransformPosition(Pos)
          End If
          'DI#84020 Ends
          x = NewPos.x
          y = NewPos.y
     End If

Exit Sub
ErrorHandler:  HandleError MODULE, METHOD
End Sub
'*************************************************************************
'Function
'CommonCheckMember
'
'Abstract
'Checks if the passed in object is a member system
'Arguments:
'object, collection
'
'Return
'1 - passed in object is member systems
'0 - passed in object is member systems
'
'Exceptions
'
'***************************************************************************
Public Function CommonCheckMember(obj As Object, elems As IJElements) As Integer
Const METHOD = "CommonCheckMember"
On Error GoTo ErrorHandler
  
    CommonCheckMember = 0
    Dim oMembSys As ISPSMemberSystem
    On Error Resume Next
    Set oMembSys = obj
    If oMembSys Is Nothing Then
        Exit Function
    Else
        If IsBuiltUPMemeberSystem(oMembSys) Then
            Exit Function
        End If
        CommonCheckMember = 1
        Exit Function
    End If
    
    

Exit Function
ErrorHandler:  HandleError MODULE, METHOD
End Function
'*************************************************************************
'Function
'CommonCheckBottomPlane
'
'Abstract
'Checks if the passed in members are below the input plane
'Arguments:
'Plane, member systems
'
'Return
'1 - member systems are below the plane
'0 - they are not below the plane
'
'Exceptions
'
'***************************************************************************
'TR#71850- Last argument added. This contains object ins selectset. This is used to check that
' plane doesn't belong to same footing which is being modified.
Public Function CommonCheckBottomPlane(obj As Object, elems As IJElements, ObjsinSelectSet As IJElements) As Integer
Const METHOD = "CommonCheckBottomPlane"
On Error GoTo ErrorHandler

    CommonCheckBottomPlane = 0
    Dim oMembSys As ISPSMemberSystem
    Dim oCurve As IJCurve
    Dim oPlane As IJPlane
    On Error Resume Next
    Set oPlane = obj
    If oPlane Is Nothing Then
    Exit Function
    End If
    On Error GoTo ErrorHandler
    
    Dim FootingObj As ISPSFooting
    If ObjsinSelectSet Is Nothing Then Exit Function
    If ObjsinSelectSet.count > 0 Then
        If TypeOf ObjsinSelectSet.Item(1) Is ISPSFooting Then
            Set FootingObj = ObjsinSelectSet.Item(1)
        End If
            
        If TypeOf obj Is IJDReferenceProxy Then
            Dim oRefProxy As IJDReferenceProxy
            Dim footingComp As ISPSFootingComponent
            Dim systemChild As IJDesignChild
            Dim desParent As IJDesignParent

            Set oRefProxy = obj
            If TypeOf oRefProxy.Reference Is ISPSFootingComponent Then
                Set footingComp = oRefProxy.Reference
                Set systemChild = footingComp
                Set desParent = systemChild.GetParent
            ElseIf TypeOf oRefProxy.Reference Is ISPSFooting Then
                Set desParent = oRefProxy.Reference
            End If
            
            If desParent Is FootingObj Then
                Set oRefProxy = Nothing
                Set FootingObj = Nothing
                Set footingComp = Nothing
                Set systemChild = Nothing
                Set desParent = Nothing
                CommonCheckBottomPlane = 0
                Exit Function
            End If
  
        End If
        
        Set oRefProxy = Nothing
        Set footingComp = Nothing
        Set systemChild = Nothing
        Set desParent = Nothing
    End If
    'TR#71850
    
    Dim i As Integer, NumInter As Long, Numover As Long
    Dim pts() As Double
    Dim Stx As Double, Sty As Double, Stz As Double, Endx As Double, Endy As Double, Endz As Double
    Dim Srcsx As Double, Srcy As Double, Srcz As Double, inx As Double, iny As Double, inz As Double
    Dim Rtx As Double, Rty As Double, Rtz As Double
    Dim Normx As Double, Normy As Double, Normz As Double
    
    Dim oline As Line3d
    Dim NewCurve As IJCurve
    Dim DummyFace As New Plane3d
    Dim GeomFactory As New IngrGeom3D.GeometryFactory
    Dim code As Geom3dIntersectConstants
    
    oPlane.GetRootPoint Rtx, Rty, Rtz
    oPlane.GetNormal Normx, Normy, Normz

    Set DummyFace = GeomFactory.Planes3d.CreateByPointNormal(Nothing, Rtx, Rty, Rtz, Normx, Normy, Normz)
        
    ' check the lowest point when in modify and there is no points and no members
    If Not FootingObj Is Nothing And elems.count = 0 Then
        FootingObj.GetPosition Stx, Sty, Stz
        If (Rtz < Stz) Then
            CommonCheckBottomPlane = 1
        Else
            CommonCheckBottomPlane = 0
        End If
        Exit Function
    End If
    
    ' get/check the lowest points when points have been picked
    Dim oPoint As IJDPosition
    If elems.count > 0 Then
        If TypeOf elems.Item(1) Is IJDPosition Then
            Dim tempX As Double, tempY As Double, tempZ As Double
            Set oPoint = elems.Item(1)
            Stz = oPoint.z
            For i = 2 To elems.count
                If TypeOf elems.Item(i) Is IJDPosition Then
                    Set oPoint = elems.Item(i)
                    tempZ = oPoint.z
                End If
                If tempZ < Stz Then
                    Stz = tempZ
                End If
            Next i
            
            If (Rtz < Stz) Then
                CommonCheckBottomPlane = 1
            Else
                CommonCheckBottomPlane = 0
            End If

            Exit Function
        End If
    End If


    For i = 1 To elems.count
        If TypeOf elems.Item(i) Is ISPSMemberSystem Then
            Set oMembSys = elems.Item(i)
        End If
        
        If Not oMembSys Is Nothing Then
            Set oCurve = oMembSys
            oCurve.EndPoints Stx, Sty, Stz, Endx, Endy, Endz
            
            If (Endz > Stz And Rtz >= Stz) Or (Stz > Endz And Rtz >= Endz) Then
                CommonCheckBottomPlane = 0
                Exit For
            End If ' TR#67109
            
            Set oline = GeomFactory.Lines3d.CreateByPtVectLength(Nothing, Stx, Sty, Stz, _
                                    Endx - Stx, Endy - Sty, Endz - Stz, 1)
            oline.Infinite = True
            Set NewCurve = oline
            NewCurve.Intersect DummyFace, NumInter, pts, Numover, code
            
            If NumInter >= 1 Then
                Dim NewPos As New DPosition
                Dim Vec1 As New DVector
                Dim Vec2 As New DVector
                NewPos.Set pts(0), pts(1), pts(2)
                Vec1.Set NewPos.x - Stx, NewPos.y - Sty, NewPos.z - Stz
                Vec2.Set NewPos.x - Endx, NewPos.y - Endy, NewPos.z - Endz
                If Vec1.Dot(Vec2) > 0# Then
                    If Normz >= 0.7071 And Normz <= 1 Then
                        CommonCheckBottomPlane = 1
                    ElseIf Normz <= -0.7071 And Normz >= -1 Then
                        CommonCheckBottomPlane = 1
                    End If
                Else
                    CommonCheckBottomPlane = 0
                End If
'            ElseIf Rtz < Endz And Rtz < Stz Then ' TR#71999- commented this. if plane doesn't intersect member line then that plane can not be valid plane.
'                CommonCheckBottomPlane = 1
            Else
                CommonCheckBottomPlane = 0
                Exit Function
            End If
        End If
    Next i
    
    Set FootingObj = Nothing
    
Exit Function
ErrorHandler:  HandleError MODULE, METHOD
End Function

Public Sub GetColumnsRange(elems As IJElements, _
                           ByRef xmin As Double, ByRef ymin As Double, ByRef xmax As Double, ByRef ymax As Double, ByRef zmin As Double)

Const METHOD = "GetColumnsRange"
On Error GoTo ErrorHandler

Dim i As Long
Dim x As Double, y As Double, z As Double
Dim dMembAng As Double
Dim sSecName As String, sRefStd As String

xmin = 100000000#
ymin = 100000000#
xmax = -100000000#
ymax = -100000000#
zmin = 100000000#


For i = 1 To elems.count
    GetMemberBottomCenterProperties elems, x, y, z, dMembAng, sSecName, sRefStd, i
            
    If xmin > x Then xmin = x
    If ymin > y Then ymin = y
    If xmax < x Then xmax = x
    If ymax < y Then ymax = y
    If zmin > z Then zmin = z
   
Next i

Exit Sub
ErrorHandler:  HandleError MODULE, METHOD
End Sub
Public Sub GetColumnsRangeAboutRotatedZ(elems As IJElements, dblZRotation As Double, _
                           ByRef xmin As Double, ByRef ymin As Double, ByRef xmax As Double, ByRef ymax As Double, ByRef zmin As Double, ByRef xMidGlobal As Double, ByRef yMidGlobal As Double)

Const METHOD = "GetColumnsRangeAboutRotatedZ"
On Error GoTo ErrorHandler

Dim i As Long
Dim x As Double, y As Double, z As Double
Dim dMembAng As Double
Dim sSecName As String, sRefStd As String
Dim oMatrix As New DT4x4
Dim oVec As New DVector
Dim oPos As New DPosition

oMatrix.LoadIdentity
oVec.Set 0, 0, 1


xmin = 100000000#
ymin = 100000000#
xmax = -100000000#
ymax = -100000000#
zmin = 100000000#

oMatrix.Rotate dblZRotation, oVec
oMatrix.Invert
For i = 1 To elems.count
    GetMemberBottomCenterProperties elems, x, y, z, dMembAng, sSecName, sRefStd, i
    oPos.Set x, y, z
    
    Set oPos = oMatrix.TransformPosition(oPos)
    oPos.Get x, y, z
    If xmin > x Then xmin = x
    If ymin > y Then ymin = y
    If xmax < x Then xmax = x
    If ymax < y Then ymax = y
    If zmin > z Then zmin = z
   
Next i
'get mid point alog rotated coord
xMidGlobal = (xmin + xmax) / 2
yMidGlobal = (ymin + ymax) / 2

'convert to global now
oMatrix.Invert
oPos.Set xMidGlobal, yMidGlobal, z
Set oPos = oMatrix.TransformPosition(oPos)
oPos.Get xMidGlobal, yMidGlobal, z

Exit Sub
ErrorHandler:  HandleError MODULE, METHOD
End Sub
Public Sub RotatePositionAboutZ(ByRef x As Double, ByRef y As Double, ByRef z As Double, dblZRotation As Double)

Const METHOD = "RotatePositionAboutZ"
On Error GoTo ErrorHandler



    Dim oMatrix As New DT4x4
    Dim oVec As New DVector
    Dim oPos As New DPosition

    oMatrix.LoadIdentity
    oVec.Set 0, 0, 1


    oMatrix.Rotate dblZRotation, oVec
    oMatrix.Invert
    oPos.Set x, y, z
    
    Set oPos = oMatrix.TransformPosition(oPos)
    oPos.Get x, y, z

Exit Sub
ErrorHandler:  HandleError MODULE, METHOD
End Sub


Public Sub InitOctCurvePoints(pts() As Double, OctOverAllDim As Double, ZDirec)
Const METHOD = "InitRectCurvePoints"
On Error GoTo ErrorHandler

Dim OctFaceLength As Double

  OctFaceLength = OctOverAllDim / (1 + 2 * Sin(PI / 4))
  
 'Build points in local XY plane at the centroid of the rectangle
    pts(0) = -(OctOverAllDim / 2#)
    pts(1) = (OctFaceLength / 2#)
    pts(2) = ZDirec
    
    pts(3) = -(OctOverAllDim / 2#)
    pts(4) = -(OctFaceLength / 2#)
    pts(5) = ZDirec
    
    pts(6) = -(OctFaceLength / 2#)
    pts(7) = -(OctOverAllDim / 2#)
    pts(8) = ZDirec
    
    pts(9) = (OctFaceLength / 2#)
    pts(10) = -(OctOverAllDim / 2#)
    pts(11) = ZDirec
    
    pts(12) = (OctOverAllDim / 2#)
    pts(13) = -(OctFaceLength / 2#)
    pts(14) = ZDirec
    
    pts(15) = (OctOverAllDim / 2#)
    pts(16) = (OctFaceLength / 2#)
    pts(17) = ZDirec
    
    pts(18) = (OctFaceLength / 2#)
    pts(19) = (OctOverAllDim / 2#)
    pts(20) = ZDirec
    
    pts(21) = -(OctFaceLength / 2#)
    pts(22) = (OctOverAllDim / 2#)
    pts(23) = ZDirec
    
    pts(24) = pts(0)
    pts(25) = pts(1)
    pts(26) = pts(2)
    
Exit Sub
ErrorHandler:   HandleError MODULE, METHOD
End Sub

'This subroutine return appropirate groutwidth/length components to be used for deriving pier size
' THis will be based on whether indiviual column is governing range box or not . If yes then get appropriate component
' based grout orienation angle and its width/length
Public Sub GetGroutWidthLengthComponents(pPropertyDescriptions As IJDPropertyDescription, _
                                             xmin As Double, xmax As Double, _
                                           ymin As Double, ymax As Double, dblPierRotation As Double, _
                                           ByRef groutwidthComp As Double, ByRef groutLengthComp As Double, _
                                           Optional oMerGePierRefColl As IMSSymbolEntities.IJDReferencesCollection)
                                           
Const METHOD = "GetGroutWidthLengthComponents"
On Error GoTo ErrorHandler

Dim i As Long
Dim x As Double, y As Double, z As Double
Dim oFooting As IJDMemberObjects
Dim oSmartOcc As IJSmartOccurrence
Dim bGroutNeeded As Boolean
Dim oReferencesCollection  As IMSSymbolEntities.IJDReferencesCollection
Dim oRefColl  As IMSSymbolEntities.IJDReferencesCollection
Dim GroutRotationAngle As Double
Dim GroutOrientation As Long
Dim RotationInherited As Double
Dim sSecName As String, sRefStd As String
Dim tmpwidthComp As Double, tmpLengthComp As Double
Dim tmpwidthComp1 As Double, tmpLengthComp1 As Double
Dim oAttr As IJDAttributes
Dim oGroutAttribs As IJDAttributes
Dim supported As IJElements

   On Error Resume Next
    Set oFooting = pPropertyDescriptions.CAO
    Set oAttr = oFooting
    bGroutNeeded = oAttr.CollectionOfAttributes(PIER_FOOTING_ASM).Item(WITH_GROUT_PAD).Value
    If Err.Number <> 0 Then
        Err.Clear
        bGroutNeeded = oAttr.CollectionOfAttributes(SLAB_FOOTING_ASM).Item(WITH_GROUT_PAD).Value
    End If
    
    If Err.Number <> 0 Then
        Err.Clear
        bGroutNeeded = oAttr.CollectionOfAttributes(PIER_SLAB_FOOTING_ASM).Item(WITH_GROUT_PAD).Value
    End If
    Err.Clear
    On Error GoTo ErrorHandler
    
    Set oRefColl = GetRefCollection(oFooting)
    Set oReferencesCollection = oRefColl.IJDEditJDArgument.GetEntityByIndex(1)
    Set oRefColl = Nothing
    Set supported = GetUpdatedRefColl(oReferencesCollection)
    
    If Not oMerGePierRefColl Is Nothing Then
        For i = 1 To oMerGePierRefColl.IJDEditJDArgument.GetCount
            supported.Add oMerGePierRefColl.IJDEditJDArgument.GetEntityByIndex(i)
        Next i
    End If
    
    
    For i = 1 To supported.count
        GetMemberBottomCenterProperties supported, x, y, z, RotationInherited, sSecName, sRefStd, i
        If TypeOf supported.Item(i) Is IJPoint Then
           RotationInherited = GetPlanAngleFromSymbolOccurrence(oFooting)
        End If
        
        RotatePositionAboutZ x, y, z, dblPierRotation
        If Abs(x - xmin) < TOL Or Abs(x - xmax) < TOL Or Abs(y - ymin) < TOL Or Abs(y - ymax) < TOL Then
            If bGroutNeeded = True Then
                Set oSmartOcc = oFooting.ItemByDispid(1, i)
                Set oGroutAttribs = oSmartOcc
                tmpLengthComp = oGroutAttribs.CollectionOfAttributes(FTG_GROUT_PAD_DIM).Item(GROUT_LENGTH).Value
                tmpwidthComp = oGroutAttribs.CollectionOfAttributes(FTG_GROUT_PAD_DIM).Item(GROUT_WIDTH).Value
                GroutRotationAngle = oGroutAttribs.CollectionOfAttributes(FTG_GROUT_PAD).Item(GROUT_ROTATION_ANGLE).Value
                GroutOrientation = oGroutAttribs.CollectionOfAttributes(FTG_GROUT_PAD).Item(GROUT_ORIENTATION).Value
                CheckForUndefinedValueAndRaiseError oFooting, GroutOrientation, STRUCT_COORD_SYS_REF, 123
                
                If GroutOrientation = 2 Then '2 is Local to the SPSMemberPart
                    GroutRotationAngle = GroutRotationAngle + RotationInherited
                End If
                
                Dim GroutShape As Long 'TR#66868
                GroutShape = oGroutAttribs.CollectionOfAttributes(FTG_GROUT_PAD).Item(GROUT_SHAPE).Value
                If GroutShape = 3 Then GroutRotationAngle = 0# ' if grout is circular then not need to consider its angle
            Else
                tmpLengthComp = GetCSAttribData(sSecName, sRefStd, CROSSSSECTION_DIMENSIONS, CROSSSECTION_DEPTH)
                tmpwidthComp = GetCSAttribData(sSecName, sRefStd, CROSSSSECTION_DIMENSIONS, CROSSSECTION_WIDTH)
                GroutRotationAngle = -RotationInherited
            End If
            

            
            GroutRotationAngle = Abs(Abs(GroutRotationAngle) - dblPierRotation)
            If GroutRotationAngle > PI / 2 Then GroutRotationAngle = Abs(PI - GroutRotationAngle)
            
            If Abs(x - xmin) < TOL Or Abs(x - xmax) < TOL Then ' This column is governing length of range box
                tmpLengthComp1 = tmpLengthComp * Cos(GroutRotationAngle) + tmpwidthComp * Sin(GroutRotationAngle)
            End If
            
            If Abs(y - ymin) < TOL Or Abs(y - ymax) < TOL Then ' This columns is governing width of range box
               tmpwidthComp1 = tmpwidthComp * Cos(GroutRotationAngle) + tmpLengthComp * Sin(GroutRotationAngle)
            End If
        End If
        
        If groutwidthComp < tmpwidthComp1 Then
           groutwidthComp = tmpwidthComp1
        End If
        If groutLengthComp < tmpLengthComp1 Then
           groutLengthComp = tmpLengthComp1
        End If
        
        Set oSmartOcc = Nothing
        Set oGroutAttribs = Nothing
    Next i

Set oFooting = Nothing
Set oAttr = Nothing
Set oReferencesCollection = Nothing
Set supported = Nothing


Exit Sub
ErrorHandler:  HandleError MODULE, METHOD
End Sub
'*************************************************************************
'Function
'GetFootingPositionFromMember
'
'Abstract
' given a member, gets the x, y, z

'
'***************************************************************************
Public Sub GetFootingPositionFromMember(oMembSys As ISPSMemberSystem, ByRef x As Double, _
                                        ByRef y As Double, ByRef z As Double, _
                                        ByRef dMembAng As Double)
                                        
        Dim x1 As Double, y1 As Double, z1 As Double
        Dim CP As Long
        Dim oMembPart As ISPSMemberPartPrismatic
        Set oMembPart = GetBottomMemberPart(oMembSys)
        oMembPart.Axis.EndPoints x, y, z, x1, y1, z1
        CP = oMembPart.CrossSection.CardinalPoint
        dMembAng = oMembPart.Rotation.BetaAngle
        
        Dim bEnd As Boolean
        If z < z1 Then
        Else
             bEnd = True
             x = x1
             y = y1
             z = z1
             dMembAng = -dMembAng
        End If
        
        If CP <> 5 Then
            If Not oMembPart Is Nothing Then
                Call GetModelCPPositionPart(CP, oMembPart, x, y, bEnd)
            End If
        End If

End Sub

Public Sub GetMemberBottomCenterProperties(elems As IJElements, _
                    ByRef x As Double, ByRef y As Double, ByRef z As Double, ByRef dMembAng As Double, _
                    ByRef sSecName As String, ByRef sRefStd As String, index As Long)
Const METHOD = "GetMemberBottomCenterProperties"
On Error GoTo ErrorHandler


    Dim oMembSys As ISPSMemberSystem
    Dim oMembPart As ISPSMemberPartPrismatic
    Dim oPnt As IJPoint
    
    sSecName = vbNullString
    sRefStd = vbNullString
    
    If TypeOf elems.Item(index) Is ISPSMemberSystem Then
        Set oMembSys = elems.Item(index)
        If Not oMembSys Is Nothing Then
            GetFootingPositionFromMember oMembSys, x, y, z, dMembAng
            Set oMembPart = GetBottomMemberPart(oMembSys)
            sRefStd = oMembPart.CrossSection.SectionStandard
            sSecName = oMembPart.CrossSection.SectionName
        End If
        Set oMembSys = Nothing
        Set oMembPart = Nothing
        
    ElseIf TypeOf elems.Item(index) Is IJPoint Then
        Set oPnt = elems.Item(index)
        oPnt.GetPoint x, y, z
    End If
    
    Exit Sub
    
ErrorHandler:  HandleError MODULE, METHOD
End Sub


'This subroutine return appropirate pier width/length components to be used for deriving slab size
' This will be based on whether indiviual column pier is governing range box or not . If yes then get appropriate component
' based pier orienation angle and its width/length
Public Sub GetPierWidthLengthComponents(pPropertyDescriptions As IJDPropertyDescription, globalPierColl As Collection, _
                                             xmin As Double, xmax As Double, _
                                           ymin As Double, ymax As Double, dblSlabRotation As Double, _
                                           ByRef pierwidthComp As Double, ByRef pierLengthComp As Double)
                                           
Const METHOD = "GetPierWidthLengthComponents"
On Error GoTo ErrorHandler

Dim i As Long
Dim x As Double, y As Double, z As Double
Dim MemberObj As IJDMemberObjects
Dim oSmartOcc As IJSmartOccurrence
Dim PierRotationAngle As Double
Dim PierOrientation As Long
Dim dMembAng As Double
Dim sSecName As String, sRefStd As String
Dim tmpwidthComp As Double, tmpLengthComp As Double
Dim tmpwidthComp1 As Double, tmpLengthComp1 As Double
Dim oPierAttribs As IJDAttributes
Dim pierColl As Collection
Dim kk As Long
Dim mergedPierHandled As Boolean
Dim supported As IJElements
     
    Set MemberObj = pPropertyDescriptions.CAO
   
    For i = 1 To globalPierColl.count
        
        Set supported = New JObjectCollection
        Set pierColl = globalPierColl.Item(i)
        mergedPierHandled = False
        
        For kk = 1 To pierColl.count
            supported.Add pierColl.Item(kk)
        Next kk
        
        For kk = 1 To pierColl.count
            GetMemberBottomCenterProperties supported, x, y, z, dMembAng, sSecName, sRefStd, kk
            RotatePositionAboutZ x, y, z, dblSlabRotation
            If mergedPierHandled = False Then
                If Abs(x - xmin) < TOL Or Abs(x - xmax) < TOL Or Abs(y - ymin) < TOL Or Abs(y - ymax) < TOL Then
                    Set oSmartOcc = MemberObj.ItemByDispid(4, i)
                    Set oPierAttribs = oSmartOcc
                    tmpLengthComp = oPierAttribs.CollectionOfAttributes(FTG_PIER_DIM).Item(PIER_LENGTH).Value
                    tmpwidthComp = oPierAttribs.CollectionOfAttributes(FTG_PIER_DIM).Item(PIER_WIDTH).Value
                    
                    PierRotationAngle = oPierAttribs.CollectionOfAttributes(FTG_PIER).Item(PIER_ROTATION_ANGLE).Value
                    PierOrientation = oPierAttribs.CollectionOfAttributes(FTG_PIER).Item(PIER_ORIENTATION).Value
                    CheckForUndefinedValueAndRaiseError pPropertyDescriptions.CAO, PierOrientation, STRUCT_COORD_SYS_REF, 126
                    
                    If pierColl.count > 1 Then ' for merged pier take appropriate pier size
                        Dim PierXmin As Double, PierYmin As Double, PierXmax As Double, PierYmax As Double, PierXMid As Double, PierYMid As Double
                        
                        Call GetColumnsRangeAboutRotatedZ(supported, PierRotationAngle, PierXmin, PierYmin, PierXmax, PierYmax, z, PierXMid, PierYMid)
                        tmpLengthComp = tmpLengthComp - (PierXmax - PierXmin)
                        tmpwidthComp = tmpwidthComp - (PierYmax - PierYmin)
                    End If
           
                    Dim PierShape As Long 'TR#66868
                    PierShape = oPierAttribs.CollectionOfAttributes(FTG_PIER).Item(PIER_SHAPE).Value
                    CheckForUndefinedValueAndRaiseError pPropertyDescriptions.CAO, PierShape, PRISMATIC_FOOTING_SHAPES, 124
                    
                    
                    If PierOrientation = 2 Then '2 is Local to the SPSMemberPart
                        PierRotationAngle = PierRotationAngle + dMembAng
                    End If
                    If PierShape = 3 Then PierRotationAngle = 0# ' if pier is circular then no need to consider its rotation.
                    
                    PierRotationAngle = Abs(Abs(PierRotationAngle) - dblSlabRotation)
                    If PierRotationAngle > PI / 2 Then PierRotationAngle = Abs(PI / 2 - PierRotationAngle)
                    
                    If Abs(x - xmin) < TOL Or Abs(x - xmax) < TOL Then ' This column is governing length of range box
                        tmpLengthComp1 = tmpLengthComp * Cos(PierRotationAngle) + tmpwidthComp * Sin(PierRotationAngle)
                    End If
                    
                    If Abs(y - ymin) < TOL Or Abs(y - ymax) < TOL Then ' This columns is governing width of range box
                       tmpwidthComp1 = tmpwidthComp * Cos(PierRotationAngle) + tmpLengthComp * Sin(PierRotationAngle)
                    End If
                    
                    mergedPierHandled = True
                End If
            End If
            
            If pierwidthComp < tmpwidthComp1 Then
               pierwidthComp = tmpwidthComp1
            End If
            If pierLengthComp < tmpLengthComp1 Then
               pierLengthComp = tmpLengthComp1
            End If
            
            Set oSmartOcc = Nothing
            Set oPierAttribs = Nothing
        Next kk
        Set supported = Nothing
    Next i

Set MemberObj = Nothing
Set pierColl = Nothing

Exit Sub
ErrorHandler:
If Err.Description = "Undefined Value" Then
    Err.Raise E_FAIL
Else
    HandleError MODULE, METHOD
End If
End Sub

' This function is used for setting filter criteria for selection of inputs for footings
'This function is differnet from existing single footing function. Here additional check is
' that all selected members should have same BOS level

Public Function CommonCombinedCheckMember(obj As Object, elems As IJElements) As Integer
Const METHOD = "CommonCombinedCheckMember"
On Error GoTo ErrorHandler
    
    CommonCombinedCheckMember = 0
    Dim oMembSys As ISPSMemberSystem

    On Error Resume Next
    Set oMembSys = obj
    On Error GoTo ErrorHandler
    If oMembSys Is Nothing Then
        Exit Function
    Else
        If IsBuiltUPMemeberSystem(oMembSys) Then
            Exit Function
        End If
        
        If elems.count < 1 Then
            CommonCombinedCheckMember = 1
            Set oMembSys = Nothing
            Exit Function
        End If
        
        If Not TypeOf elems.Item(1) Is ISPSMemberSystem Then ' not a member system must be point
            CommonCombinedCheckMember = 1
            Set oMembSys = Nothing
            Exit Function
        End If
        
        Dim oMembPart As ISPSMemberPartPrismatic
        Dim xstart As Double, ystart As Double, zstart As Double
        Dim xend As Double, yend As Double, zend As Double
        Dim zBotCommon As Double, zBotNew As Double
        
        Set oMembPart = GetBottomMemberPart(oMembSys)
        oMembPart.Axis.EndPoints xstart, ystart, zstart, xend, yend, zend
        If zstart < zend Then
           zBotCommon = zstart
        Else
           zBotCommon = zend
        End If
        
        Set oMembSys = Nothing
        Set oMembPart = Nothing

        Set oMembSys = elems.Item(1)
        Set oMembPart = GetBottomMemberPart(oMembSys)
        oMembPart.Axis.EndPoints xstart, ystart, zstart, xend, yend, zend
        If zstart < zend Then
           zBotNew = zstart
        Else
           zBotNew = zend
        End If
        
        If Abs(zBotNew - zBotCommon) < TOL Then
            CommonCombinedCheckMember = 1
        End If
        
        Set oMembSys = Nothing
        Set oMembPart = Nothing
        
        Exit Function
    End If
    
    

Exit Function
ErrorHandler:  HandleError MODULE, METHOD
End Function

Public Function ValidateInputMembersCombFtg(pMemberDescription As IJDMemberDescription, bosFlag As Boolean)
Const METHOD = "ValidateInputMembersCombFtg"
On Error GoTo ErrorHandler
    
    Dim oReferencesCollection  As IMSSymbolEntities.IJDReferencesCollection
    Dim oRefColl1  As IMSSymbolEntities.IJDReferencesCollection
    Dim oMembSys As ISPSMemberSystem
    Dim oMembPart As ISPSMemberPartPrismatic
    Dim x As Double, y As Double, z As Double
    Dim x1 As Double, y1 As Double, z1 As Double
    Dim bz As Double
    Dim cnt As Integer
    Dim oPlane As IJPlane
    Dim i As Integer
    Dim supported As IJElements

    ValidateInputMembersCombFtg = vbNullString
    
    Set oReferencesCollection = GetRefCollection(pMemberDescription.CAO)
    Set oRefColl1 = oReferencesCollection.IJDEditJDArgument.GetEntityByIndex(1)
    Set oReferencesCollection = Nothing
    Set supported = GetUpdatedRefColl(oRefColl1)
    Set oRefColl1 = Nothing
    Set oLocalizer = New IMSLocalizer.Localizer
    oLocalizer.Initialize App.Path & "\" & FOOTING_MACROS
    
    cnt = supported.count

    If cnt < 1 Then
        ValidateInputMembersCombFtg = oLocalizer.GetString(IDS_COMBINEDFTG_REQDMIN_TWOMEMBS, "Combined footings require a minimum of two supported members.")
        Set supported = Nothing
        Exit Function
    End If

    If bosFlag = True Then
    
        If TypeOf supported.Item(1) Is ISPSMemberSystem Then
            Set oMembSys = supported.Item(1)
            Set oMembPart = GetBottomMemberPart(oMembSys)
            oMembPart.Axis.EndPoints x, y, z, x1, y1, z1
            Set oMembSys = Nothing
            Set oMembPart = Nothing
        
            If z < z1 Then bz = z Else bz = z1
          
            For i = 2 To cnt
                Set oMembSys = supported.Item(i)
                Set oMembPart = GetBottomMemberPart(oMembSys)
                oMembPart.Axis.EndPoints x, y, z, x1, y1, z1
                Set oMembSys = Nothing
                Set oMembPart = Nothing
                If z > z1 Then z = z1
                If Abs(bz - z) > TOL Then
                  Set supported = Nothing
                  ValidateInputMembersCombFtg = oLocalizer.GetString(IDS_BOSLEVEL_SUPPORTEDMEMBS_TOBESAME, "Combined footings require all supported members to have the same bottom-of-steel elevation.")
                  Exit For
                End If
            Next i
            
        Else    ' placed by points free in space
            ValidateInputMembersCombFtg = vbNullString
        End If
        
    End If

    Set supported = Nothing
    Set oLocalizer = Nothing
    
    Exit Function
    
ErrorHandler:  HandleError MODULE, METHOD
End Function
 
Public Sub CombineFootings_SetInputs(ByVal FtgObject As Object, ByVal FtgDefinitionObject As Object, ByVal supported As SPSFootings.IJElements, ByVal supporting As SPSFootings.IJElements)
Const METHOD = "CombineFootings_SetInputs"
On Error GoTo ErrorHandler
    
    Dim strUserType As String
    Dim strSCName As String
    Dim oSmartItem As IJSmartItem
    Dim oSmartClass As IJSmartClass
    Dim oUserType As IJDUserType
    Dim oSmartOcc As IJSmartOccurrence
    Dim oFtgFactory As SPSFootingFactory
    Dim ijdObject As ijdObject
    Dim NewSO As Boolean
    NewSO = False
    Dim i As Integer, j As Integer
    
    ' Create the reference collection
    Dim oSymbolEntitiesFactory As New IMSSymbolEntities.DSymbolEntitiesFactory
    Dim oReferencesCollection As IMSSymbolEntities.IJDReferencesCollection
    Dim oReferencesCollection1 As IMSSymbolEntities.IJDReferencesCollection
    Dim oReferencesCollection2 As IMSSymbolEntities.IJDReferencesCollection
    
    Set oSmartOcc = FtgObject
    
    Set oReferencesCollection = GetRefCollection(oSmartOcc)
    If Not oReferencesCollection Is Nothing Then
        Set oReferencesCollection1 = oReferencesCollection.IJDEditJDArgument.GetEntityByIndex(1)
        Set oReferencesCollection2 = oReferencesCollection.IJDEditJDArgument.GetEntityByIndex(2)
        If Not oReferencesCollection1 Is Nothing Then
            For j = 1 To oReferencesCollection1.IJDEditJDArgument.GetCount
                Dim oObject As ijdObject
                If TypeOf oReferencesCollection1.IJDEditJDArgument.GetEntityByIndex(j) Is IJPoint Then
                    Set oObject = oReferencesCollection1.IJDEditJDArgument.GetEntityByIndex(j)
                    oObject.Remove
                End If
             Next j
            oReferencesCollection1.IJDEditJDArgument.RemoveAll
        Else
            Set oReferencesCollection1 = oSymbolEntitiesFactory.CreateEntity(referencesCollection, _
                                                     GetResourceMgr())
        End If
    
        If Not oReferencesCollection2 Is Nothing Then
            oReferencesCollection2.IJDEditJDArgument.RemoveAll
        Else
            If supporting.count >= 1 Then
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
        If supporting.count >= 1 Then
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
    On Error GoTo ErrorHandler
    
    If Not OldItem Is Nothing Then
        strOldItemName = OldItem.name
    End If
    
    Set oSmartItem = FtgDefinitionObject
    If strOldItemName <> oSmartItem.name Then
        Set oSmartClass = oSmartItem.Parent
        strUserType = oSmartClass.SOUserType
        Set oUserType = oSmartOcc
        oUserType.UserType = strUserType
        strSCName = oSmartClass.SCName
        oSmartOcc.RootSelectorClass = strSCName
        oSmartOcc.RootSelection = oSmartItem.name
    End If
        
    For i = 1 To supported.count
        If TypeOf supported.Item(i) Is ISPSMemberSystem Then
            oReferencesCollection1.IJDEditJDArgument.SetEntity i, supported.Item(i), ISPSMemberSystemSuppingNotify1, "MembSysSuppingNotify1RC_DEST"
        Else
            oReferencesCollection1.IJDEditJDArgument.SetEntity i, supported.Item(i), IJFullObject, "SPSRCtoRC_DEST"
        End If
    Next i
    
    If supporting.count >= 1 Then
        oReferencesCollection2.IJDEditJDArgument.SetEntity 1, supporting.Item(1), IJPlane, "SPSSuppPlaneToRC_DEST"
    End If
    
    oReferencesCollection.IJDEditJDArgument.SetEntity 1, oReferencesCollection1, IJFullObject, "SPSRCToRC_1_DEST"
    If supporting.count >= 1 Then
        oReferencesCollection.IJDEditJDArgument.SetEntity 2, oReferencesCollection2, IJFullObject, "SPSRCToRC_2_DEST"
    Else
        If Not oReferencesCollection2 Is Nothing Then
            Set ijdObject = oReferencesCollection2
            ijdObject.Remove
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
    
Exit Sub
ErrorHandler:  HandleError MODULE, METHOD
End Sub

Public Sub CombineFootings_GetInputs(ByVal FtgObject As Object, ByVal supported As SPSFootings.IJElements, ByVal supporting As SPSFootings.IJElements)
Const METHOD = "CombineFootings_GetInputs"
On Error GoTo ErrorHandler

    Dim oSmartOcc As IJSmartOccurrence
    Set oSmartOcc = FtgObject
    Dim oRefColl As IJDReferencesCollection
    Set oRefColl = GetRefCollection(oSmartOcc)

    Dim oRefColl1 As IJDReferencesCollection
    Dim oRefColl2 As IJDReferencesCollection
    Dim SupportedElms As IJElements
    Dim supportingElms As IJElements
    
    Set oRefColl1 = oRefColl.IJDEditJDArgument.GetEntityByIndex(1)

    Set oRefColl2 = oRefColl.IJDEditJDArgument.GetEntityByIndex(2)
    Set supportingElms = GetUpdatedRefColl(oRefColl2)
    Set SupportedElms = GetUpdatedRefColl(oRefColl1)
        
    Set oRefColl = Nothing
    Set oRefColl1 = Nothing
    Set oRefColl2 = Nothing
    Set oSmartOcc = Nothing
    
    Dim i As Integer

    If supportingElms.count > 0 Then
        supporting.Add supportingElms.Item(1)
    End If
    
    For i = 1 To SupportedElms.count
        supported.Add SupportedElms.Item(i)
    Next i

    Set supportingElms = Nothing
    Set SupportedElms = Nothing
    
 Exit Sub
ErrorHandler:  HandleError MODULE, METHOD

End Sub

Public Sub EvaluateGrout(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
Const METHOD = "EvaluateGrout"
On Error GoTo ErrorHandler
    
     Dim oReferencesCollection0fCAO  As IMSSymbolEntities.IJDReferencesCollection
     Dim oRefColl  As IMSSymbolEntities.IJDReferencesCollection
     Dim strSecname As String, strRefStd As String
     Dim x As Double, y As Double, z As Double
     Dim RotationInherited As Double
     Dim groutIndex As Long
     Dim supported As IJElements
     
     Set oRefColl = GetRefCollection(pPropertyDescriptions.CAO)
     Set oReferencesCollection0fCAO = oRefColl.IJDEditJDArgument.GetEntityByIndex(1)
     Set oRefColl = Nothing
     Set supported = GetUpdatedRefColl(oReferencesCollection0fCAO)
     groutIndex = pPropertyDescriptions.index
     strSecname = vbNullString
     strRefStd = vbNullString
     RotationInherited = 0#
     GetMemberBottomCenterProperties supported, x, y, z, RotationInherited, strSecname, strRefStd, groutIndex
    
     If TypeOf supported.Item(groutIndex) Is IJPoint Then
        RotationInherited = GetPlanAngleFromSymbolOccurrence(pPropertyDescriptions.CAO)
     End If
     
     Dim pSymbol As IJDSymbol
     Set pSymbol = pObject ' use this object instead pPropertyDescriptions.Object?
     Dim pOcc As IJDOccurrence

    Set pOcc = pSymbol
    If pOcc Is Nothing Then
         Exit Sub
    End If

     
     Dim Matrix As IJDT4x4
     Set Matrix = New DT4x4
     
     Matrix.LoadIdentity
     
     Matrix.IndexValue(12) = x
     Matrix.IndexValue(13) = y
     Matrix.IndexValue(14) = z
     
     
     Dim Vec As DVector
     Set Vec = New DVector
     Vec.Set 0, 0, 1
     Dim GroutRotationAngle As Double
     Dim GroutOrientation As Long
     Dim GroutSizingRule As Long
     Dim oGroutAttribs As IJDAttributes
     Dim SmartOcc As IJSmartOccurrence
     Dim MemberObj As IJDMemberObjects
     Set MemberObj = pPropertyDescriptions.CAO
     Set SmartOcc = MemberObj.ItemByDispid(1, groutIndex)
     Set oGroutAttribs = SmartOcc 'pObject
     
     GroutRotationAngle = oGroutAttribs.CollectionOfAttributes(FTG_GROUT_PAD).Item(GROUT_ROTATION_ANGLE).Value
     GroutOrientation = oGroutAttribs.CollectionOfAttributes(FTG_GROUT_PAD).Item(GROUT_ORIENTATION).Value
     GroutSizingRule = oGroutAttribs.CollectionOfAttributes(FTG_GROUT_PAD).Item(GROUT_SIZE_RULE).Value
     
     CheckForUndefinedValueAndRaiseError pPropertyDescriptions.CAO, GroutSizingRule, FOOTING_COMP_SIZE_RULE, 122
     CheckForUndefinedValueAndRaiseError pPropertyDescriptions.CAO, GroutOrientation, STRUCT_COORD_SYS_REF, 123
     
     If GroutOrientation = 1 Then '1 is Global
        Matrix.Rotate GroutRotationAngle, Vec
     ElseIf GroutOrientation = 2 Then '2 is Local to the SPSMemberPart
        GroutRotationAngle = GroutRotationAngle + RotationInherited
        Matrix.Rotate GroutRotationAngle, Vec
     End If
            
     If strSecname = vbNullString Then ' no cross section name, placed by points
         pOcc.Matrix = Matrix
         Exit Sub
     End If
            
     If GroutSizingRule = 1 Or GroutSizingRule = 2 Then
        Dim Depth As Double
        Dim SecWidth As Double
        Dim GroutEdgeClearance As Double

        If strSecname <> vbNullString Then
            Depth = GetCSAttribData(strSecname, strRefStd, CROSSSSECTION_DIMENSIONS, CROSSSECTION_DEPTH)
            SecWidth = GetCSAttribData(strSecname, strRefStd, CROSSSSECTION_DIMENSIONS, CROSSSECTION_WIDTH)
        End If
        
        GroutEdgeClearance = oGroutAttribs.CollectionOfAttributes(FTG_GROUT_PAD).Item(GROUT_EDGE_CLEARANCE).Value

        Dim GroutShape As Long
        GroutShape = oGroutAttribs.CollectionOfAttributes(FTG_GROUT_PAD).Item(GROUT_SHAPE).Value
        CheckForUndefinedValueAndRaiseError pPropertyDescriptions.CAO, GroutShape, PRISMATIC_FOOTING_SHAPES, 121
        
        Dim tmpAng As Double
        If GroutShape = 3 Then GroutRotationAngle = 0#
        tmpAng = Abs(GroutRotationAngle - RotationInherited)
        If tmpAng > PI / 2 Then tmpAng = Abs(PI - tmpAng)
        
        Dim secDepthComp As Double
        Dim SecWidthComp As Double
        
        secDepthComp = Depth * Cos(tmpAng) + SecWidth * Sin(tmpAng)
        SecWidthComp = SecWidth * Cos(tmpAng) + Depth * Sin(tmpAng)
        If GroutShape = 2 Then
            oGroutAttribs.CollectionOfAttributes(FTG_GROUT_PAD_DIM).Item(GROUT_WIDTH).Value = _
                            SecWidthComp + (GroutEdgeClearance * 2)
            oGroutAttribs.CollectionOfAttributes(FTG_GROUT_PAD_DIM).Item(GROUT_LENGTH).Value = _
                                    secDepthComp + (GroutEdgeClearance * 2)
        ElseIf GroutShape = 3 Then
        
            'TR#72794- if grout is circular then take diagonal + clearance=diameter. i.e. width & length
            oGroutAttribs.CollectionOfAttributes(FTG_GROUT_PAD_DIM).Item(GROUT_WIDTH).Value = Sqr(SecWidthComp * SecWidthComp + secDepthComp * secDepthComp) + (GroutEdgeClearance * 2)
            oGroutAttribs.CollectionOfAttributes(FTG_GROUT_PAD_DIM).Item(GROUT_LENGTH).Value = Sqr(SecWidthComp * SecWidthComp + secDepthComp * secDepthComp) + (GroutEdgeClearance * 2)
        End If
     End If

     pOcc.Matrix = Matrix
     
    
    Exit Sub
    
ErrorHandler:  HandleError MODULE, METHOD
End Sub

Public Function GetGroutHt(pPropertyDescriptions As IJDPropertyDescription) As Double
Const METHOD = "GetGroutHt"
On Error GoTo ErrorHandler

Dim MemberObj As IJDMemberObjects
Dim oAttr As IJDAttributes
Dim bIsNeeded As Boolean
Dim oSmartOcc As IJSmartOccurrence
Dim oGroutAttribs As IJDAttributes

     Set MemberObj = pPropertyDescriptions.CAO
     Set oAttr = MemberObj
     On Error Resume Next
     bIsNeeded = oAttr.CollectionOfAttributes(PIER_FOOTING_ASM).Item(WITH_GROUT_PAD).Value
     If Err.Number <> 0 Then
        Err.Clear
        bIsNeeded = oAttr.CollectionOfAttributes(SLAB_FOOTING_ASM).Item(WITH_GROUT_PAD).Value
     End If
     
     If Err.Number <> 0 Then
        Err.Clear
        bIsNeeded = oAttr.CollectionOfAttributes(PIER_SLAB_FOOTING_ASM).Item(WITH_GROUT_PAD).Value
     End If
     GetGroutHt = 0#
     If bIsNeeded Then 'check for existence of the GroutPad
        Set oSmartOcc = MemberObj.ItemByDispid(1, 1) ' get first grout height. It will be same for all in case of mergedpier & combined slab asm
        Set oGroutAttribs = oSmartOcc
        GetGroutHt = oGroutAttribs.CollectionOfAttributes(FTG_GROUT_PAD_DIM).Item(GROUT_HEIGHT).Value
     End If
     
Set MemberObj = Nothing
Set oGroutAttribs = Nothing
Set oAttr = Nothing
Set oSmartOcc = Nothing

Exit Function
ErrorHandler:  HandleError MODULE, METHOD
End Function
Public Function GetUpdatedRefColl(ByVal oRefColl As IJDReferencesCollection) As IJElements
Const METHOD = "GetUpdatedRefColl"
On Error GoTo ErrorHandler

Dim pEnumJDArgument As IEnumJDArgument
Dim arg1 As IJDArgument
Dim found As Long
Dim i As Long

    Set GetUpdatedRefColl = New JObjectCollection
    Set pEnumJDArgument = oRefColl
    If Not oRefColl Is Nothing Then

        If Not pEnumJDArgument Is Nothing Then
            pEnumJDArgument.Reset
            i = 1
            Do
               pEnumJDArgument.Next 1, arg1, found
               If found <> 0 Then
                    GetUpdatedRefColl.Add arg1.Entity
                    Set arg1 = Nothing
                    i = i + 1
               Else: Exit Do
               End If
            Loop
        End If
        
        Set pEnumJDArgument = Nothing
    End If
    
Exit Function
ErrorHandler:  HandleError MODULE, METHOD
End Function

Public Function GetGroutCount(ByVal pMemberDescription As IJDMemberDescription) As Long
Const METHOD = "GetGroutCount"
On Error GoTo ErrorHandler

Dim oReferencesCollection0fCAO  As IMSSymbolEntities.IJDReferencesCollection
Dim oRefColl  As IMSSymbolEntities.IJDReferencesCollection
Dim supported As IJElements

     Set oRefColl = GetRefCollection(pMemberDescription.CAO)
     Set oReferencesCollection0fCAO = oRefColl.IJDEditJDArgument.GetEntityByIndex(1)
     
     Set supported = GetUpdatedRefColl(oReferencesCollection0fCAO)
     
    GetGroutCount = supported.count

Set oRefColl = Nothing
Set oReferencesCollection0fCAO = Nothing
Set supported = Nothing
    
Exit Function
ErrorHandler:  HandleError MODULE, METHOD
End Function

Public Sub CreateContolPointRelation(pMemberDesc As IJDMemberDescription)
Const METHOD = "CreateContolPointRelation"
On Error GoTo ErrorHandler

Dim oParent As ijdObject
Dim oChild As ijdObject
Dim lConditionId As Long
Dim CollHlpr As IMSRelation.DCollectionHelper
Dim TOColl As IJDTargetObjectCol
Dim AssocRel As IJDAssocRelation
'Dim IIDCtrlPt As Variant
Dim oNamedItem As IJNamedItem
Dim oRevision As IJRevision
Dim oRelationship As IJDRelationship
    
    lConditionId = -1
    Set oParent = pMemberDesc.CAO
    Set oChild = pMemberDesc.object
    If Not oParent Is Nothing Then
        lConditionId = oParent.PermissionGroup
        If Not lConditionId = -1 Then
            oChild.PermissionGroup = lConditionId
        End If
    End If
    Set oParent = Nothing
    Set oChild = Nothing
    

    Set AssocRel = pMemberDesc.object
    Set CollHlpr = AssocRel.CollectionRelations(IJControlPoint, "Parent")
    
    If Not CollHlpr Is Nothing Then
        Set TOColl = CollHlpr
        On Error Resume Next
        If TOColl.count > 0 Then
            TOColl.Remove 1
        End If
        On Error GoTo ErrorHandler
        TOColl.Add pMemberDesc.CAO, vbNullString, oRelationship
        Set oRevision = New JRevision
        oRevision.AddRelationShip oRelationship
        Set CollHlpr = Nothing
        Set TOColl = Nothing
        Set AssocRel = Nothing
    End If

Exit Sub
ErrorHandler:  HandleError MODULE, METHOD
End Sub

Public Function GetBottomMemberPart(oMembSys As ISPSMemberSystem) As ISPSMemberPartPrismatic
Const METHOD = "GetBottomMemberPart"
On Error GoTo ErrorHandler
        
        Dim x As Double, y As Double, z As Double
        Dim x1 As Double, y1 As Double, z1 As Double
        Dim oEles As IJElements
        Set oEles = oMembSys.DesignParts
        
        oMembSys.PhysicalAxis.EndPoints x, y, z, x1, y1, z1
        If z < z1 Then
            Set GetBottomMemberPart = oMembSys.DesignParts.Item(1)
        Else
            Set GetBottomMemberPart = oMembSys.DesignParts.Item(oEles.count)
        End If
        
        
Exit Function
ErrorHandler:      HandleError MODULE, METHOD
End Function
'TR#71850- Last argument added. This contains object ins selectset. This is used to check that
' plane doesn't belong to same footing which is being modified.
Public Function CommonCombinedCheckBottomPlane(obj As Object, elems As IJElements, ObjsinSelectSet As IJElements) As Integer
Const METHOD = "CommonCombinedCheckBottomPlane"
On Error GoTo ErrorHandler

    CommonCombinedCheckBottomPlane = 0
    Dim oMembSys As ISPSMemberSystem
    Dim oCurve As IJCurve
    Dim oPlane As IJPlane
    On Error Resume Next
    Set oPlane = obj
    If oPlane Is Nothing Then
    Exit Function
    End If
    On Error GoTo ErrorHandler
    
        'TR#71850
     If ObjsinSelectSet Is Nothing Then Exit Function
     
     If ObjsinSelectSet.count > 0 Then
        If TypeOf ObjsinSelectSet.Item(1) Is ISPSFooting Then
            Dim FootingObj As ISPSFooting
            Set FootingObj = ObjsinSelectSet.Item(1)
            
            If TypeOf obj Is IJDReferenceProxy Then
                Dim oRefProxy As IJDReferenceProxy
                Dim footingComp As ISPSFootingComponent
                Dim systemChild As IJDesignChild
                Dim desParent As IJDesignParent
    
                Set oRefProxy = obj
                If TypeOf oRefProxy.Reference Is ISPSFootingComponent Then
                    Set footingComp = oRefProxy.Reference
                    Set systemChild = footingComp
                    Set desParent = systemChild.GetParent
                ElseIf TypeOf oRefProxy.Reference Is ISPSFooting Then
                    Set desParent = oRefProxy.Reference
                End If
                
                If desParent Is FootingObj Then
                    Set oRefProxy = Nothing
                    Set FootingObj = Nothing
                    Set footingComp = Nothing
                    Set systemChild = Nothing
                    Set desParent = Nothing
                    CommonCombinedCheckBottomPlane = 0
                    Exit Function
                End If
      
            End If
            
            Set oRefProxy = Nothing
            Set FootingObj = Nothing
            Set footingComp = Nothing
            Set systemChild = Nothing
            Set desParent = Nothing
        End If
    End If
    'TR#71850
    
    Dim i As Integer, NumInter As Long, Numover As Long
    Dim pts() As Double
    Dim Stx As Double, Sty As Double, Stz As Double, Endx As Double, Endy As Double, Endz As Double
    Dim Srcsx As Double, Srcy As Double, Srcz As Double, inx As Double, iny As Double, inz As Double
    Dim Rtx As Double, Rty As Double, Rtz As Double
    Dim Normx As Double, Normy As Double, Normz As Double
    
    Dim oline As Line3d
    Dim NewCurve As IJCurve
    Dim DummyFace As New Plane3d
    Dim GeomFactory As New IngrGeom3D.GeometryFactory
    Dim code As Geom3dIntersectConstants
    
    oPlane.GetRootPoint Rtx, Rty, Rtz
    oPlane.GetNormal Normx, Normy, Normz

    If Abs(Normx) > 0.01 Or Abs(Normy) > 0.01 Then
      CommonCombinedCheckBottomPlane = 0
      Exit Function
    End If
    Set DummyFace = GeomFactory.Planes3d.CreateByPointNormal(Nothing, Rtx, Rty, Rtz, Normx, Normy, Normz)
        
    ' get/check the lowest points when points have been picked
    Dim oPoint As IJDPosition
    If elems.count > 0 Then
        If TypeOf elems.Item(1) Is IJDPosition Then
            Dim tempX As Double, tempY As Double, tempZ As Double
            Set oPoint = elems.Item(1)
            Stz = oPoint.z
            For i = 2 To elems.count
                If TypeOf elems.Item(i) Is IJDPosition Then
                    Set oPoint = elems.Item(i)
                    tempZ = oPoint.z
                End If
                If tempZ < Stz Then
                    Stz = tempZ
                End If
            Next i
            
            If (Rtz < Stz) Then
                CommonCombinedCheckBottomPlane = 1
            Else
                CommonCombinedCheckBottomPlane = 0
            End If
            Exit Function
        End If
    End If
    
    For i = 1 To elems.count
        Set oMembSys = elems.Item(i)
        If Not oMembSys Is Nothing Then
            
            Set oCurve = oMembSys
            oCurve.EndPoints Stx, Sty, Stz, Endx, Endy, Endz
            
            If (Endz > Stz And Rtz >= Stz) Or (Stz > Endz And Rtz >= Endz) Then
                CommonCombinedCheckBottomPlane = 0
                Exit For
            End If
            Set oline = GeomFactory.Lines3d.CreateByPtVectLength(Nothing, Stx, Sty, Stz, _
                                Endx - Stx, Endy - Sty, Endz - Stz, 1)
            oline.Infinite = True
            Set NewCurve = oline
            NewCurve.Intersect DummyFace, NumInter, pts, Numover, code
            
            If NumInter >= 1 Then
                Dim NewPos As New DPosition
                Dim Vec1 As New DVector
                Dim Vec2 As New DVector
                NewPos.Set pts(0), pts(1), pts(2)
                Vec1.Set NewPos.x - Stx, NewPos.y - Sty, NewPos.z - Stz
                Vec2.Set NewPos.x - Endx, NewPos.y - Endy, NewPos.z - Endz
                If Vec1.Dot(Vec2) > 0# Then
                    If Normz >= 0.7071 And Normz <= 1 Then
                        CommonCombinedCheckBottomPlane = 1
                    ElseIf Normz <= -0.7071 And Normz >= -1 Then
                        CommonCombinedCheckBottomPlane = 1
                    End If
                Else
                    CommonCombinedCheckBottomPlane = 0
                End If
'            ElseIf Rtz < Endz And Rtz < Stz Then ' TR#71999- commented this. if plane doesn't intersect member line then that plane can not be valid plane.
'                CommonCombinedCheckBottomPlane = 1
            Else
                CommonCombinedCheckBottomPlane = 0
                Exit Function
            End If
        End If
    Next i
    
Exit Function
ErrorHandler:  HandleError MODULE, METHOD
End Function


Public Sub GenerateNameForFooting(obj As Object)
Const METHOD = "GenerateNameForFooting"
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
    Call oNameRuleHlpr.GetEntityNamingRulesGivenProgID(FOOTINGCOMPPROGID, NamingRules)
    Dim ncount As Integer
    Dim oNameRuleAE As GSCADGenNameRuleAE.IJNameRuleAE
    
        For ncount = 1 To NamingRules.count
            Set oNameRuleHolder = NamingRules.Item(1)
        Next ncount
    
    Call oNameRuleHlpr.AddNamingRelations(obj, oNameRuleHolder, oNameRuleAE)
    Set oNameRuleHolder = Nothing
    
    Set oActiveNRHolder = Nothing
    Set oNameRuleHolder = Nothing
    Set oNameRuleAE = Nothing
    Set oNameRuleHlpr = Nothing
Exit Sub

ErrorHandler:     HandleError MODULE, METHOD
End Sub
Public Sub AddSystemAndNameRule(pMemberDescription As IJDMemberDescription)
Const METHOD = "AddSystemAndNameRule"
On Error GoTo ErrorHandler
     Dim oDesignParent As IJDesignParent
     Set oDesignParent = pMemberDescription.CAO
     oDesignParent.AddChild pMemberDescription.object
          
     Dim oAssemblyparent As IJAssembly
     Set oAssemblyparent = oDesignParent
     oAssemblyparent.AddChild pMemberDescription.object
     Set oAssemblyparent = Nothing
     
     Call GenerateNameForFooting(pMemberDescription.object)
     Set oDesignParent = Nothing
Exit Sub
Errx:
ErrorHandler:      HandleError MODULE, METHOD
End Sub
Public Sub SetControlFlags(controlFlag As Long, m_oControlPoint As IJControlPoint)
Dim oControlFlags As IJControlFlags

     Set oControlFlags = m_oControlPoint
     oControlFlags.ControlFlags(CTL_FLAG_SYSTEM_MASK) = controlFlag
     
     Set oControlFlags = Nothing
End Sub

' This function will create control points with default inuts. Its properties like type, subtype, diameter
' position, control flags can be set later by different symbols

Public Function CreateControlPoint(ByVal pResourceManager As IUnknown, IsSymplePhysical As Boolean) As IJControlPoint
Const METHOD = "CreateControlPoint"
On Error GoTo ErrorHandler

    Dim m_oGBSFactory  As IJGeneralBusinessObjectsFactory
    Dim m_oControlPoint As IJControlPoint

    Set m_oGBSFactory = New GeneralBusinessObjectsFactory
    Set m_oControlPoint = m_oGBSFactory.CreateControlPoint(pResourceManager, 0#, 0#, 0#, 0.001, , , IsSymplePhysical)
    Set CreateControlPoint = m_oControlPoint
     
    Set m_oGBSFactory = Nothing
    Set m_oControlPoint = Nothing

Exit Function

ErrorHandler:      HandleError MODULE, METHOD
End Function

Public Sub SetOccurrenceMatrix(pObj As Object)
Const METHOD = "SetOccurrenceMatrix"
On Error GoTo Errx

    Dim oMembSys As ISPSMemberSystem
    Dim oMembObj As Object
    Dim oMembPart As ISPSMemberPartPrismatic
    Dim oReferencesCollection0fCAO  As IMSSymbolEntities.IJDReferencesCollection
    Dim oRefColl  As IMSSymbolEntities.IJDReferencesCollection
    Dim CP As Long
    Dim oFooting As ISPSFooting
    
    Set oReferencesCollection0fCAO = GetRefCollection(pObj)
    
    If oReferencesCollection0fCAO Is Nothing Then
        Exit Sub
    Else
        If oReferencesCollection0fCAO.IJDEditJDArgument Is Nothing Then
            Exit Sub
        Else
            If oReferencesCollection0fCAO.IJDEditJDArgument.GetCount = 0 Then
                Exit Sub
            End If
        End If
    End If
    
    If TypeOf oReferencesCollection0fCAO.IJDEditJDArgument.GetEntityByIndex(1) Is IJDReferencesCollection Then
        Set oRefColl = oReferencesCollection0fCAO.IJDEditJDArgument.GetEntityByIndex(1)
        Set oMembObj = oRefColl.IJDEditJDArgument.GetEntityByIndex(1)
        If Not oMembObj Is Nothing Then
            If TypeOf oMembObj Is ISPSMemberSystem Then
                Set oMembSys = oMembObj
            End If
        End If
    Else
        Set oMembObj = oReferencesCollection0fCAO.IJDEditJDArgument.GetEntityByIndex(1)
        If Not oMembObj Is Nothing Then
            If TypeOf oMembObj Is ISPSMemberSystem Then
                Set oMembSys = oMembObj
            End If
        End If
    End If
    If Not oMembSys Is Nothing Then
        Dim x As Double, y As Double, z As Double
        Dim x1 As Double, y1 As Double, z1 As Double
        Set oMembPart = GetBottomMemberPart(oMembSys)
        oMembPart.Axis.EndPoints x, y, z, x1, y1, z1
        CP = oMembPart.CrossSection.CardinalPoint
        Dim bEnd As Boolean
        If z < z1 Then
        Else
            bEnd = True
            x = x1
            y = y1
            z = z1
        End If

        If CP <> 5 Then
            If Not oMembPart Is Nothing Then
                Call GetModelCPPositionPart(CP, oMembPart, x, y, bEnd)
            End If
        End If
    Else
        If TypeOf pObj Is SPSFooting Then
            Set oFooting = pObj
            oFooting.GetPosition x, y, z
        End If
    End If
    
    Dim pOcc As IJDOccurrence
    Dim Matrix As IJDT4x4
    Set Matrix = New DT4x4
    Matrix.LoadIdentity

    Matrix.IndexValue(12) = x
    Matrix.IndexValue(13) = y
    Matrix.IndexValue(14) = z
    Set pOcc = pObj
    pOcc.Matrix = Matrix

Exit Sub
Errx:
ErrorHandler:      HandleError MODULE, METHOD
End Sub

'*************************************************************************
'Function
'CMMigrateAggregator
'
'Abstract
'Migrates thr footing to the correct surface if it is split.
'
'Arguments
'IJDMemberDescription interface of the member
'
'Return
'
'Exceptions
'
'***************************************************************************
Public Sub MigrateFootingAggregator(oAggregatorDesc As IJDAggregatorDescription, oMigrateHelper As IJMigrateHelper)

  Const MT = "CMMigrateAggregator"
  On Error GoTo ErrorHandler
  
    Dim oSmartOcc As IJSmartOccurrence
    Dim oRefCollAsm As IJDReferencesCollection
    Dim oEditArgs As IJDEditJDArgument
    Dim oObjectsReplacing() As Object
    Dim bIsInputMigrated As Boolean
    Dim oMembSys As ISPSMemberSystem
    Dim oMembPart As ISPSMemberPartPrismatic
    Dim oPoint As IJPoint
    Dim oGeom3DFactory As New GeometryFactory
    Dim iRCCount As Integer
    Dim iPlaneIdx As Long
    Dim ii As Long
    Dim oEntity As Object
        
    Set oPoint = oGeom3DFactory.Points3d.CreateByPoint(Nothing, 0#, 0#, 0#)
    Set oGeom3DFactory = Nothing
    
    Set oSmartOcc = oAggregatorDesc.CAO
    Set oRefCollAsm = GetRefCollection(oSmartOcc)
    
    Dim x As Double, y As Double, z As Double
    Dim x1 As Double, y1 As Double, z1 As Double
    
    Set oEditArgs = oRefCollAsm.IJDEditJDArgument
     
    iRCCount = oEditArgs.GetCount
    'Only call migrate for the refcoll with support plane (number of entities in RC more than 1)
    If iRCCount > 1 Then
        'Though Memebr system is first in the RC, buut still to make sure find out from the RC
        iPlaneIdx = 0
        For ii = 1 To iRCCount
           Set oEntity = oEditArgs.GetEntityByIndex(ii)
           If TypeOf oEntity Is ISPSMemberSystem Then
                Set oMembSys = oEntity
                Set oMembPart = GetBottomMemberPart(oMembSys)
                oMembPart.Axis.EndPoints x, y, z, x1, y1, z1
                If z < z1 Then
                    oPoint.SetPoint x, y, z
                Else
                    oPoint.SetPoint x1, y1, z1
                End If
           Else
                iPlaneIdx = ii
           End If
        Next ii
        
        'RC did not have all the inputs correctly defined
        If oMembSys Is Nothing Or iPlaneIdx = 0 Then
            Exit Sub
        End If
        
        MigrateRefColl oRefCollAsm, oMigrateHelper, oObjectsReplacing, bIsInputMigrated, oPoint
     
        If bIsInputMigrated And UBound(oObjectsReplacing) > 0 Then
            'If any of the inputs are indeed migrated, reset them on the ref coll
            Call oEditArgs.RemoveByIndex(iPlaneIdx)
            
            oEditArgs.SetEntity iPlaneIdx, oObjectsReplacing(2), IJPlane, "SPSSuppPlaneToRC_DEST"
        End If
    End If
    
  Exit Sub
ErrorHandler:  HandleError MODULE, MT
End Sub


Public Sub ChangePointsZToSameZAsFirstPoint(ByVal Points As IJElements)
Const MT = "ChangePointsZToSameZAsFirstPoint"
On Error GoTo ErrorHandler
    
    If Points.count < 2 Then ' nothing to do
        Exit Sub
    End If
    
    Dim z1 As Double
    Dim x As Double, y As Double, z As Double
    Dim oPnt As IJPoint
    
    If TypeOf Points.Item(1) Is IJPoint Then
        Set oPnt = Points.Item(1)
        oPnt.GetPoint x, y, z1
        Set oPnt = Nothing
    End If
    
    Dim i As Long
    For i = 2 To Points.count
        If TypeOf Points.Item(i) Is IJPoint Then
            Set oPnt = Points.Item(i)
            oPnt.GetPoint x, y, z
            If (z <> z1) Then
                oPnt.SetPoint x, y, z1
            End If
            Set oPnt = Nothing
        End If
    Next i
  
    Exit Sub
    
ErrorHandler:  HandleError MODULE, MT
End Sub

Private Function IsBuiltUPMemeberSystem(oMembSystem As ISPSMemberSystem) As Boolean
Const MT = "IsBuiltUPMemeberSystem"
On Error GoTo ErrorHandler
    Dim oSPSMemberPartCommon As ISPSMemberPartCommon
    IsBuiltUPMemeberSystem = False
    
    If Not oMembSystem Is Nothing Then
        Set oSPSMemberPartCommon = oMembSystem.MemberPartAtEnd(SPSMemberAxisStart)
        If Not oSPSMemberPartCommon Is Nothing Then
            If oSPSMemberPartCommon.IsPrismatic = False Then
                IsBuiltUPMemeberSystem = True
            End If
        End If
    End If
    Exit Function
ErrorHandler:  HandleError MODULE, MT
End Function

Public Function GetPlanAngleFromSymbolOccurrence(oSymbOcc As IJDOccurrence) As Double
   
    GetPlanAngleFromSymbolOccurrence = GetPlanAngleFromMatrix(oSymbOcc.Matrix)
End Function

Public Function GetPlanAngleFromMatrix(oMatrix As IJDT4x4) As Double
    
    Dim oVecX As New DVector
    Dim oVecXGlobal As New DVector
    Dim oVecZGlobal As New DVector
    Dim dotprod As Double
    Dim angle As Double
    
    oVecXGlobal.Set 1, 0, 0
    oVecZGlobal.Set 0, 0, 1
    
    oVecX.Set oMatrix.IndexValue(0), oMatrix.IndexValue(1), oMatrix.IndexValue(2)
    oVecX.length = 1

    angle = oVecXGlobal.angle(oVecX, oVecZGlobal)
    If angle > PI Then ' footing code is not able to handle values greater than 180. for example 350 is not ok but -10 is ok
    'easier to do it here than fix all the footing code
        angle = angle - 2 * PI
    End If
    GetPlanAngleFromMatrix = angle

    
End Function

'This method updates the rotation property of a footing component (component is a seperate Object).
'Well,Rotation is not updated if the footing is attached to member and the component's orientation is local.
'The Rotation is set to a value that would be the roation of the component, about global Z
'measured from global X, if the component is transformed by the input matrix.
Public Sub UpdateRotationAngle(oFootingComponent As SPSFootingComponent, strComponent As String, strRotation As String, strOrientation As String, Trans4x4 As IJDT4x4, Optional bIgnoreOrientation As Boolean = False)
    Dim oCompAttribs As IJDAttributes
    Dim CompRotationAngle As Double
    Dim CompOrientation As Long
    Dim pOcc As IJDOccurrence
    Dim Matrix As IJDT4x4
    Dim angle As Double
    
    
    Set oCompAttribs = oFootingComponent
    CompRotationAngle = oCompAttribs.CollectionOfAttributes(strComponent).Item(strRotation).Value
    CompOrientation = oCompAttribs.CollectionOfAttributes(strComponent).Item(strOrientation).Value
    If CompOrientation = 1 Or bIgnoreOrientation Then '1 is global and 2 is Local to the SPSMemberPart
                                                      'for components which are combined the orientation   value is ignored
                                                      'as it is always oriented globally.orientation is also ignored
                                                      'for footing placed by point
        
        Set pOcc = oFootingComponent
        Set Matrix = pOcc.Matrix.Clone
        
        Matrix.MultMatrix Trans4x4
        angle = GetPlanAngleFromMatrix(Matrix)
        oCompAttribs.CollectionOfAttributes(strComponent).Item(strRotation).Value = angle
    End If

End Sub
'This method updates the rotation property of a footing component (component is not a seperate Object).
'Well,Rotation is not updated if the footing is attached to member and the component's orientation is local.
'The Rotation is set to a value that would be the roation of the component, about global Z
'measured from global X, if the component is transformed by the input matrix.
Public Sub UpdateComponentRotationAngles(oFooting As SPSFooting, strComponent As String, strRotation As String, strOrientation As String, Trans4x4 As IJDT4x4, Optional bIgnoreOrientation As Boolean = False)
    Dim oFootingAttribs As IJDAttributes
    Dim CompRotationAngle As Double
    Dim CompOrientation As Long
    Dim Matrix As IJDT4x4
    Dim angle As Double
    Dim oVec As New DVector
    
    Set oFootingAttribs = oFooting
    CompRotationAngle = oFootingAttribs.CollectionOfAttributes(strComponent).Item(strRotation).Value
    CompOrientation = oFootingAttribs.CollectionOfAttributes(strComponent).Item(strOrientation).Value
    If CompOrientation = 1 Or bIgnoreOrientation Then '1 is global and 2 is Local to the SPSMemberPart
                                                      'for components which are combined the orientation   value is ignored
                                                      'as it is always oriented globally. orientation is also ignored
                                                      'for footing placed by point

        Set Matrix = New DT4x4
        Matrix.LoadIdentity
        oVec.Set 0, 0, 1
        Matrix.Rotate CompRotationAngle, oVec
        
        Matrix.MultMatrix Trans4x4
        angle = GetPlanAngleFromMatrix(Matrix)
        oFootingAttribs.CollectionOfAttributes(strComponent).Item(strRotation).Value = angle
    End If

End Sub

Public Function IsFootingPlacedByMember(oFooting As SPSFooting) As Boolean

    Dim oSuppedObj As Object
    Dim oRefCollCAO  As IJDReferencesCollection
    Dim oRefColl  As IJDReferencesCollection

    IsFootingPlacedByMember = False

    Set oRefCollCAO = GetRefCollection(oFooting)
    
    If Not oRefCollCAO Is Nothing Then
        If oRefCollCAO.IJDEditJDArgument.GetCount > 0 Then
            If TypeOf oRefCollCAO.IJDEditJDArgument.GetEntityByIndex(1) Is IJDReferencesCollection Then
                Set oRefColl = oRefCollCAO.IJDEditJDArgument.GetEntityByIndex(1)
                Dim enumJDArgument As IEnumJDArgument
                Dim argument As IJDArgument
                Dim found As Long
                
                'TR-256127 if RefColl not properly up to date, named relations may no longer exist
                'using an enumerator instead
                Set enumJDArgument = oRefColl
                If Not enumJDArgument Is Nothing Then
                    enumJDArgument.Reset
                    found = 1
                    
                    Do While (found <> 0)
                        enumJDArgument.Next 1, argument, found
                        If found <> 0 Then
                            Set oSuppedObj = argument.Entity
                            If TypeOf oSuppedObj Is ISPSMemberSystem Then
                                IsFootingPlacedByMember = True
                                Exit Do 'We are placed by member.  Nothing left to do
                            End If
                        End If
                    Loop
                End If
                Set enumJDArgument = Nothing
            Else
                Set oSuppedObj = oRefCollCAO.IJDEditJDArgument.GetEntityByIndex(1)
                If TypeOf oSuppedObj Is ISPSMemberSystem Then
                    IsFootingPlacedByMember = True
                End If
            End If
        End If
    End If
End Function
'this method is called to update the Z elevation value of a footing, when the footing
'is being transformed
Public Sub UpdateElevationDatumOnTransform(oFooting As SPSFooting, Trans4x4 As IJDT4x4)
Const MT = "UpdateElevationDatumOnTransform"
On Error GoTo ErrHandler
    Dim oAttr As IJDAttributes
    Dim oCollProxy As IJDAttributesCol
    Dim oDatumPos As New DPosition
    Dim vZVal As Variant
   
    Set oAttr = oFooting
    oDatumPos.Set 0, 0, 0
    On Error Resume Next ' this interface was added later so may not exist in a migrated model
    Set oCollProxy = oAttr.CollectionOfAttributes(STRUCT_ELEVATION_DATUM)
    On Error GoTo ErrHandler
        
    If Not oCollProxy Is Nothing Then
        vZVal = oCollProxy.Item(BOTTOM_ELEVATION).Value
        If Not IsEmpty(vZVal) Then
            oDatumPos.z = vZVal
            Set oDatumPos = Trans4x4.TransformPosition(oDatumPos)
            oCollProxy.Item(BOTTOM_ELEVATION).Value = oDatumPos.z
        End If
    End If
    Exit Sub
    
ErrHandler:
    HandleError MODULE, MT
End Sub

Public Sub UpdateElevationDatumFromPlane(oFooting As SPSFooting)
Const MT = "UpdateElevationDatumFromPlane"
On Error GoTo ErrHandler
    
    Dim oPlane As IJPlane
    Dim ii As Long, jj As Long, count As Long, countInner As Long
    Dim obj As Object
    Dim oReferencesCollection0fCAO  As IJDReferencesCollection
    Dim oReferencesCollectionInner  As IJDReferencesCollection
    Dim x As Double, y As Double, z As Double
    Dim oAttr As IJDAttributes
    Dim bUseDatum As Boolean
    Dim oCollProxy As IJDAttributesCol
   
    Set oAttr = oFooting
    On Error Resume Next ' this interface was added later so may not exist in a migrated model
    Set oCollProxy = oAttr.CollectionOfAttributes(STRUCT_ELEVATION_DATUM)
    On Error GoTo ErrHandler
        
 
    Set oReferencesCollection0fCAO = GetRefCollection(oFooting) ' refcoll of the footing
    count = oReferencesCollection0fCAO.IJDEditJDArgument.GetCount
    For ii = 1 To count
        Set obj = oReferencesCollection0fCAO.IJDEditJDArgument.GetEntityByIndex(ii)
        If Not obj Is Nothing Then
            If TypeOf obj Is IJPlane Then 'This is a single footing
                Set oPlane = obj
                Exit For
            ElseIf TypeOf obj Is IJDReferencesCollection Then ' This is combined footing which has 2 inner refcolls. One for
            'members or points and another one for supporting plane
                Set oReferencesCollectionInner = obj
                countInner = oReferencesCollectionInner.IJDEditJDArgument.GetCount
                For jj = 1 To countInner
                    Set obj = oReferencesCollectionInner.IJDEditJDArgument.GetEntityByIndex(jj)
                    If Not obj Is Nothing Then
                        If TypeOf obj Is IJPlane Then
                            Set oPlane = obj
                            Exit For
                        End If
                    End If
                Next jj
                If Not oPlane Is Nothing Then
                    Exit For
                End If
            End If
        End If
    Next ii
    'update the Z elevation value using the root point of the plane
    If (Not oCollProxy Is Nothing) And (Not oPlane Is Nothing) Then
        oPlane.GetRootPoint x, y, z
        oCollProxy.Item(BOTTOM_ELEVATION).Value = z
        oCollProxy.Item(USE_ELEVATION_DATUM).Value = False
    End If
    
    Exit Sub
    
ErrHandler:
    HandleError MODULE, MT
End Sub





Public Function ValidateBottomZValueForCombinedFooting(oFooting As ISPSFooting, zVal As Double) As Boolean
Const METHOD = "ValidateBottomZValueForCombinedFooting"
On Error GoTo ErrorHandler

    Dim jj As Long, count As Long, countInner As Long
    Dim obj As Object
    Dim oReferencesCollection0fCAO  As IJDReferencesCollection
    Dim oReferencesCollectionInner  As IJDReferencesCollection
    Dim x1 As Double, y1 As Double, z1 As Double
    Dim x2 As Double, y2 As Double, z2 As Double
    Dim oPoint As IJPoint
    Dim oCurve As IJCurve
    
    ValidateBottomZValueForCombinedFooting = False 'set to invalid till we compare against inputs
 
    Set oReferencesCollection0fCAO = GetRefCollection(oFooting) ' refcoll of the footing
    count = oReferencesCollection0fCAO.IJDEditJDArgument.GetCount

    If count > 0 Then
        Set obj = oReferencesCollection0fCAO.IJDEditJDArgument.GetEntityByIndex(1)
        If Not obj Is Nothing Then
            If TypeOf obj Is IJDReferencesCollection Then ' This is a combined footing which has 2 inner refcolls. first One for
            'members or points and second one for supporting plane.
                Set oReferencesCollectionInner = obj
                countInner = oReferencesCollectionInner.IJDEditJDArgument.GetCount
                For jj = 1 To countInner
                    Set obj = oReferencesCollectionInner.IJDEditJDArgument.GetEntityByIndex(jj)
                    If Not obj Is Nothing Then
                        If TypeOf obj Is IJPoint Then
                            Set oPoint = obj
                            oPoint.GetPoint x1, y1, z1
                            If zVal >= (z1 - TOL) Then
                                GoTo InValidZValue
                            End If
                        ElseIf TypeOf obj Is IJCurve Then
                            Set oCurve = obj
                            oCurve.EndPoints x1, y1, z1, x2, y2, z2
                            If (zVal >= (z1 - TOL)) Or (zVal >= (z2 - TOL)) Then
                                GoTo InValidZValue
                            End If
                        End If
                    End If
                Next jj
            End If
        End If
    End If
    ValidateBottomZValueForCombinedFooting = True ' ZVal is good
Exit Function
ErrorHandler:  HandleError MODULE, METHOD
InValidZValue:
End Function



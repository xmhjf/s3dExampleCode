Attribute VB_Name = "Common"
'*******************************************************************
'
'Copyright (C) 2006 Intergraph Corporation. All rights reserved.
'
'File : common.bas
'
'Author : RP
'
'Description :
'    SmartPlant Structural common feature macro bas file
'
'History:
'
' 05/21/03  JS    Replaced the Msgbox in HandleError
'                           with JServerErrors AddFromErr method
' 02/02/04  JS    Modified ComputePlanarCutback to check if
'                 and cutback planes exist on the member and
'                 if not add the plane from the AC. This
'                 should only occur during a copy operation (TR#52040).
' 02/16/04  RP    Added a parameter, resource manager, to the
'                 cope and cutback creation methods. When resouce
'                 manager is nothing, the plane or wire body is not
'                 persisted.
'                 Moved the code that hides the plane/wirebody to the
'                 CMConstruct().
' 06/13/06  RP    Changes due to impact from curved members. Aded methods to create cutback
'                 surface, solid cutter and axis for the cutter (DI#84001)
' 11/02/06  RP              TR#108733 and DI#84749 - Copelength and depth not computed. Also
'                           moved compute of these properties of aggregator section to avoid
'                           extra compute
' 06/09/08  AS    TR#132553. Added PostMissingInputsError to post correct error msgs for feature obejcts with missing inputs
'***************************************************************************************************
Option Explicit
Private Const MODULE = "Common"
Public Const IJPlane = "{4317C6B3-D265-11D1-9558-0060973D4824}"
Public Const E_FAIL = -2147467259
Public Sub ComputePlanarCutback(ByVal oSmartOcc As IJSmartOccurrence, ByRef oPlane As IJPlane)
Const MT = "ComputePlanarCutback"
    On Error GoTo ErrorHandler

    Dim oSuppedPort As ISPSSplitAxisPort
    Dim oRefColl   As IJDReferencesCollection
    Dim oSuppedPart As ISPSMemberPartPrismatic
    Dim oGeomFactory As New GeometryFactory
    Dim oInputPlane As IJPlane
    Dim oX#, oY#, oZ#, nx#, ny#, nz#
    Dim pIJAttribsCAO As IJDAttributes
    Dim oAttrbs As IJDAttributes
    Dim Clearance As Double
    Dim SquareEnd As Boolean
    Dim oVec1 As New DVector, oVec2 As New DVector
    Dim sX#, sY#, sZ#, eX#, eY#, eZ#
    Dim cosA As Double
    Dim pObject As IJDObject
    Dim oStructHelper As IJStructSymbolTools
    Dim oPlate As IJStructCustomPlatePart
    Dim oPort As IJPort
    Dim oInputObj2 As Object, oInputObj3 As Object
    Dim oSuppingPort As ISPSSplitAxisPort
    Dim oSuppingPart As ISPSMemberPartPrismatic
    Dim oInputPlane1 As IJPlane
    Dim iPort As SPSMemberAxisPortIndex
    Dim oPos As IJDPosition, oPos1 As IJDPosition
    Dim dist1 As Double, dist As Double
    
    Dim computeIt As Boolean

    Set pObject = oSmartOcc
    Set oRefColl = GetRefCollFromSmartOccurrence(oSmartOcc)
    Set pIJAttribsCAO = oSmartOcc
    
    computeIt = False
    If IsAttributeModified(oSmartOcc) Then
        computeIt = True
    End If

    Clearance = pIJAttribsCAO.CollectionOfAttributes("IJUASPSPlanarCutback").Item("Clearance").Value
    SquareEnd = pIJAttribsCAO.CollectionOfAttributes("IJUASPSPlanarCutback").Item("SquaredEnd").Value

    Set oSuppedPort = oRefColl.IJDEditJDArgument.GetEntityByIndex(1)
    Set oSuppedPart = oSuppedPort.Part
    iPort = oSuppedPort.portIndex
   
    If Not computeIt Then
        If IsItModified(oSuppedPart) Then
            computeIt = True
        End If
    End If

    oSuppedPart.Axis.EndPoints sX, sY, sZ, eX, eY, eZ
    'get the vector towards the other end
    If iPort = SPSMemberAxisStart Then
        oVec2.Set eX - sX, eY - sY, eZ - sZ
    Else
        oVec2.Set sX - eX, sY - eY, sZ - eZ
    End If
    'normalize
    oVec2.length = 1
   
    Set oInputObj2 = oRefColl.IJDEditJDArgument.GetEntityByIndex(2)
    Set oInputObj3 = oRefColl.IJDEditJDArgument.GetEntityByIndex(3)
    
    If Not computeIt Then
        If IsItModified(oInputObj2) Then
            computeIt = True
        End If
    End If
    
    If TypeOf oInputObj2 Is IJPlane Then

        Set oInputPlane = oInputObj2
        oInputPlane.GetNormal nx, ny, nz
        oVec1.Set nx, ny, nz
        'normalize
        oVec1.length = 1
        oInputPlane.GetRootPoint oX, oY, oZ
        cosA = oVec1.Dot(oVec2)
        'project the plane along the normal to get the required clearance between the input plane and member
        If cosA < 0 Then ' the plane normal is away from oVec2
            oX = oX - Clearance * oVec1.x
            oY = oY - Clearance * oVec1.y
            oZ = oZ - Clearance * oVec1.z
        Else
            oX = oX + Clearance * oVec1.x
            oY = oY + Clearance * oVec1.y
            oZ = oZ + Clearance * oVec1.z
        End If
        Set oInputPlane = oGeomFactory.Planes3d.CreateByPointNormal(Nothing, oX, oY, oZ, nx, ny, nz)
    ElseIf TypeOf oInputObj2 Is ISPSSplitAxisPort Then

        Set oSuppingPort = oInputObj2
        Set oSuppingPart = oSuppingPort.Part
        'get the cutback plane
        CreateCutback oSuppingPart, oSuppedPart, iPort, pIJAttribsCAO, Nothing, oInputPlane
    ElseIf TypeOf oInputObj2 Is IJStructCustomPlatePart Then
        If Not oPlane Is Nothing Then
            Set oPlate = oInputObj2
            Set oStructHelper = New StructSymbolTools
            'get top face
            oStructHelper.GetLateBindingPort oPlate, JS_TOPOLOGY_PROXY_LFACE, 0, 0, 138, 0, oPort
            Set oInputPlane = oPort
            'get the max cut dist and pos for the cutback plane
            GetCutDistAndPosFromPlane oSuppedPart, iPort, oInputPlane, oPos, dist
            
            'get bottom face
            oStructHelper.GetLateBindingPort oPlate, JS_TOPOLOGY_PROXY_LFACE, 0, 0, 138, -1, oPort
            Set oInputPlane1 = oPort
            GetCutDistAndPosFromPlane oSuppedPart, iPort, oInputPlane1, oPos1, dist1
            If dist1 < dist Then
                Set oInputPlane = oInputPlane1
                Set oPos = oPos1
            End If
        Else

        End If
    End If
    If Not oInputObj3 Is Nothing Then

        If Not computeIt Then
            If IsItModified(oInputObj3) Then
                computeIt = True
            End If
        End If

        If TypeOf oInputObj3 Is IJPlane Then
            Set oInputPlane1 = oInputObj3
        ElseIf TypeOf oInputObj3 Is ISPSSplitAxisPort Then
            Set oSuppingPort = oInputObj3
            Set oSuppingPart = oSuppingPort.Part
            'get the cutback plane
            CreateCutback oSuppingPart, oSuppedPart, iPort, pIJAttribsCAO, Nothing, oInputPlane1

        ElseIf TypeOf oInputObj3 Is IJStructCustomPlatePart Then
        
        End If
    End If
    'get the max cut dist and pos for the cutback plane
    GetCutDistAndPosFromPlane oSuppedPart, iPort, oInputPlane, oPos, dist
    
    If Not oInputPlane1 Is Nothing Then
        'get the max cut dist and pos for the cutback plane
        GetCutDistAndPosFromPlane oSuppedPart, iPort, oInputPlane1, oPos1, dist1
        If dist1 < dist Then
            Set oInputPlane = oInputPlane1
            Set oPos = oPos1
        End If
    End If
    
    
    oPos.Get oX, oY, oZ
    If SquareEnd = True Then
        oVec2.Get nx, ny, nz
    Else
        oInputPlane.GetNormal nx, ny, nz
    End If
   

    If oPlane Is Nothing Then

        Set oPlane = oGeomFactory.Planes3d.CreateByPointNormal(pObject.ResourceManager, oX, oY, oZ, nx, ny, nz)
        
        oSuppedPart.AddCutbackSurface oSuppedPort.portIndex, oPlane
        
    ElseIf computeIt Then
        oPlane.SetRootPoint oX, oY, oZ
        oPlane.SetNormal nx, ny, nz
    
    End If

        ' For a copy operation the cutback plane has not be established
    If oSuppedPart.Cutbacks(oSuppedPort.portIndex).count < 1 Then
        oSuppedPart.AddCutbackSurface oSuppedPort.portIndex, oPlane
    End If

  Exit Sub
ErrorHandler:      HandleError MODULE, MT
End Sub
Public Sub ModifyCopeShape(oSupping As ISPSMemberPartPrismatic, oSupped As ISPSMemberPartPrismatic, iEnd As SPSMemberAxisPortIndex, bTopCope As Boolean, ByRef oAttribsCAO As IJDAttributes, ByRef oCmplx As ComplexString3d)
Const MT = "ModifyCopeShape"
    On Error GoTo ErrorHandler
    
    Dim oMembObjetcs As IJDMemberObjects
    Dim oPlane As IJPlane
    Dim oSurf As IJSurface
    Dim oElms As IJElements
    Dim code As Geom3dIntersectConstants
    Dim oCopePlane As IJPlane
    Dim oPoint(1 To 2) As IJPoint
    Dim oCopePoint As IJPoint
    Dim x#, y#, z#, nx#, ny#, nz#, x1#, y1#, z1#
    Dim dblCopeLength As Double, dblCopedepth As Double
    Dim oLine3d(1 To 4) As Line3d
    Dim oGeomFactory As New GeometryFactory
    Dim oMatCtoG As IJDT4x4, oMatMtoG As IJDT4x4, oMatGtoC As IJDT4x4, oMatGtoM As IJDT4x4
    Dim idx As Integer
    Dim oVec As New DVector
    Dim oPosAlong As IJDPosition
    Dim bMirror As Boolean, bInPosY As Boolean, bModifyStartPos As Boolean
    
    'we need to get the cutback plane from the custom assembly
    Set oMembObjetcs = oAttribsCAO
    Set oSurf = oMembObjetcs.ItemByDispid(1)
    ' for top cope shape the curves of the cope shape are as below. The supported member is on curve 2 side
    '     1
    '  --------
    ' 4|      | 2
    '  |      |
    '  --------
    '     3
    
    ' for bottom cope shape the curves of the cope shape are as below. The supported member is on curve 2 side
    '     3
    '  --------
    ' 4|      | 2
    '  |      |
    '  --------
    '     1
    Set oPosAlong = GetConnectionPositionOnSupping(oSupping, oSupped, iEnd)
    
    oSupping.Rotation.GetTransformAtPosition oPosAlong.x, oPosAlong.y, oPosAlong.z, oMatMtoG, Nothing
    Set oMatGtoM = oMatMtoG.Clone()
    oMatGtoM.Invert ' matrix to transform from global to member coordinates
    
    bMirror = oSupping.Rotation.Mirror
    Set oMatCtoG = CreateCSToMembTransform(oMatMtoG, bMirror)
    Set oMatGtoC = oMatCtoG.Clone()
    
    oMatGtoC.Invert 'global to cross section coordinates
    
    oCmplx.GetCurve 3, oLine3d(3) 'get the third curve added
    
    
    'transform to global
    oLine3d(3).Transform oMatCtoG
    oSurf.Intersect oLine3d(3), oElms, code
  
    If Not oElms Is Nothing Then
        If oElms.count > 0 Then
            Set oPoint(1) = oElms.Item(1) ' the intersection point
        
            'get cope depth and length
            dblCopedepth = oAttribsCAO.CollectionOfAttributes("IJUASPSCope").Item("Depth").Value
            dblCopeLength = oAttribsCAO.CollectionOfAttributes("IJUASPSCope").Item("Length").Value
        
            oCmplx.GetCurves oElms
            For idx = 1 To 4
                oCmplx.RemoveCurve idx
            Next
            oPoint(1).GetPoint x, y, z ' curve3 intersection point with cutback plane
            
            bInPosY = IsSupportedAxisInPositiveY(oSupping, oSupped, iEnd)
            If (bInPosY And Not bMirror) Or (Not bInPosY And bMirror) Then
                bModifyStartPos = True
            End If
            If bModifyStartPos Then
                oLine3d(3).GetEndPoint x1, y1, z1
            Else
                oLine3d(3).GetStartPoint x1, y1, z1
            End If
        
            oVec.Set x1 - x, y1 - y, z1 - z
            oVec.length = 1
            'size curve3 vector to cope length
            x = x + oVec.x * dblCopeLength
            y = y + oVec.y * dblCopeLength
            z = z + oVec.z * dblCopeLength
            
            Set oPoint(2) = oGeomFactory.Points3d.CreateByPoint(Nothing, x, y, z)
            
            Set oCopePlane = oGeomFactory.Planes3d.CreateByPointNormal(Nothing, x, y, z, oVec.x, oVec.y, oVec.z)
            'based on the section range box of the supported, get the top most or bottom most intersection point
            'on the cope plane. bTopCope=True for Top, Bottom otherwise
            GetCopePoint oSupping, oSupped, iEnd, oCopePlane, bTopCope, oCopePoint
            
            oPoint(2).Transform oMatGtoC 'global to section
            'oPoint(2) in supporting local coordinates now
            
            oPoint(2).GetPoint x, y, z

            oCopePoint.Transform oMatGtoC 'global to section
            oCopePoint.GetPoint x1, y1, z1

            oVec.Set 0, y - y1, 0 'oVec parallel to section y axis
            oVec.length = 1 ' normalize
            y = y1 + oVec.y * dblCopedepth ' adjust copepoint to take care of copedepth
           
                               
           
            
            Set oLine3d(1) = oElms.Item(1) 'get the first curve added
            
            If bModifyStartPos Then
                oLine3d(1).GetEndPoint x1, y1, z1
                oLine3d(1).SetStartPoint x, y1, 0
                oLine3d(1).SetEndPoint x1, y1, 0
            Else
                oLine3d(1).GetStartPoint x1, y1, z1
                oLine3d(1).SetStartPoint x1, y1, 0
                oLine3d(1).SetEndPoint x, y1, 0
            End If
            
            Set oLine3d(2) = oElms.Item(2)
            If bModifyStartPos Then
                oLine3d(2).GetStartPoint x1, y1, z1
                oLine3d(2).SetStartPoint x1, y1, 0
                oLine3d(2).SetEndPoint x1, y, 0
            Else
                oLine3d(2).GetStartPoint x1, y1, z1
                oLine3d(2).SetStartPoint x, y1, 0
                oLine3d(2).SetEndPoint x, y, 0
            End If
            
            Set oLine3d(4) = oElms.Item(4)
            
            If bModifyStartPos Then
                oLine3d(4).SetStartPoint x, y, 0
                oLine3d(4).GetEndPoint x1, y1, z1
                oLine3d(4).SetEndPoint x, y1, 0
            Else
                oLine3d(4).GetEndPoint x1, y1, z1
                oLine3d(4).SetStartPoint x1, y, 0
                oLine3d(4).SetEndPoint x1, y1, 0
            End If
           
            
            Set oLine3d(3) = oElms.Item(3) 'get the third curve added
            'we transformed line3 to global before so need to transform back
'            oLine3d(3).Transform oMatGtoC
            If bModifyStartPos Then
                oLine3d(3).GetStartPoint x1, y1, z1
                oLine3d(3).SetStartPoint x1, y, 0
                oLine3d(3).SetEndPoint x, y, 0
            Else
                oLine3d(3).SetStartPoint x, y, 0
                oLine3d(3).GetEndPoint x1, y1, z1
                oLine3d(3).SetEndPoint x1, y, 0
            End If
            
            oCmplx.SetCurves oElms
        End If
    End If
    Exit Sub
ErrorHandler:      HandleError MODULE, MT
End Sub
Public Function RoundOff(dblIncrement As Double, dblDist As Double) As Double
    Const METHOD = "RoundOff"
    On Error GoTo ErrorHandler
    Dim dblResult#
    Dim intResult As Integer
    If dblIncrement > distTol Then
    dblResult = dblDist / dblIncrement ' returns double value
    intResult = Int(-1 * dblResult) ' to make sure that we rounding to the next higher integer
    ' int(-6.1) returns -7. So we are converting dblResult to negative and then intResult to positive
    RoundOff = (-1 * intResult) * dblIncrement
    Else
        RoundOff = dblDist
    End If
    Exit Function
ErrorHandler:      HandleError MODULE, METHOD
End Function

''///////////////////////////////////////////////////////////////////////
''// IsReadOnlyObject()
''// Desc:
''// This function will return True or False depending on whether the
''// user has ReadOnly permissions on the object
''////////////////////////////////////////////////////////////////////////
Public Function IsReadOnlyObject(obj As IJDObject) As Boolean
    Const METHOD = "IsReadOnlyObject"
    On Error GoTo ErrHandler
    
    If obj.AccessControl And acUpdate Then
        IsReadOnlyObject = False
    Else
        IsReadOnlyObject = True
    End If
    Exit Function

ErrHandler:
     HandleError MODULE, METHOD
End Function

Public Function IsNewObject(obj As Object) As Boolean

    Const METHOD = "IsNewObject"
    On Error GoTo ErrHandler
    
    Dim flags As Long
    Dim iCompute As IJStructAssocCompute

    If obj Is Nothing Then      'if obj is nothing, we want to compute it.
        flags = &H200
    Else
        Set iCompute = New StructAssocTools
        iCompute.GetAssocFlags obj, flags
        flags = flags And &H200     'RELATION_INSERTED_IN_TRANSACTION
    End If

    If flags = &H200 Then
        IsNewObject = True
    Else
        IsNewObject = False
    End If
    
    Exit Function
    
ErrHandler:
     HandleError MODULE, METHOD
End Function

Public Function IsAttributeModified(obj As Object) As Boolean

    Const METHOD = "IsAttributeModified"
    On Error GoTo ErrHandler
    
    Dim flags As Long
    Dim iCompute As IJStructAssocCompute

    Set iCompute = New StructAssocTools
    iCompute.GetAssocFlags obj, flags, "IJDAttributes"

    ' pending TR 60413, we just set this to always be modified which makes us compute too much.
    flags = &H200000
    
    flags = flags And &H200000      'RELATION_MODIFIED_IN_TRANS

    If flags = &H200000 Then
        IsAttributeModified = True
    Else
        IsAttributeModified = False
    End If
    
    Exit Function
    
ErrHandler:
     HandleError MODULE, METHOD
End Function

Public Function IsItModified(PartOrPort As Object) As Boolean

    Const METHOD = "IsItModified"
    On Error GoTo ErrHandler
    Dim iPort As IJPort
    Dim iObj As Object
    Dim iDesignNotify As ISPSPartPrismaticDesignNotify
    Dim IsNew As Boolean, TypeChanged As Boolean, AxisChanged As Boolean, RotationChanged As Boolean, CrossSectionChanged As Boolean, MaterialChanged As Boolean
    Dim flags As Long
    Dim iCompute As IJStructAssocCompute
    
    If TypeOf PartOrPort Is IJPort Then
        Set iPort = PartOrPort
        Set iObj = iPort.Connectable
    Else
        Set iObj = PartOrPort
    End If
    
    If TypeOf iObj Is ISPSPartPrismaticDesignNotify Then
        Set iDesignNotify = iObj
        iDesignNotify.GetDesignChanges IsNew, TypeChanged, AxisChanged, RotationChanged, CrossSectionChanged, MaterialChanged
    
        If IsNew Or AxisChanged Or RotationChanged Or CrossSectionChanged Then
            IsItModified = True
        Else    'nothing, or type, or material changed
            IsItModified = False
        End If

    Else
        Set iCompute = New StructAssocTools
        iCompute.GetAssocFlags iObj, flags

        flags = flags And &H200000     'RELATION_MODIFIED_IN_TRANS

        If flags = &H200000 Then
            IsItModified = True
        Else
            IsItModified = False
        End If
    End If

    Exit Function
    
ErrHandler:
     HandleError MODULE, METHOD
End Function

Public Sub AdaptOnPasteFeatureToPart(pMD As IJDMemberDescription)
    Const MT = "AdaptOnPasteFeatureToPart"
    On Error GoTo ErrorHandler
    
    Dim oSuppedPort As ISPSSplitAxisPort
    Dim oSuppedPart As ISPSMemberPartPrismatic
    Dim oRefColl   As IMSSymbolEntities.IJDReferencesCollection

    Dim oCollectionOfOperators As IJElements
    Dim oStructCutoutOperationAE As StructCutoutOperationAE
    Dim oStructOperationPattern As IJStructOperationPattern

    Set oRefColl = GetRefCollFromSmartOccurrence(pMD.CAO)
    
    Set oSuppedPort = oRefColl.IJDEditJDArgument.GetEntityByIndex(1)
    Set oSuppedPart = oSuppedPort.Part
    
    Set oStructOperationPattern = oSuppedPart
    oStructOperationPattern.GetOperationPattern "StructGeneric.StructCutoutOperationAE.1", oCollectionOfOperators, oStructCutoutOperationAE
    If oCollectionOfOperators Is Nothing Then ' Collection must exist so create it if it does not
        Set oCollectionOfOperators = New JObjectCollection
    End If
    oCollectionOfOperators.Add pMD.object
    oStructOperationPattern.SetOperationPattern "StructGeneric.StructCutoutOperationAE.1", oCollectionOfOperators, oStructCutoutOperationAE
    
    Set oSuppedPort = Nothing
    Set oSuppedPart = Nothing
    Set oRefColl = Nothing
    Set oCollectionOfOperators = Nothing

    Exit Sub

ErrorHandler:
    HandleError MODULE, MT
End Sub

Public Sub AdaptOnPasteCutbackToPart(pMD As IJDMemberDescription)
    Const MT = "AdaptOnPasteCutbackToPart"
    On Error GoTo ErrorHandler
    
    Dim oSuppedPort As ISPSSplitAxisPort
    Dim oSuppedPart As ISPSMemberPartPrismatic
    Dim oRefColl   As IMSSymbolEntities.IJDReferencesCollection

    Set oRefColl = GetRefCollFromSmartOccurrence(pMD.CAO)
    Set oSuppedPort = oRefColl.IJDEditJDArgument.GetEntityByIndex(1)
    Set oSuppedPart = oSuppedPort.Part
    
    oSuppedPart.AddCutbackSurface oSuppedPort.portIndex, pMD.object
    
    Set oSuppedPort = Nothing
    Set oSuppedPart = Nothing
    Set oRefColl = Nothing

    Exit Sub

ErrorHandler:
    HandleError MODULE, MT
End Sub

' checks whether the operator is already in the list of operators.
' if not, adds it as an operator, hides it and sets PG
'
' False is returned if the operator is ALREADY an operator
' True is returned when it was added.

Public Function AddStructCutoutOperator(operator As Object, Operand As Object) As Boolean
    Const MT = "AddStructCutoutOperator"
    On Error GoTo ErrorHandler

    Dim iControlFlags As IJControlFlags
    Dim iIJDObjectOperator As IJDObject, iIJDObjectOperand As IJDObject
    Dim oCollectionOfOperators As IJElements
    Dim oStructCutoutOperationAE As StructCutoutOperationAE
    Dim iStructOperationPattern As IJStructOperationPattern
    
    Set iStructOperationPattern = Operand
    
    iStructOperationPattern.GetOperationPattern "StructGeneric.StructCutoutOperationAE.1", oCollectionOfOperators, oStructCutoutOperationAE
    
    If oCollectionOfOperators Is Nothing Then
        Set oCollectionOfOperators = New JObjectCollection

    ElseIf oCollectionOfOperators.Contains(operator) Then
        AddStructCutoutOperator = False
        Exit Function

    End If

    oCollectionOfOperators.Add operator
    iStructOperationPattern.SetOperationPattern "StructGeneric.StructCutoutOperationAE.1", oCollectionOfOperators, oStructCutoutOperationAE

    ' hide the operator
    Set iControlFlags = operator
    iControlFlags.ControlFlags(&H4) = &H4

    ' assign PG
    Set iIJDObjectOperator = operator
    Set iIJDObjectOperand = Operand
    iIJDObjectOperator.PermissionGroup = iIJDObjectOperand.PermissionGroup
    
    AddStructCutoutOperator = True

    Exit Function

ErrorHandler:
    HandleError MODULE, MT
End Function

Public Function OwnedByConnectable(Operand As IJConnectable, operator As Object) As Boolean
    Const MT = "OwnedByConnectable"
    On Error GoTo ErrorHandler

    Dim oIJPort As IJPort
    
    OwnedByConnectable = False
    
    If TypeOf operator Is IJPort Then
        Set oIJPort = operator
        If oIJPort.Connectable Is Operand Then
            OwnedByConnectable = True
        End If
    End If

    Exit Function
    
ErrorHandler:
    HandleError MODULE, MT
End Function

Public Function GetStablePort(objPort As Object) As Object
    Const MT = "GetStablePort"
    On Error GoTo ErrorHandler

    Dim iIJPort As IJPort
    Dim oConnectable As Object
    Dim iIJStablePort As IJStablePort
    Dim oOutput As Object, oStableOutput As Object
    
    Set oOutput = objPort
    If TypeOf objPort Is IJPort Then
        Set iIJPort = objPort
        Set oConnectable = iIJPort.Connectable
        If TypeOf oConnectable Is IJStablePort Then
            Set iIJStablePort = oConnectable
            On Error Resume Next
            Set oStableOutput = iIJStablePort.StablePort(objPort)
            Err.Clear
            On Error GoTo ErrorHandler
            If oStableOutput Is Nothing Then
                Set oStableOutput = objPort
            End If
            Set oOutput = oStableOutput
        End If
    End If

    Set GetStablePort = oOutput
    Exit Function

ErrorHandler:
    HandleError MODULE, MT
End Function

Public Sub AssignPG(Feature As IJDObject, Parent As IJDObject)
    Const MT = "AssignPG"
    On Error GoTo ErrorHandler

    Feature.PermissionGroup = Parent.PermissionGroup
    
    Exit Sub

ErrorHandler:
    HandleError MODULE, MT
End Sub

Public Sub AddPartFeatureRln(Part As Object, Feature As Object)
    Const MT = "AddPartFeatureRln"
    On Error GoTo ErrorHandler
    
    Dim iIJDAssocRelation As IMSRelation.IJDAssocRelation
    Dim iIJDTargetObjectCol As IMSRelation.IJDTargetObjectCol
    Dim iIJDRelationshipCol As IMSRelation.IJDRelationshipCol

    Dim pRel As Object, existingPart As Object
    Dim pRevision As IJRevision
    
    ' first check if the feature is already connected to this part
    Set iIJDAssocRelation = Feature
    Set iIJDTargetObjectCol = iIJDAssocRelation.CollectionRelations("{C419912B-4014-4718-87DC-E476ED3B1D10}", "Part")
    
    If iIJDTargetObjectCol Is Nothing Then
        Err.Raise E_FAIL
    End If

    If iIJDTargetObjectCol.count > 0 Then
        Set existingPart = iIJDTargetObjectCol.Item(1)
        
        If existingPart Is Part Then
            GoTo wrapup
        Else
            Set iIJDRelationshipCol = iIJDTargetObjectCol
            Set pRel = iIJDRelationshipCol.Item(1)
            Set pRevision = New JRevision
            pRevision.RemoveRelationship pRel
        End If
    End If

    Set iIJDAssocRelation = Part
    Set iIJDTargetObjectCol = iIJDAssocRelation.CollectionRelations("{53B6E606-78C6-4BCA-A640-43A2258EDED1}", "Feature")
    iIJDTargetObjectCol.Add Feature, "", pRel
    
    If pRevision Is Nothing Then
        Set pRevision = New JRevision
    End If
    pRevision.AddRelationship pRel

wrapup:
    Exit Sub
    
ErrorHandler:
    HandleError MODULE, MT
End Sub

Public Sub MakeFeatureChildOfPart(oParent As Object, oFeatureChild As Object)
    Const MT = "MakeFeatureChildOfPart"
    On Error GoTo ErrorHandler

    Dim iDesignChild As IJDesignChild
    Dim iDesignParent As IJDesignParent
    
    Set iDesignChild = oFeatureChild
    If Not oParent Is iDesignChild.GetParent Then
        Set iDesignParent = oParent
        iDesignParent.AddChild oFeatureChild
    End If

    AddPartFeatureRln oParent, oFeatureChild
    AssignPG oFeatureChild, oParent

    Exit Sub

ErrorHandler:
    HandleError MODULE, MT
End Sub
'*************************************************************************
'
'<CreateCutbackSurface>
'
'Abstract
'
'<Creates/Modifes cutback surface based on the input cutback plane>
'
'Arguments
'
'<Peristenet object manager, Supported member, Supporting Member, Connected end, Cutback plane>
'
'Return
'
'<Cutback Surface>
'
'Exceptions
'*************************************************************************
Public Sub CreateCutbackSurface(pResourceManager As Object, oSupping As ISPSMemberPartPrismatic, _
oSupped As ISPSMemberPartPrismatic, iEnd As SPSMemberAxisPortIndex, _
oCutbackPlane As IJPlane, oSurface As IJSurface)
    Const MT = "CreateCutbackSurface"
    On Error GoTo ErrorHandler
    Dim oAxisCurve As IJCurve
    Dim oShortCurve As IJDModelBody
    Dim oPosAlong As IJDPosition
    Dim lengthSupped#
    Dim sX#, sY#, sZ#, eX#, eY#, eZ#
    
    'first get the supported member's intersection position on supporting curve
    Set oPosAlong = GetConnectionPositionOnSupping(oSupping, oSupped, iEnd)
    
    'get the curve segment on which the supported is framing in
    Set oAxisCurve = GetCurve(oSupping, oSupped, iEnd)
    
    'get the distance  between ends of supped member axis
    oSupped.Axis.EndPoints sX, sY, sZ, eX, eY, eZ
    lengthSupped = Sqr((sX - eX) * (sX - eX) + (sY - eY) * (sY - eY) + (sZ - eZ) * (sZ - eZ))

    'create a curve along the axis of supporting with connection point as mid point
    'and length equal to the distance between Supped axis end points
    'such that the cutter that we create has very high chance of cutting across the entire cross section
    'of the supported member at the same time not cutting the other end of the supped member
    
    Set oShortCurve = CreateShorterCurve(oAxisCurve, oPosAlong, lengthSupped)
    
    CreateSurface pResourceManager, oShortCurve, oCutbackPlane, oSurface
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, MT

End Sub


'*************************************************************************
'
'<CreateSurface>
'
'Abstract
'
'<Creates/Modifes surface by projection or revolution  using the axis and plane
'A line is first created from the plane (represents cutback position
'and normal vector to the cutback plane) and this line is either projected or revolved
'using the axis curve >
'
'Arguments
'
'<Persisten object manager, axis curve, Plane, Surface >
'
'Return
'
'<Surface>
'
'Exceptions
'*************************************************************************

Public Function CreateSurface(pResourceManager As Object, oAxisCurve As IJCurve, oPlane As IJPlane, oSurface As IJSurface)
    Const MT = "CreateSurface"
    On Error GoTo ErrorHandler

    Dim oGeomFactory As New GeometryFactory
    Dim oArc As IJArc
    Dim oLine As IJLine, oEdgeLine As IJLine
    Dim xVec As New DVector, yVec As New DVector, zVec As New DVector, oVec As New DVector
    
    
    Dim x#, y#, z#, axisLength#, dblRadius1#, dblRadius2#, xPar#, yPar#
    Dim oPos As New DPosition, oCentrePos As New DPosition
    Dim oCutbackPos As New DPosition

    If TypeOf oAxisCurve Is IJArc Then
        
        Dim oMatLtoG As New DT4x4, oMatGtoL As IJDT4x4
        Dim oPlanePos1 As New DPosition, oPlanePos2 As New DPosition
        Dim oOrigin As New DPosition
        Dim oPos1 As New DPosition, oPos2 As New DPosition
        
        Set oArc = oAxisCurve
        
        'create local coord system for the arc
        oArc.GetNormal x, y, z
        zVec.Set x, y, z
        zVec.length = 1 'normalize
        oArc.GetMajorAxis x, y, z
        xVec.Set x, y, z
        xVec.length = 1 'normalize
        Set yVec = zVec.Cross(xVec)
        yVec.length = 1 'normalize
                
        
        'get center of arc
        oArc.GetCenterPoint x, y, z
        oCentrePos.Set x, y, z
        oMatLtoG.LoadIdentity
        'create arc's transformation matrix
        oMatLtoG.IndexValue(0) = xVec.x
        oMatLtoG.IndexValue(1) = xVec.y
        oMatLtoG.IndexValue(2) = xVec.z
        
        oMatLtoG.IndexValue(4) = yVec.x
        oMatLtoG.IndexValue(5) = yVec.y
        oMatLtoG.IndexValue(6) = yVec.z
        
        oMatLtoG.IndexValue(8) = zVec.x
        oMatLtoG.IndexValue(9) = zVec.y
        oMatLtoG.IndexValue(10) = zVec.z
        
        oMatLtoG.IndexValue(12) = oCentrePos.x
        oMatLtoG.IndexValue(13) = oCentrePos.y
        oMatLtoG.IndexValue(14) = oCentrePos.z
        
        Set oMatGtoL = oMatLtoG.Clone
        
        oMatGtoL.Invert 'global to local
              
        oPlane.GetRootPoint x, y, z
        oCutbackPos.Set x, y, z
               
        'opos is in arc's local coordinates now
        Set oPos = oMatGtoL.TransformPosition(oCutbackPos)
        'remove z component.
        
        'we need to get a vector normal to plane's normal and _
        'axis of the arc
        oPos.z = 0
        oOrigin.Set 0, 0, 0
        Set yVec = oPos.Subtract(oOrigin)
        yVec.length = 1
        Set yVec = oMatLtoG.TransformVector(yVec)
        Set xVec = zVec.Cross(yVec)
        xVec.length = 1
        'get plane normal
        oPlane.GetNormal x, y, z
        zVec.Set x, y, z
        'get a vector along the cutback plane normal to arc axis and arc notmal
        Set yVec = zVec.Cross(xVec)
        
        yVec.[Scale] -1 'reverse direction
        Set oPlanePos1 = oCutbackPos.Offset(yVec)
        
        yVec.[Scale] -1 'rverse direction
        Set oPlanePos2 = oCutbackPos.Offset(yVec)
        
        'we need to move these 2 point to the start of the arc
        
        'transform to local
        Set oPlanePos1 = oMatGtoL.TransformPosition(oPlanePos1)
        Set oPlanePos2 = oMatGtoL.TransformPosition(oPlanePos2)

        Set oPos1 = oPlanePos1.Clone()
        Set oPos2 = oPlanePos2.Clone()
        
        'project to arc's plane
        oPos1.z = 0
        oPos2.z = 0
        
        dblRadius1 = oPos1.DistPt(oOrigin)
        dblRadius2 = oPos2.DistPt(oOrigin)
        
        'get start position of the arc
        oArc.GetStartPoint x, y, z
        oPos.Set x, y, z
        
        Set oPos = oMatGtoL.TransformPosition(oPos)
        
        Set yVec = oPos.Subtract(oOrigin)
        yVec.length = dblRadius1
        
        'basically create the cutback plane root position at the start of the arc
        Set oPos1 = oOrigin.Offset(yVec)
        
        yVec.length = dblRadius2
        Set oPos2 = oOrigin.Offset(yVec)
        
        'set Z to match of the original cutback position
        oPos1.z = oPlanePos1.z
        oPos2.z = oPlanePos2.z
        
        Set oPos1 = oMatLtoG.TransformPosition(oPos1)
        Set oPos2 = oMatLtoG.TransformPosition(oPos2)
        'create line of 2m length
        Set oEdgeLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oPos1.x, oPos1.y, oPos1.z, oPos2.x, oPos2.y, oPos2.z)
      
   
    ElseIf TypeOf oAxisCurve Is IJLine Then
        Set oLine = oAxisCurve
        oLine.GetDirection x, y, z
        xVec.Set x, y, z
        oPlane.GetNormal x, y, z
        yVec.Set x, y, z
        Set zVec = xVec.Cross(yVec)
    
        'we will create a line of 10 m length
    
        oPlane.GetRootPoint x, y, z

        'need to move this towards  the start of the axiscurve so that the surface
        'we generate later cut across the entire member
        'we will move this by the 1/2 the length of the line and later project by the length
        'this way the surface extend half the length on either side
        axisLength = oLine.length
        x = x - xVec.x * axisLength * 0.5
        y = y - xVec.y * axisLength * 0.5
        z = z - xVec.z * axisLength * 0.5
        oPos.Set x, y, z
    
        zVec.length = 1
        zVec.[Scale] -5 'reverse direction and scale by 5 - so the length is 5m
        
        'create another position is 5m below the current position - 5m is just arbitray
        'just need to make sure that the surface is big
        Set oPos = oPos.Offset(zVec)
        
        'reverse the direction
        zVec.[Scale] -1
        
        'create line of 10m length
        Set oEdgeLine = oGeomFactory.Lines3d.CreateByPtVectLength(Nothing, oPos.x, oPos.y, oPos.z, zVec.x, zVec.y, zVec.z, 10)
    
    End If
    
    If Not oArc Is Nothing Then
        oArc.GetNormal x, y, z
        zVec.Set x, y, z
        
        oArc.GetCenterPoint x, y, z
        
        'todo - increase sweep angle and modify start point to handle supported member framing to the end of the curve
        If oSurface Is Nothing Then
            Set oSurface = oGeomFactory.Revolutions3d.CreateByCurve(pResourceManager, oEdgeLine, zVec.x, zVec.y, zVec.z, x, y, z, oArc.SweepAngle, False)
        Else
            Dim oRevSurface As IJRevolution
            Set oRevSurface = oSurface
            oRevSurface.SweepAngle = oArc.SweepAngle
            oRevSurface.Curve = oEdgeLine
            oRevSurface.SetAxis zVec.x, zVec.y, zVec.z
            oRevSurface.SetCenter x, y, z
            oRevSurface.Capped = False
        End If
        
        'need to check if we need to reverse surface's normal
        'needed if the surface is used as a cutback surface
        oPlane.GetNormal x, y, z
        yVec.Set x, y, z
        

        oSurface.Parameter oCutbackPos.x, oCutbackPos.y, oCutbackPos.z, xPar, yPar
        'surface normal at cutback position
        oSurface.Normal xPar, yPar, x, y, z
        xVec.Set x, y, z
        
        If xVec.Dot(yVec) < 0 Then 'need to reverse otherwise the cutback won't be applied correctly
            Set oRevSurface = oSurface
            oRevSurface.ReverseNormal = True
        End If
        
        
    ElseIf Not oLine Is Nothing Then
        oLine.GetDirection x, y, z
        'todo - increase projection length to handle supported member framing to the end of the curve
        If oSurface Is Nothing Then
            Set oSurface = oGeomFactory.Projections3d.CreateByCurve(pResourceManager, oEdgeLine, x, y, z, oLine.length, False)
        Else
            Dim oProjSurface As IJProjection
            Set oProjSurface = oSurface
            oProjSurface.SetProjection x, y, z
            oProjSurface.length = oLine.length
            oProjSurface.Curve = oEdgeLine
            oProjSurface.Capped = False
        End If
        oPlane.GetNormal x, y, z
        yVec.Set x, y, z
        oLine.GetStartPoint x, y, z
               
        oSurface.Parameter x, y, z, xPar, yPar
        'surface normal at cutback position same as at start position
        oSurface.Normal xPar, yPar, x, y, z
        xVec.Set x, y, z
        
        If xVec.Dot(yVec) < 0 Then 'need to reverse otherwise the cutback won't be applied correctly
            Set oProjSurface = oSurface
            oProjSurface.ReverseNormal = True
        End If
    
    End If
    
    Exit Function

ErrorHandler:
    HandleError MODULE, MT
End Function
'*************************************************************************
'
'<CreateShorterCurve>
'
'Abstract
'
'<Creates a shorter curve from the input curve with the input position as the center>
'
'Arguments
'
'<Input curve, position along the curve, length specified for shorter curve >
'
'Return
'
'<Curve as IJModelBody>
'
'Exceptions
'*************************************************************************
Public Function CreateShorterCurve(oInputCurve As IJCurve, oPosAlong As IJDPosition, ByVal length As Double) As IJDModelBody
    Const MT = "CreateShorterCurve"
    On Error GoTo ErrorHandler
    Dim oGeomFac As New GeometryFactory
    Dim x#, y#, z#, uX#, uY#, uZ#
    Dim oGeometryMisc As IJGeometryMisc
    Dim oShorterCurve As IJCurve
    
    If TypeOf oInputCurve Is IJArc Then
        Dim oArc As IJArc
        Dim dblRadius As Double, dblSide As Double
        Dim oCntrPos As New DPosition
        Dim oSecndPos As IJDPosition, oThrdPos As IJDPosition
        Dim oVec As IJDVector, oNorm As New DVector
        
        
        Set oArc = oInputCurve
        'get the radius of the arc
        dblRadius = oArc.Radius
        'Imagine a right triagle with one corner as the center of the arc, second point as the start of the
        'arc we will create and the third point the mid point of the chord between the ends of the arc we will create
        
        
        'if dblRadius < length/2 then it is an error during dblside compute, so prevent this case
        If dblRadius <= length / 2 Then
            length = dblRadius
        End If
        'let us calculate the length of third point from the center
        dblSide = Sqr(dblRadius * dblRadius - length * length / 4#)
        
        'get center
        oArc.GetCenterPoint x, y, z
        oCntrPos.Set x, y, z
        
        'get vector from center to third point
        Set oVec = oPosAlong.Subtract(oCntrPos)
        oVec.length = dblSide 'set to that of distance from center to third point
        
        'now get the thirdpoint by adding the above vector to the center point
        Set oThrdPos = oCntrPos.Offset(oVec)
        
        'get normal to the arc
        oArc.GetNormal x, y, z
        oNorm.Set x, y, z
        
        'get the vector from third point to second point
        
        Set oVec = oVec.Cross(oNorm)
        oVec.length = length / 2# 'half the chord length of the new arc
        
        Set oSecndPos = oThrdPos.Offset(oVec)
        
        'we will create the arc with the start position, mid position and end position
        'we already have the first two, let us create the end position now
        oVec.[Scale] -1  'revers direction
        
        'we will offset othrdPos in the other direction to get the end position of the arc
        Set oThrdPos = oThrdPos.Offset(oVec)
        
        'now create the arc
        Set oShorterCurve = oGeomFac.Arcs3d.CreateBy3Points(Nothing, oSecndPos.x, oSecndPos.y, oSecndPos.z, _
        oPosAlong.x, oPosAlong.y, oPosAlong.z, oThrdPos.x, oThrdPos.y, oThrdPos.z)
    ElseIf TypeOf oInputCurve Is IJLine Then
        Dim oLine As IJLine
        Set oLine = oInputCurve

        
        oLine.GetDirection uX, uY, uZ
        oPosAlong.Get x, y, z
        
        x = x - 0.5 * length * uX
        y = y - 0.5 * length * uY
        z = z - 0.5 * length * uZ
        
        Set oShorterCurve = oGeomFac.Lines3d.CreateByPtVectLength(Nothing, x, y, z, uX, uY, uZ, length)
    Else ' todo- handle the case of complex string here- currently the caller doesn't send in a complex sting
        
    End If
    
    
    If Not oShorterCurve Is Nothing Then
        Dim oWireBody As IJWireBody
        Set oGeometryMisc = New DGeomOpsMisc
        oGeometryMisc.CreateModelGeometryFromGType Nothing, oShorterCurve, Nothing, oWireBody
        Set CreateShorterCurve = oWireBody
    End If
    
    Exit Function
ErrorHandler:
    HandleError MODULE, MT

End Function

'*************************************************************************
'
'<CreateSolidCutter>
'
'Abstract
'
'<Creates/Modifes solid cutter based on  axis and 2D cutter profile
'if axis and orientation vector are not given that of supporting is used>
'
'Arguments
'
'<Peristenet object manager, Supported member, Supporting Member, Connected end,2D cutter shape,
'Solid Cutter, optional axis, optional orientation vector>
'
'Return
'
'<Solid geometry>
'
'Exceptions
'*************************************************************************
Public Sub CreateSolidCutter(pResourceManager As Object, oSuppingPart As ISPSMemberPartPrismatic, _
oSuppedPart As ISPSMemberPartPrismatic, iEnd As SPSMemberAxisPortIndex, _
oCutterContour As IJComplexString, oCutter As IJDModelBody, Optional oAxis As IJLine, _
Optional oOrientationVec As IJDVector)
    Const MT = "CreateSolidCutter"
    On Error GoTo ErrorHandler
    
    Dim oCurve As IJCurve
    Dim oVec As New DVector
    Dim x#, y#, z#
    Dim sX#, sY#, sZ#, eX#, eY#, eZ#
    Dim oFac As New SPSMembers.SPSMemberFactory
    Dim oFeatureHlper As ISPSMemberFeatureServices
    Dim jobject As IJDObject
    Dim oShortCurve As IJDModelBody
    Dim oPosAlong As IJDPosition
    Dim oArrayOfids() As Long
    Dim numCurves As Long
    Dim oProfile2dhelper As IJDProfile2dHelper
    Dim oGeometryMisc As IJGeometryMisc
    Dim oWireBody As IJWireBody
    Dim idx As Long
    Dim bMirror As Boolean
    Dim oMat As IJDT4x4
    Dim lengthSupped#
    
    Set oGeometryMisc = New DGeomOpsMisc
    Set oFeatureHlper = oFac.CreateMemberFeatureServices
    
    'first get the supported member's intersection position on supporting curve
    Set oPosAlong = GetConnectionPositionOnSupping(oSuppingPart, oSuppedPart, iEnd)
    
    If oOrientationVec Is Nothing Then
        oSuppingPart.Rotation.GetTransformAtPosition oPosAlong.x, oPosAlong.y, oPosAlong.z, oMat, Nothing
        oVec.Set oMat.IndexValue(8), oMat.IndexValue(9), oMat.IndexValue(10)
    Else
        'squared cope situation
        Set oVec = oOrientationVec
    End If
    
    If oAxis Is Nothing Then
        Set oCurve = GetCurve(oSuppingPart, oSuppedPart, iEnd)
        'we need to create another curve, a shorter one, from this so that the cutter we
        'create below doesn't cut the other end of the member. This is an issue when
        ' members we deal with are curved
                
        'get the distance  between ends of supped member axis
        If oSuppedPart.Axis.Scope = CURVE_SCOPE_COLINEAR Then
            Set oShortCurve = oCurve
        Else
            oSuppedPart.Axis.EndPoints sX, sY, sZ, eX, eY, eZ
            lengthSupped = Sqr((sX - eX) * (sX - eX) + (sY - eY) * (sY - eY) + (sZ - eZ) * (sZ - eZ))
    
            'create a curve along the axis of supporting with connection point as mid point
            'and length equal to the distance between Supped axis end points
            'such that the cutter that we create has very high chance of cutting across the entire cross section
            'of the supported member at the same time not cutting the other end of the supped member
            Set oShortCurve = CreateShorterCurve(oCurve, oPosAlong, lengthSupped)
        End If
    Else
        'squared cope situation
        oGeometryMisc.CreateModelGeometryFromGType Nothing, oAxis, Nothing, oShortCurve
    End If
    
    If Not oCutter Is Nothing Then
        If TypeOf oCutter Is IJWireBody Then
            'cutter create in pre v7 model so need to handle it
            Set oWireBody = oCutter
            'need to transform the contour to global
            If oAxis Is Nothing Or oOrientationVec Is Nothing Then
                oSuppingPart.Rotation.GetTransform oMat
                Set oMat = CreateCSToMembTransform(oMat, oSuppingPart.Rotation.Mirror)
            Else
                'squared cope situation
                Dim xVec As New DVector, zVec As New DVector
                oAxis.GetDirection x, y, z
                zVec.Set x, y, z
                Set xVec = oVec.Cross(zVec)
                xVec.length = 1
                Set oMat = New DT4x4
                oMat.LoadIdentity
                oMat.IndexValue(0) = xVec.x
                oMat.IndexValue(1) = xVec.y
                oMat.IndexValue(2) = xVec.z
                oMat.IndexValue(4) = oVec.x
                oMat.IndexValue(5) = oVec.y
                oMat.IndexValue(6) = oVec.z
                oMat.IndexValue(8) = zVec.x
                oMat.IndexValue(9) = zVec.y
                oMat.IndexValue(10) = zVec.z
                oAxis.GetStartPoint x, y, z
                oMat.IndexValue(12) = x
                oMat.IndexValue(13) = y
                oMat.IndexValue(14) = z
            
            End If
            oCutterContour.Transform oMat
            
            oGeometryMisc.ModifyModelGeometryFromGType oCutterContour, oWireBody
        Else
            oGeometryMisc.CreateModelGeometryFromGType Nothing, oCutterContour, Nothing, oWireBody
        End If
    Else
        oGeometryMisc.CreateModelGeometryFromGType Nothing, oCutterContour, Nothing, oWireBody
    End If
    
    'attribute the wirebody
    numCurves = oCutterContour.CurveCount
    ReDim oArrayOfids(numCurves)
    For idx = 0 To numCurves
      oArrayOfids(idx) = idx + 1
    Next idx
    
    Set oProfile2dhelper = New DProfile2dHelper
    oProfile2dhelper.SetIntegerAttributesOnWireBodyEdges oWireBody, "JSXid", numCurves, oArrayOfids(0)
    
    If oAxis Is Nothing Then
        bMirror = oSuppingPart.Rotation.Mirror
        'for cross sections the 2d coord system and the supporting axis doesn't make right hand coord system
        'so the underlying G&T call used by CreateCuttingTool/ModifyCuttingTool call below
        'has mirror set to true by default
        'so when we send mirror=true inside the call it is set to false
    Else
        bMirror = True '
        'so the CreateCuttingTool/ModifyCuttingTool call will set mirror=false in the underlying G&T call
    End If
    If oCutter Is Nothing Then
        'create a solid cutter by projecting the cutting shape along the axis curve
        oFeatureHlper.CreateCuttingTool pResourceManager, oWireBody, oShortCurve, oVec, bMirror, oCutter
    Else 'modify the cutter
        If Not TypeOf oCutter Is IJWireBody Then
            Set jobject = oCutter
            oFeatureHlper.ModifyCuttingTool jobject.ResourceManager, oWireBody, oShortCurve, oVec, bMirror, oCutter
        Else
            'cutter created in pre v7 model. do nothing as we modified it above
        End If
    End If
    Exit Sub
ErrorHandler:
    HandleError MODULE, MT

End Sub

'*************************************************************************
'
'<PostMissingInputsError>
'
'Abstract
'
'< Adds the feature in TDL with missing inputs error. Features were previosuly not modifiable. Now they do, but we have a new codelist msg
'    which may or may not exist in user catalog (if new codelist not bulklaoded. In that case, stick with old msg
' >
'
'Arguments
'
'<Feature Smart Occ >
'
'Return
'
'
'Exceptions
'*************************************************************************

Public Sub PostMissingInputsError(oSmartOcc As IJSmartOccurrence, strSourceFile As String, strError As String)
  Const METHOD = "PostMissingInputsError"
  On Error GoTo ErrorHandler
  Dim oInfosCol As IJDInfosCol
  Dim lMsgKey As Long
  
  'default to old msg
  lMsgKey = TDL_FEATUREMACROS_INVALIDINPUTS_TRIMCUTBACK
  Set oInfosCol = GetCodeList(FeatureToDoMsgCodelist)
  If Not oInfosCol Is Nothing Then
    If oInfosCol.count > 6 Then
        ' assuming user did not add msgs on its own. IN that case we are not in good shape anyway as user
        'might have deleted our msgs. Should not do that.
        'this is new msg number
        lMsgKey = TDL_FEATUREMACROS_INVALIDINPUTS_EDIT_FEATURE
    Else
        lMsgKey = TDL_FEATUREMACROS_INVALIDINPUTS_TRIMCUTBACK
    End If
  End If
  
  SPSToDoErrorNotify FeatureToDoMsgCodelist, lMsgKey, oSmartOcc, Nothing
  Err.Raise SPS_MACRO_WARNING
    
Exit Sub
ErrorHandler:
    Err.Raise ReportError(Err, strSourceFile, METHOD, strError).Number

End Sub

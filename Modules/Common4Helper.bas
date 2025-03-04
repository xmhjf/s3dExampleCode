Attribute VB_Name = "Common4Helper"
Option Explicit
Const g_bPermute As Boolean = True
Private Const MODULE = "Common4Helper:"  'Used for error messages

Public Sub ConnectSmartOccurrence(pSO As IJSmartOccurrence, pRefColl As IJDReferencesCollection)
Dim sSourceFile As String: Dim sMethod As String: Dim sError As String
Let sSourceFile = "Common4Helper.bas": Let sMethod = "ConnectSmartOccurrence"
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
ErrorHandler:
    Err.Raise ReportError(Err, sSourceFile, sMethod, sError).Number
End Sub


Public Function GetRefCollFromSmartOccurrence(pSO As IJSmartOccurrence) As IJDReferencesCollection
Dim sSourceFile As String: Dim sMethod As String: Dim sError As String
Let sSourceFile = "Common4Helper.bas": Let sMethod = "GetRefCollFromSmartOccurrence"
On Error GoTo ErrorHandler
  
   'connect the reference collection to the smart occurrence
    Dim pRelationHelper As IMSRelation.DRelationHelper
    Dim pCollectionHelper As IMSRelation.DCollectionHelper
    Dim pRelationshipHelper As DRelationshipHelper
    Dim pRevision As IJRevision
    
    Set pRelationHelper = pSO
    Set pCollectionHelper = pRelationHelper.CollectionRelations("{A2A655C0-E2F5-11D4-9825-00104BD1CC25}", "toArgs_O")
    If Not pCollectionHelper Is Nothing Then
        If pCollectionHelper.Count = 1 Then
            Set GetRefCollFromSmartOccurrence = pCollectionHelper.Item("RC")
        End If
    End If
  
Exit Function
ErrorHandler:
    Err.Raise ReportError(Err, sSourceFile, sMethod, sError).Number
End Function

'*********************************************************************************
'Get's the default surface from the equipment, if it is defined for the equipment.
'*********************************************************************************
Public Function GetDefaultSurface(oEquip As IJEquipment) As Object
Dim sSourceFile As String: Dim sMethod As String: Dim sError As String
Let sSourceFile = "Common4Helper.bas": Let sMethod = "GetDefaultSurface"
On Error GoTo ErrorHandler

    Dim oSymbol As IJDSymbol
    Dim oISD As IJDSymbolDefinition
    Dim oReps As IJDRepresentations
    Dim oRep As IJDRepresentation
    Dim oDefSurface As IJSurface

    On Error Resume Next
    Set oSymbol = oEquip
    
    If Not oSymbol Is Nothing Then
    Debug.Assert Not oSymbol Is Nothing

    Set oISD = oSymbol.IJDSymbolDefinition(1)
    Debug.Assert Not oISD Is Nothing

    Set oReps = oISD.IJDRepresentations
    Debug.Assert Not oReps Is Nothing

    Set oRep = oReps.GetRepresentationAtIndex(1)
    Debug.Assert Not oRep Is Nothing
    
    Dim strRepName As String
    strRepName = oRep.Name
    
    On Error Resume Next
    Set oDefSurface = oSymbol.BindToOutput(strRepName, "DefaultSurface")
    
    If oDefSurface Is Nothing Then
        Set oDefSurface = oSymbol.BindToOutput(strRepName, "DefaultSurface1")
    End If
    
    On Error GoTo ErrorHandler

    If Not oDefSurface Is Nothing Then
        Set GetDefaultSurface = oDefSurface
    Else
        Set GetDefaultSurface = Nothing
    End If
    End If

    Set oSymbol = Nothing
    Set oISD = Nothing
    Set oReps = Nothing
    Set oRep = Nothing
    Set oDefSurface = Nothing
    
Exit Function
ErrorHandler:
    Err.Raise ReportError(Err, sSourceFile, sMethod, sError).Number

End Function

Public Sub Normalize(vect As DVector)
Dim sSourceFile As String: Dim sMethod As String: Dim sError As String
Let sSourceFile = "Common4Helper.bas": Let sMethod = "Normalize"
On Error GoTo ErrorHandler
    Dim Nx As Double, Ny As Double, Nz As Double
    Dim length As Double
    
    vect.Get Nx, Ny, Nz
    length = Sqr(Nx * Nx + Ny * Ny + Nz * Nz)
    If length > 0 Then
        vect.Set Nx / length, Ny / length, Nz / length
    End If
Exit Sub
ErrorHandler:
    Err.Raise ReportError(Err, sSourceFile, sMethod, sError).Number

End Sub

Public Sub GetCurveProjVectorOnSurface(iCurve As IJCurve, oSurface2 As IJSurface, projVector As IJDVector)
Dim sSourceFile As String: Dim sMethod As String: Dim sError As String
Let sSourceFile = "Common4Helper.bas": Let sMethod = "GetCurveProjVectorOnSurface"
On Error GoTo ErrorHandler
On Error GoTo ErrorHandler

    Dim ptX As Double, ptY As Double, ptZ As Double
    Dim cvX As Double, cvY As Double, cvZ As Double
    Dim dist As Double
    
    'surface have outward normals, we want the cutting tools inward
    'return normal to curve oriented as the reverse normal of surface
    Dim u As Double, v As Double
    Dim Nx As Double, Ny As Double, Nz As Double
    Dim cvNorm As New DVector
    Dim sfNorm As New DVector
    Dim dotProd As Double
    Dim scope As Geom3dCurveScopeConstants
    
    iCurve.DistanceBetween oSurface2, dist, cvX, cvY, cvZ, ptX, ptY, ptZ
    
    iCurve.Normal scope, Nx, Ny, Nz
    cvNorm.Set Nx, Ny, Nz
    oSurface2.Parameter ptX, ptY, ptZ, u, v
    oSurface2.Normal u, v, Nx, Ny, Nz
    sfNorm.Set Nx, Ny, Nz
    
    dotProd = cvNorm.Dot(sfNorm)
    If (dotProd < 0) Then
        projVector.Set cvNorm.x, cvNorm.y, cvNorm.z
    Else
        projVector.Set -cvNorm.x, -cvNorm.y, -cvNorm.z
    End If

'return surf normal inward
projVector.Set -Nx, -Ny, -Nz

Normalize projVector

Exit Sub
ErrorHandler:
    Err.Raise ReportError(Err, sSourceFile, sMethod, sError).Number

End Sub


Public Function GetDistBetweenContourAndSurface(oContour As Object, oSurface2 As IJSurface, projVector As IJDVector) As Double
Dim sSourceFile As String: Dim sMethod As String: Dim sError As String
Let sSourceFile = "Common4Helper.bas": Let sMethod = "GetDistBetweenContourAndSurface"
On Error GoTo ErrorHandler
Dim oWireBody As IJWireBody
On Error GoTo ErrorHandler
GetDistBetweenContourAndSurface = 0#

Dim startPos As IJDPosition
Dim endPos As IJDPosition
Dim startDir As IJDVector
Dim sText As String


If TypeOf oContour Is IJCurve Then
    Dim iCurve As IJCurve
    Dim ptX As Double, ptY As Double, ptZ As Double
    Dim cvX As Double, cvY As Double, cvZ As Double
    Dim dist As Double
    Set iCurve = oContour
    iCurve.DistanceBetween oSurface2, dist, cvX, cvY, cvZ, ptX, ptY, ptZ
    If Not projVector Is Nothing Then
        If (Abs(dist) <= 0.001) Then
            GetCurveProjVectorOnSurface iCurve, oSurface2, projVector
        Else
            projVector.Set ptX - cvX, ptY - cvY, ptZ - cvZ
            Normalize projVector
        End If
    End If
    GetDistBetweenContourAndSurface = dist
End If

Exit Function
ErrorHandler:
    Err.Raise ReportError(Err, sSourceFile, sMethod, sError).Number
End Function

Public Function GetSupportThickness(oConnectable As IJConnectable) As Double
Dim sSourceFile As String: Dim sMethod As String: Dim sError As String
Let sSourceFile = "Common4Helper.bas": Let sMethod = "GetSupportThickness"
On Error GoTo ErrorHandler
GetSupportThickness = 0#

If TypeOf oConnectable Is ISPSWallPart Then
   Dim oWallPart As ISPSWallPart
   Set oWallPart = oConnectable
   GetSupportThickness = oWallPart.Thickness
   Exit Function
ElseIf TypeOf oConnectable Is ISPSSlabEntity Then
   Dim oOperationPattern As IJStructOperationPattern
   Set oOperationPattern = oConnectable
   Dim pGeomParents As IJElements
   Dim pTempAEDisp As Object
   Const sPlanaSlabAEProgId As String = "SPSSlabs.SPSPlanarSlabAE.1"
   
   oOperationPattern.GetOperationPattern sPlanaSlabAEProgId, pGeomParents, pTempAEDisp
   Dim oSlabBoundOperationAE As ISPSSlabBoundOperationAE
   Set oSlabBoundOperationAE = pTempAEDisp
   GetSupportThickness = oSlabBoundOperationAE.Thickness
   Exit Function
Else
    sError = "Invalid Support type"
    Err.Raise ReportError(Err, sSourceFile, sMethod, sError).Number
End If

Exit Function
ErrorHandler:
    Err.Raise ReportError(Err, sSourceFile, sMethod, sError).Number
End Function
Public Function Line_Permute(pLine As IJLine) As IJLine
Dim sSourceFile As String: Dim sMethod As String: Dim sError As String
Let sSourceFile = "Common4Helper.bas": Let sMethod = "Line_Permute"
On Error GoTo ErrorHandler
    If g_bPermute Then
        Dim dX0 As Double, dY0 As Double, dZ0 As Double
        Dim dX1 As Double, dY1 As Double, dZ1 As Double
        Call pLine.GetStartPoint(dX0, dY0, dZ0)
        Call pLine.GetEndPoint(dX1, dY1, dZ1)
        Call pLine.DefineBy2Points(dX0, -dZ0, dY0, _
                                   dX1, -dZ1, dY1)
    End If
    Set Line_Permute = pLine
Exit Function
ErrorHandler:
    Err.Raise ReportError(Err, sSourceFile, sMethod, sError).Number
End Function

Public Sub Vector_Permute(pVector As IJDVector)
 Dim sSourceFile As String: Dim sMethod As String: Dim sError As String
Let sSourceFile = "Common4Helper.bas": Let sMethod = "Vector_Permute"
On Error GoTo ErrorHandler
   If g_bPermute Then
        Dim dU As Double
        Dim dV As Double
        Dim dW As Double
        Call pVector.Get(dU, dV, dW)
        Call pVector.Set(dU, -dW, dV)
    End If
Exit Sub
ErrorHandler:
    Err.Raise ReportError(Err, sSourceFile, sMethod, sError).Number
End Sub

Public Sub Position_Permute(pPosition As IJDPosition)
Dim sSourceFile As String: Dim sMethod As String: Dim sError As String
Let sSourceFile = "Common4Helper.bas": Let sMethod = "Position_Permute"
On Error GoTo ErrorHandler
    If g_bPermute Then
        Dim dX As Double
        Dim dY As Double
        Dim dZ As Double
        Call pPosition.Get(dX, dY, dZ)
        Call pPosition.Set(dX, -dZ, dY)
    End If
Exit Sub
ErrorHandler:
    Err.Raise ReportError(Err, sSourceFile, sMethod, sError).Number
End Sub

Public Function Geometry_Transform(pGeometry As IJDGeometry, pT4x4 As DT4x4) As IJDGeometry
Dim sSourceFile As String: Dim sMethod As String: Dim sError As String
Let sSourceFile = "Common4Helper.bas": Let sMethod = "Geometry_Transform"
On Error GoTo ErrorHandler
    If g_bPermute Then
        Call pGeometry.DTransform(pT4x4)
    End If
    Set Geometry_Transform = pGeometry
Exit Function
ErrorHandler:
    Err.Raise ReportError(Err, sSourceFile, sMethod, sError).Number
End Function

Public Function CurvesAsElements(ParamArray Curves() As Variant) As IJElements
    Const METHOD = "CurvesAsElements:"
    On Error GoTo ErrorHandler
    Dim pElements As IJElements
    Set pElements = New JObjectCollection
    
    Dim lSize As Long
    Let lSize = 0
    
    ' >> Warning: Kludge incomming! <<
    Dim vUgly As Variant
    Set vUgly = Curves(0)   ' This is a bit dirty, isn't it?
    
    Dim vSB As Variant
    For Each vSB In Curves()
        lSize = lSize + 1
    Next vSB
    
    If lSize <= 0 Then GoTo ErrorHandler
    
    ReDim pDoubles(1 To (lSize * 3)) As Double
    Dim i As Long
    For i = LBound(Curves()) To UBound(Curves())
        Dim pCurve As IJCurve
        Set pCurve = Curves(i)
        pElements.Add pCurve
    Next i
    Set CurvesAsElements = pElements
    Exit Function

ErrorHandler:
    ReportUnanticipatedError MODULE, METHOD
End Function
 
Public Function Vector_Scale(pVector As IJDVector, dScale As Double) As IJDVector
Const METHOD = "Vector_Scale"
On Error GoTo ErrorHandler
    Dim pVectorNew As New DVector
    Call pVectorNew.Set(pVector.x * dScale, _
                        pVector.y * dScale, _
                        pVector.z * dScale)

    Set Vector_Scale = pVectorNew
    Exit Function

ErrorHandler:
    ReportUnanticipatedError MODULE, METHOD
End Function

Public Function ProjectSplitComplexString1(ByVal pOutputCollection As IJDOutputCollection, ByVal sOutputName As String, pElementsOfCurves As IJElements, pVectorOfProjection As IJDVector, dLengthOfProjection As Double, pT4x4OfTranslation As DT4x4)
Const METHOD = "ProjectSplitComplexString1"
On Error GoTo ErrorHandler
    Dim pObject As IJDObject
    Set pObject = pOutputCollection
    
    With New GeometryFactory
    
        Dim i As Integer
        For i = 1 To pElementsOfCurves.Count
            Dim pLine As IJLine
            Set pLine = pElementsOfCurves.Item(i)
            
            Dim pPositionOfLowerStartPoint As IJDPosition
            Set pPositionOfLowerStartPoint = Line_GetPosition(pLine, 0)
            
            Dim pPositionOfLowerEndPoint As IJDPosition
            Set pPositionOfLowerEndPoint = Line_GetPosition(pLine, 1)
            
            Dim pPositionOfTopStartPoint As IJDPosition
            Set pPositionOfTopStartPoint = pPositionOfLowerStartPoint.offset(Vector_Scale(pVectorOfProjection, dLengthOfProjection))
            
            Dim pPositionOfTopEndPoint As IJDPosition
            Set pPositionOfTopEndPoint = pPositionOfLowerEndPoint.offset(Vector_Scale(pVectorOfProjection, dLengthOfProjection))
            
            Dim dPoints(1 To 12) As Double
            Dim k As Integer
            k = 1
            dPoints(k) = pPositionOfLowerStartPoint.x: k = k + 1
            dPoints(k) = pPositionOfLowerStartPoint.y: k = k + 1
            dPoints(k) = pPositionOfLowerStartPoint.z: k = k + 1
    
            dPoints(k) = pPositionOfLowerEndPoint.x: k = k + 1
            dPoints(k) = pPositionOfLowerEndPoint.y: k = k + 1
            dPoints(k) = pPositionOfLowerEndPoint.z: k = k + 1
    
            dPoints(k) = pPositionOfTopEndPoint.x: k = k + 1
            dPoints(k) = pPositionOfTopEndPoint.y: k = k + 1
            dPoints(k) = pPositionOfTopEndPoint.z: k = k + 1
            
            dPoints(k) = pPositionOfTopStartPoint.x: k = k + 1
            dPoints(k) = pPositionOfTopStartPoint.y: k = k + 1
            dPoints(k) = pPositionOfTopStartPoint.z: k = k + 1
            
            Dim pPlane3d As IJPlane
            Set pPlane3d = .Planes3d.CreateByPoints(pObject.ResourceManager, 4, dPoints)
    
            Dim sAutoName As String
            Call pOutputCollection.AddOutput(sOutputName, Geometry_Transform(pPlane3d, pT4x4OfTranslation), sAutoName)
            'extract normal from extruded surface
            Dim pSurface As IJSurface
            Set pSurface = pPlane3d
            
            Dim eScope As Geom3dSurfaceScopeConstants
            Dim dU As Double, dV As Double, dW As Double
            Call pSurface.ScopeN(eScope, dU, dV, dW)
'            MsgBox "OutputName = " + sOutputName + " U = " + Str(dU) + " V = " + Str(dV) + " W = " + Str(dW)
        Next
        
    End With
    Exit Function

ErrorHandler:
    ReportUnanticipatedError MODULE, METHOD
End Function

Public Function GetMatrixFromCoordinateSystem(ByVal pCoordinateSystem As IJLocalCoordinateSystem) As IJDT4x4
Const METHOD = "GetMatrixFromCoordinateSystem"
On Error GoTo ErrorHandler
    Dim dMatrix(0 To 15) As Double
    Set GetMatrixFromCoordinateSystem = New DT4x4
        
    dMatrix(0) = pCoordinateSystem.XAxis.x
    dMatrix(1) = pCoordinateSystem.XAxis.y
    dMatrix(2) = pCoordinateSystem.XAxis.z
    dMatrix(3) = 0
        
    dMatrix(4) = pCoordinateSystem.YAxis.x
    dMatrix(5) = pCoordinateSystem.YAxis.y
    dMatrix(6) = pCoordinateSystem.YAxis.z
    dMatrix(7) = 0
        
    dMatrix(8) = pCoordinateSystem.ZAxis.x
    dMatrix(9) = pCoordinateSystem.ZAxis.y
    dMatrix(10) = pCoordinateSystem.ZAxis.z
    dMatrix(11) = 0
        
    dMatrix(12) = pCoordinateSystem.Position.x
    dMatrix(13) = pCoordinateSystem.Position.y
    dMatrix(14) = pCoordinateSystem.Position.z
    dMatrix(15) = 1
        
    GetMatrixFromCoordinateSystem.Set dMatrix(0)
    Exit Function

ErrorHandler:
    ReportUnanticipatedError MODULE, METHOD
End Function

Public Function Line_GetPosition(pLine As IJLine, iIndex As Integer) As IJDPosition
Const METHOD = "Line_GetPosition"
On Error GoTo ErrorHandler
    Dim dX As Double, dY As Double, dZ As Double
    Select Case iIndex
        Case 0:
            Call pLine.GetStartPoint(dX, dY, dZ)
        Case 1:
            Call pLine.GetEndPoint(dX, dY, dZ)
    End Select
    
    Dim pPosition As New DPosition
    Call pPosition.Set(dX, dY, dZ)
    
    Set Line_GetPosition = pPosition
    Exit Function

ErrorHandler:
    ReportUnanticipatedError MODULE, METHOD
End Function

Public Function Type_Dump(lType As Long) As String
    Select Case lType
        Case 1: Let Type_Dump = "Unknown"
        Case 2: Let Type_Dump = "Points"
        Case 3: Let Type_Dump = "Wire"
        Case 4: Let Type_Dump = "Surface"
        Case 5: Let Type_Dump = "Solid"
        Case 6: Let Type_Dump = "Degenerate"
        Case Else: Let Type_Dump = "?"
    End Select
End Function

Public Function Context_Dump(lContext As Long) As String
    Select Case lContext
        Case -1: Let Context_Dump = "JSCTX_INVALID"
        Case 0: Let Context_Dump = "JSCTX_NOP"
        Case 1: Let Context_Dump = "JSCTX_BASE"
        Case 2: Let Context_Dump = "JSCTX_OFFSET"
        Case 4: Let Context_Dump = "JSCTX_LATERAL"
        Case 8: Let Context_Dump = "JSCTX_NPLUS"
        Case 16: Let Context_Dump = "JSCTX_NMINUS"
        Case 32: Let Context_Dump = "JSCTX_INTERNAL_LEDGE"
        Case 33: Let Context_Dump = "JSCTX_BASE_INTERNAL_LEDGE"
        Case 34: Let Context_Dump = "JSCTX_OFFSET_INTERNAL_LEDGE"
        Case 36: Let Context_Dump = "JSCTX_LATERAL_INTERNAL_LEDGE"
        Case 64: Let Context_Dump = "JSCTX_EXTERNAL_LEDGE"
        Case 65: Let Context_Dump = "JSCTX_BASE_EXTERNAL_LEDGE"
        Case 66: Let Context_Dump = "JSCTX_OFFSET_EXTERNAL_LEDGE"
        Case 68: Let Context_Dump = "JSCTX_LATERAL_EXTERNAL_LEDGE"
        Case 128: Let Context_Dump = "JSCTX_LFACE"
        Case 132: Let Context_Dump = "JSCTX_LATERAL_LFACE"
        Case 137: Let Context_Dump = "JSCTX_BASE_NPLUS_LFACE"
        Case 138: Let Context_Dump = "JSCTX_OFFSET_NPLUS_LFACE"
        Case 145: Let Context_Dump = "JSCTX_BASE_NMINUS_LFACE"
        Case 146: Let Context_Dump = "JSCTX_OFFSET_NMINUS_LFACE"
        Case Else: Let Context_Dump = "?"
    End Select
End Function


'From the pattern of the cutout we look if a StructCutoutContour has an input that is our symbol output
Public Sub RemoveCutoutFromContour(oConnectable As IJConnectable, oSblCutoutOutput As Object, sOperationProgId As String)
Dim sSourceFile As String: Dim sMethod As String: Dim sError As String
Let sSourceFile = "SimpleDoor_1_Asm.bas": Let sMethod = "RemoveCutoutFromContour"
On Error GoTo ErrorHandler

Dim OperationPattern As IJStructOperationPattern
Set OperationPattern = oConnectable

'Get cutouts on pattern
Dim oCollectionOfOperators As IJElements
Dim oStructCutoutOperationAE As StructCutoutOperationAE
Dim oStructCutoutContour As StructCutoutContour
OperationPattern.GetOperationPattern sOperationProgId, oCollectionOfOperators, oStructCutoutOperationAE
If (Not oCollectionOfOperators Is Nothing) Then
    For Each oStructCutoutContour In oCollectionOfOperators
        ' If the symbol contour is in the collection remove it
        If oStructCutoutContour.InputContour Is oSblCutoutOutput Then
            Dim colIndex As Long
            
'            colIndex = oCollectionOfOperators.GetIndex(oStructCutoutContour)
'            oCollectionOfOperators.Remove (colIndex)
'            OperationPattern.SetOperationPattern sOperationProgId, oCollectionOfOperators, oStructCutoutOperationAE
            'Remove existing cutout before exit
            Dim ocutoutDel As IJDObject
            Set ocutoutDel = oStructCutoutContour
            ocutoutDel.Remove
            Exit For
        End If
    Next
End If
        
Exit Sub
ErrorHandler:
    Err.Raise ReportError(Err, sSourceFile, sMethod, sError).Number

End Sub

'If the SO has created a cutout, a refcoll is connected that references the opposite face of the wall
'From the pattern of the cutout we look if a StructCutoutContour has an input that is our symbol output
Public Sub RemoveSOCutoutObjects(oRefColl As IJDReferencesCollection, oSblCutoutOutput As Object, sOperationProgId As String)
Dim sSourceFile As String: Dim sMethod As String: Dim sError As String
Let sSourceFile = "SimpleDoor_1_Asm.bas": Let sMethod = "RemoveSOCutoutObjects"
On Error GoTo ErrorHandler

Dim OperationPattern As IJStructOperationPattern
Dim oOppositeDeckPort As IJPort
Dim nbRef As Long

nbRef = oRefColl.IJDEditJDArgument.GetCount
If nbRef < 1 Then
    'a refcoll is connected but no more reference, cutoutContour has already been deleted
    Exit Sub
End If
Set oOppositeDeckPort = oRefColl.IJDEditJDArgument.GetEntityByIndex(1)
If oOppositeDeckPort Is Nothing Then
    sError = "no port connected to refcoll"
    Err.Raise ReportError(Err, sSourceFile, sMethod, sError).Number
End If
Set OperationPattern = oOppositeDeckPort.Connectable
If OperationPattern Is Nothing Then
    sError = "Port connected to refcoll has no Connectable"
    Err.Raise ReportError(Err, sSourceFile, sMethod, sError).Number
End If

'Get cutouts on pattern
Dim oCollectionOfOperators As IJElements
Dim oStructCutoutOperationAE As StructCutoutOperationAE
Dim oStructCutoutContour As StructCutoutContour
OperationPattern.GetOperationPattern sOperationProgId, oCollectionOfOperators, oStructCutoutOperationAE
If (Not oCollectionOfOperators Is Nothing) Then
    For Each oStructCutoutContour In oCollectionOfOperators
        ' If the symbol contour is in the collection remove it
        If oStructCutoutContour.InputContour Is oSblCutoutOutput Then
            Dim colIndex As Long
            
            colIndex = oCollectionOfOperators.GetIndex(oStructCutoutContour)
            oCollectionOfOperators.Remove (colIndex)
            OperationPattern.SetOperationPattern sOperationProgId, oCollectionOfOperators, oStructCutoutOperationAE
            'Remove existing cutout before exit
            Dim ocutoutDel As IJDObject
            Set ocutoutDel = oStructCutoutContour
            ocutoutDel.Remove
            Exit For
        End If
    Next
End If
'delete refcoll connected to Oppositeface is deleting the SO,
'Just remove its reference
If Not oRefColl Is Nothing Then
    oRefColl.IJDEditJDArgument.RemoveAll
End If
        
Exit Sub
ErrorHandler:
    Err.Raise ReportError(Err, sSourceFile, sMethod, sError).Number

End Sub


Public Function CutoutCurve_GetCutoutContour(oCurve As Object) As IJDStructCutoutTool
Dim sSourceFile As String: Dim sMethod As String: Dim sError As String
Let sSourceFile = "SimpleDoor_1_Asm.bas": Let sMethod = "CutoutCurve_GetCutoutContour"
On Error GoTo ErrorHandler
        Dim pAssocRelations As IJDAssocRelation
        Dim pRelationshipCol As IJDRelationshipCol
        Dim pRelationship As IJDRelationship
        
        Set CutoutCurve_GetCutoutContour = Nothing
        
        Set pAssocRelations = oCurve
        Set pRelationshipCol = pAssocRelations.CollectionRelations("IJGeometry", "StructCutoutToolGeomCurve_DEST")
        If Not pRelationshipCol Is Nothing And pRelationshipCol.Count > 0 Then
            Set pRelationship = pRelationshipCol.Item(1)
            Set CutoutCurve_GetCutoutContour = pRelationship.Target
        End If

Exit Function
ErrorHandler:
    Err.Raise ReportError(Err, sSourceFile, sMethod, sError).Number
End Function

Public Function StructPort_GetPortSelector(oPort As IJPort) As IJStructPortSelector
Dim sSourceFile As String: Dim sMethod As String: Dim sError As String
Let sSourceFile = "SimpleDoor_1_Asm.bas": Let sMethod = "StructPort_GetPortSelector"
On Error GoTo ErrorHandler
        Dim pAssocRelations As IJDAssocRelation
        Dim pRelationshipCol As IJDRelationshipCol
        Dim pRelationship As IJDRelationship
        Set StructPort_GetPortSelector = Nothing
        
        Set pAssocRelations = oPort
        Set pRelationshipCol = pAssocRelations.CollectionRelations("IJPort", "StructPortItem_DEST")
        If pRelationshipCol.Count = 0 Then
            Exit Function
        End If
        Set pRelationship = pRelationshipCol.Item(1)
        Set StructPort_GetPortSelector = pRelationship.Target

Exit Function
ErrorHandler:
    Err.Raise ReportError(Err, sSourceFile, sMethod, sError).Number
End Function

Public Function GetConnectableFromCutoutContour(oStructCutoutContour As StructCutoutContour) As IJConnectable
Dim sSourceFile As String: Dim sMethod As String: Dim sError As String
Let sSourceFile = "Doors_1_Asm.bas": Let sMethod = "GetConnectableFromCutoutContour"
On Error GoTo ErrorHandler
        Dim iAssocRelations As IJDAssocRelation
        Dim iRelationshipCol As IJDRelationshipCol
        Dim iRelationship As IJDRelationship
        Dim oCutoutAE As Object
        Dim iStructGeometryToEntity As IJStructGeometryToEntity
        Dim iStructCompactGraph As IJStructCompactGraph
        
        Set GetConnectableFromCutoutContour = Nothing
        'Get CutoutOperation
        Set iAssocRelations = oStructCutoutContour
        Set iRelationshipCol = iAssocRelations.CollectionRelations("IJGeometry", "StructCutoutOperators_DEST")
            
        If iRelationshipCol.Count = 0 Then
            'No connectable
            Exit Function
        End If
        
        Set iRelationship = iRelationshipCol.Item(1)
        
        Set oCutoutAE = iRelationship.Target
        
        'Now ask for IJStructCompactGraph to see if we are on the geometry (AE=geom for walls)
        If TypeOf oCutoutAE Is IJStructCompactGraph Then
            Set iStructGeometryToEntity = oCutoutAE
        Else
            'Slab (not compact graph) get the geometry through result
            Set iAssocRelations = oCutoutAE
            Set iRelationshipCol = iAssocRelations.CollectionRelations("IJStructOperationAE", "StructResult_ORIG")
            If iRelationshipCol.Count = 0 Then
                Exit Function
            End If
            Set iRelationship = iRelationshipCol.Item(1)
            
            Set iStructGeometryToEntity = iRelationship.Target
        End If
        
        'Now get the connectable from the geometry
        Set iAssocRelations = iStructGeometryToEntity
        Set iRelationshipCol = iAssocRelations.CollectionRelations("IJGeometry", "StructEntityGeometry_DEST")
        If iRelationshipCol.Count = 0 Then
            Exit Function
        End If
        Set iRelationship = iRelationshipCol.Item(1)
        Set GetConnectableFromCutoutContour = iRelationship.Target
        
Exit Function
ErrorHandler:
    Err.Raise ReportError(Err, sSourceFile, sMethod, sError).Number
End Function

Public Function GetDoorCutoutContour(oSblCutoutOutput As Object) As StructCutoutContour
Dim sSourceFile As String: Dim sMethod As String: Dim sError As String
Let sSourceFile = "SimpleDoor_1_Asm.bas": Let sMethod = "GetDoorCutoutContour"
On Error GoTo ErrorHandler
        Dim pAssocRelations As IJDAssocRelation
        Dim pRelationshipCol As IJDRelationshipCol
        Dim pRelationship As IJDRelationship
        
        Set pAssocRelations = oSblCutoutOutput
        Set pRelationshipCol = pAssocRelations.CollectionRelations("IJGeometry", "StructCutoutToolGeomCurve_DEST")
        
        If Not (pRelationshipCol Is Nothing) And pRelationshipCol.Count > 0 Then
            Set pRelationship = pRelationshipCol.Item(1)
            Set GetDoorCutoutContour = pRelationship.Target
          Else
            Set GetDoorCutoutContour = Nothing
        End If
Exit Function
ErrorHandler:
    Err.Raise ReportError(Err, sSourceFile, sMethod, sError).Number
   
End Function

Public Sub GetOppositeFaceFromSupport(oPort As IJPort, oOppositeDeckPort As IJPort, oSupportType As SupportType)
Dim sSourceFile As String: Dim sMethod As String: Dim sError As String
Let sSourceFile = "SimpleDoor_1_Asm.bas": Let sMethod = "GetOppositeFaceFromSupport"
On Error GoTo ErrorHandler

'Decode given oPort ( if NPLUS, ask for NMINUS ...)
Dim oGraphConnectable As IJStructGraphConnectable
Dim oPortElts As IJElements
Dim oPortEltMk As New MonikeredElement
Dim oStructPort As IJStructPort
Dim lType As Long
Dim lContext As Long
Dim lContextOpposite As Long
Dim lOperation As Long
Dim lOperationOpposite As Long
Dim lOperator As Long
Dim lOperatorOpposite As Long
Dim lXid As Long
Dim sText As String
Dim vOperation As Variant
Dim geomSelector As StructGeometrySelector

Set oOppositeDeckPort = Nothing
Set oStructPort = oPort
Set oGraphConnectable = oPort.Connectable

'Retrieve Port moniker to decode
Dim oPortHelper As IPortHelper
Dim pMonikerOfPort As IMoniker
If Not oStructPort Is Nothing Then
    Set pMonikerOfPort = oStructPort.PortMoniker
Else
    sError = "Error, oStructPort not implemented by support! "
    Err.Raise ReportError(Err, sSourceFile, sMethod, sError).Number
End If

If pMonikerOfPort Is Nothing Then
    sError = "Error, pMonikerOfPort not found ! "
    Err.Raise ReportError(Err, sSourceFile, sMethod, sError).Number
End If

'Decode Port momiker
Set oPortHelper = New PortHelper
Call oPortHelper.DecodeTopologyProxyMoniker(pMonikerOfPort, lType, lContext, lOperation, lOperator, lXid)
Let sText = " port: typ= " + CStr(lType) + " [" + Type_Dump(lType) + "]" + _
                    " ctx= " + CStr(lContext) + " [" + Context_Dump(lContext) + "]" + _
                    " opn= " + CStr(lOperation) + _
                    " opr= " + CStr(lOperator) + _
                    " xid= " + Str(lXid)

'For Slab, support surface can only be JSCTX_BASE_NPLUS_LFACE(137),JSCTX_OFFSET_NPLUS_LFACE (138)
'JSCTX_BASE_NMINUS_LFACE(145) ,JSCTX_OFFSET_NMINUS_LFACE(146) (should be filtered by command ?)
'For walls, it will be on lateral faces (JSCTX_LATERAL_LFACE=132), and operators JXSEC_LEFT_WEB(257)/JXSEC_RIGHT_WEB(258)

lOperatorOpposite = lOperator
lOperationOpposite = lOperation
lContextOpposite = lContext

'Get opposite face by encoding the moniker
Dim sCutoutOperationProgId As String
If oSupportType = SLAB Then
    sCutoutOperationProgId = "StructGeneric.StructCutoutOperationAE.1"
    If Not (lContext = 137 Or lContext = 138 Or lContext = 145 Or lContext = 146) Then
        sError = "Error, support face is not valid (should be a lateral face)"
        Err.Raise ReportError(Err, sSourceFile, sMethod, sError).Number
    End If
    
    Select Case lContext
            Case 137: Let lContextOpposite = 146
            Case 138: Let lContextOpposite = 145
            Case 145: Let lContextOpposite = 138
            Case 146: Let lContextOpposite = 137
            Case Else: Let lContextOpposite = 0
    End Select
Else 'Wall
    sCutoutOperationProgId = "SP3DStructGeneric.StructCutoutOperation"
    If Not (lContext = 132) Then
        sError = "Error, support face is not valid (should be a lateral face)"
        Err.Raise ReportError(Err, sSourceFile, sMethod, sError).Number
    End If
    If Not (lOperator = 257 Or lOperator = 258) Then
        sError = "Error, support face is not valid (should be a lateral face)"
        Err.Raise ReportError(Err, sSourceFile, sMethod, sError).Number
    End If
    Select Case lOperator
            Case 257: Let lOperatorOpposite = 258
            Case 258: Let lOperatorOpposite = 257
            Case Else: Let lOperatorOpposite = lOperator
    End Select
End If


'Find opposite port on the geometry before cutout
'Retrieve portSelector from this port
Dim pStructPortSelector As IJStructPortSelector
Dim pGeometry As IJStructGeometryToEntity

Set pStructPortSelector = StructPort_GetPortSelector(oPort)
geomSelector = pStructPortSelector.geometrySelector
vOperation = "IJStructCutoutOperationAE"
If geomSelector = GeometryInGraph Then vOperation = sCutoutOperationProgId

oGraphConnectable.FindPortsByMonikerIdEx oPortElts, lType, lOperationOpposite, lOperatorOpposite, lContextOpposite, lXid, geomSelector, cmnstrBefore, vOperation
If oPortElts Is Nothing Then
    sError = "Error, FindPortsByMonikerIdEx fails ! "
    Err.Raise ReportError(Err, sSourceFile, sMethod, sError).Number
End If

 If Not oPortElts Is Nothing And oPortElts.Count > 0 Then
     If Not oPortElts.Item(1) Is Nothing Then
         Set oOppositeDeckPort = oPortElts.Item(1)
     Else
        sError = "Error: oPortElts.Item(1) Is Nothing "
        Err.Raise ReportError(Err, sSourceFile, sMethod, sError).Number
     End If
     oPortElts.Clear
 Else
     sError = "Error: No oPortElts found"
     Err.Raise ReportError(Err, sSourceFile, sMethod, sError).Number
 End If

If oOppositeDeckPort Is Nothing Then
     sError = "Error, oOppositeDeckPort not found "
     Err.Raise ReportError(Err, sSourceFile, sMethod, sError).Number
End If

Exit Sub
ErrorHandler:
    Err.Raise ReportError(Err, sSourceFile, sMethod, sError).Number

End Sub


'   Fix TR 107247 : do not use EquipRlnHelper to retrieve relations because it uses IMSElements
'   that needs an active connection (client collection).
Public Function GetDoorMatingSurfaces(oSmartOcc As IJSmartOccurrence, MateOffset As Double) As IJPort
Dim sSourceFile As String: Dim sMethod As String: Dim sError As String
Let sSourceFile = "SimpleDoor_1_Asm.bas": Let sMethod = "GetDoorMatingSurfaces"
On Error GoTo ErrorHandler

Dim oEquip As IJEquipment
Dim oSblDefaultSurface As Object

Set GetDoorMatingSurfaces = Nothing

Set oEquip = oSmartOcc
Set oSblDefaultSurface = GetDefaultSurface(oEquip)
If oSblDefaultSurface Is Nothing Then
    sError = "GetDoorMatingSurfaces::oSblDefaultSurface is nothing"
    Err.Raise ReportError(Err, sSourceFile, sMethod, sError).Number
End If
'The equipment should have a specific face named "DefaultSurface" or "DefaultSurface1" defining the default
'surface for the mate constraint. We will take this surface to recognize the constraint for the opening.


'Look at all mating constraint to see if one is defined with the default surface
Dim activeEntity As IJAssemblyConstraintAE
Dim nbrel As Long
Dim Index As Long

Dim pAssocRelations As IJDAssocRelation
Dim pRelationshipCol As IJDRelationshipCol
Dim pRelationship As IJDRelationship

Set pAssocRelations = oSblDefaultSurface
Set pRelationshipCol = pAssocRelations.CollectionRelations("IJGeometry", "pattern3MatePartSurfAE")

nbrel = pRelationshipCol.Count

For Index = 1 To nbrel
    Dim deckSurface As Object 'Mating surface on support
    Dim relAE As Object
    
    Set pRelationship = pRelationshipCol.Item(Index)
    Set relAE = pRelationship.Target
    If TypeOf relAE Is IJAssemblyConstraintAE Then
        Set activeEntity = relAE
        Dim reltype As Long '0 mate, 1 align
        Dim offset As Double
        reltype = activeEntity.RelationType
        
        'Get mating surfaces
        If reltype = 0 Then
            Dim pAssocRel As IJDAssocRelation
            Dim pRelCol As IJDRelationshipCol
            
            Set pAssocRel = activeEntity
            Set pRelCol = pAssocRel.CollectionRelations("IJAssemblyConstraintAE", "pattern3Surf")
            If pRelCol.Count = 1 Then
                Set deckSurface = pRelCol.Item(1).Target
            End If
            Set pRelCol = Nothing
            If Not deckSurface Is Nothing Then
                'We got opening constraint
                If TypeOf deckSurface Is IJPort Then
                    Set GetDoorMatingSurfaces = deckSurface
                    MateOffset = activeEntity.offset
                End If
                Exit For  ' Dont't care about other relations
            End If
        End If
    End If
Next
    
    
Exit Function
ErrorHandler:
    Err.Raise ReportError(Err, sSourceFile, sMethod, sError).Number

End Function

'  New function to replace GetDoorMatingSurfaces kept for compatibility
Public Function GetDoorMatingSurfaces2(oSmartOcc As IJSmartOccurrence, outAE As IJAssemblyConstraintAE) As IJPort
Dim sSourceFile As String: Dim sMethod As String: Dim sError As String
Let sSourceFile = "SimpleDoor_1_Asm.bas": Let sMethod = "GetDoorMatingSurfaces2"
On Error GoTo ErrorHandler

Dim oEquip As IJEquipment
Dim oSblDefaultSurface As Object

Set GetDoorMatingSurfaces2 = Nothing
Set outAE = Nothing

Set oEquip = oSmartOcc
Set oSblDefaultSurface = GetDefaultSurface(oEquip)
If oSblDefaultSurface Is Nothing Then
    sError = "GetDoorMatingSurfaces::oSblDefaultSurface is nothing"
    Err.Raise ReportError(Err, sSourceFile, sMethod, sError).Number
End If
'The equipment should have a specific face named "DefaultSurface" or "DefaultSurface1" defining the default
'surface for the mate constraint. We will take this surface to recognize the constraint for the opening.


'Look at all mating constraint to see if one is defined with the default surface
Dim activeEntity As IJAssemblyConstraintAE
Dim nbrel As Long
Dim Index As Long

Dim pAssocRelations As IJDAssocRelation
Dim pRelationshipCol As IJDRelationshipCol
Dim pRelationship As IJDRelationship

Set pAssocRelations = oSblDefaultSurface
Set pRelationshipCol = pAssocRelations.CollectionRelations("IJGeometry", "pattern3MatePartSurfAE")

nbrel = pRelationshipCol.Count

For Index = 1 To nbrel
    Dim deckSurface As Object 'Mating surface on support
    Dim relAE As Object
    
    Set pRelationship = pRelationshipCol.Item(Index)
    Set relAE = pRelationship.Target
    If TypeOf relAE Is IJAssemblyConstraintAE Then
        Set activeEntity = relAE
        Dim reltype As Long '0 mate, 1 align
        Dim offset As Double
        reltype = activeEntity.RelationType
        
        'Get mating surfaces
        If reltype = 0 Then
            Dim pAssocRel As IJDAssocRelation
            Dim pRelCol As IJDRelationshipCol
            
            Set pAssocRel = activeEntity
            Set pRelCol = pAssocRel.CollectionRelations("IJAssemblyConstraintAE", "pattern3Surf")
            If pRelCol.Count = 1 Then
                Set deckSurface = pRelCol.Item(1).Target
            End If
            Set pRelCol = Nothing
            If Not deckSurface Is Nothing Then
                'We got opening constraint
                If TypeOf deckSurface Is IJPort Then
                    Set GetDoorMatingSurfaces2 = deckSurface
                    Set outAE = activeEntity
                End If
                Exit For  ' Dont't care about other relations
            End If
        End If
    End If
Next
    
    
Exit Function
ErrorHandler:
    Err.Raise ReportError(Err, sSourceFile, sMethod, sError).Number

End Function

Public Function GetAttribute(pObject As IJDObject, strInterfaceName As String, strAttributeName As String) As IJDAttribute
Dim sSourceFile As String: Dim sMethod As String: Dim sError As String
Let sSourceFile = "SimpleDoor_1_Asm.bas": Let sMethod = "GetAttribute"
On Error GoTo ErrorHandler

    ' inialize access to metadata
    Dim pAttributeMetaData As IJDAttributeMetaData
    Set pAttributeMetaData = pObject.ResourceManager
    
    Dim sInterfaceID As String
    Let sInterfaceID = pAttributeMetaData.IID(strInterfaceName)
    
    Dim pSmartOccurrence As IJSmartOccurrence
    Set pSmartOccurrence = pObject

    ' retrieve the attibute
    Dim pAttributes As IJDAttributes
    Dim pAttributesCol As IJDAttributesCol
    
    Set pAttributes = pSmartOccurrence
    On Error Resume Next
    Set pAttributesCol = pAttributes.CollectionOfAttributes(pAttributeMetaData.IID(strInterfaceName))
    On Error GoTo 0
    
    If Not pAttributesCol Is Nothing Then
        Set GetAttribute = pAttributesCol.Item(strAttributeName)
    End If
Exit Function
ErrorHandler:
    Err.Raise ReportError(Err, sSourceFile, sMethod, sError).Number
    
End Function

' This Sub creates output for Doors/Windows Operation aspect.
' It is called by both SP3DDoorsAsm and SimpleDoorAsm projects.
Sub RunForDoorsWindowsOperationAspect(ByVal pOutputCollection As IJDOutputCollection, ByRef arrayOfInputs(), arrayOfOutputs() As String, ByVal isSimpleDoor As Boolean)
        Const METHOD = "RunForDoorsWindowsOperationAspect:"
        On Error GoTo ErrorHandler
        Dim pObjectOfOutputCollection As IJDObject
        Set pObjectOfOutputCollection = pOutputCollection

        Dim pGeometryFactory As GeometryFactory
        Set pGeometryFactory = New GeometryFactory
        With pGeometryFactory

            'input taken from catalog data

            Dim Kinematics As Long
            Dim lPush As Long

            '    Dim DoorAxisPosition As Long

            Dim dDoorXPosition As Double
            Dim dDoorYPosition As Double
            Dim dDoorZPosition As Double

            Dim dOpeningRatio As Double
            Dim dWidth As Double
            Dim dHeight As Double

            Dim CTLength As Double
            Dim CTwidth As Double
            Dim CTThickness As Double
            Dim CTEdge As Double

            Dim CLLength As Double
            Dim CLwidth As Double
            Dim CLThickness As Double
            Dim CLEdge As Double

            Dim CRLength As Double
            Dim CRwidth As Double
            Dim CRThickness As Double
            Dim CREdge As Double

            Dim CBLength As Double
            Dim CBwidth As Double
            Dim CBThickness As Double
            Dim CBEdge As Double

            
            Dim pT4x4OfTranslation As New DT4x4
            Dim pVectorOfTranslation As New DVector
            

            If isSimpleDoor = False Then    'Door construction for SP3DDoorsAsm
            
                dOpeningRatio = arrayOfInputs(3) ' 0
                Kinematics = arrayOfInputs(4) ' 0
                lPush = arrayOfInputs(5) ' 0
                dHeight = arrayOfInputs(6) ' 800
                dWidth = arrayOfInputs(7) '400
    
                CTLength = arrayOfInputs(8) '400
                CTwidth = arrayOfInputs(9) '150
                CTThickness = arrayOfInputs(10) '150
                CTEdge = arrayOfInputs(11)
    
                CBLength = arrayOfInputs(12) ' 400
                CBwidth = arrayOfInputs(13) ' 50
                CBThickness = arrayOfInputs(14) ' 10
                CBEdge = arrayOfInputs(15)
    
                CLLength = arrayOfInputs(16) ' 800
                CLwidth = arrayOfInputs(17) ' 50
                CLThickness = arrayOfInputs(18) '10
                CLEdge = arrayOfInputs(19)
    
                CRLength = arrayOfInputs(20) ' 800
                CRwidth = arrayOfInputs(21) ' 50
                CRThickness = arrayOfInputs(22) '10
                CREdge = arrayOfInputs(23)
    
                Dim dPannelThickness As Double
    
                dPannelThickness = arrayOfInputs(24) ' 2
    
                dDoorXPosition = arrayOfInputs(25) '0
                dDoorYPosition = arrayOfInputs(27) 'dHeight / 2
                dDoorZPosition = -arrayOfInputs(26)

                ' make the left bottom corner of the mating surface coincident with the origin of the coordinate system
                Call pVectorOfTranslation.Set(-(0 - dWidth / 2 - CLThickness - CLEdge), -(0 - dHeight / 2 - CBThickness - CBEdge), -(0))
                Call Vector_Permute(pVectorOfTranslation)
                Call pT4x4OfTranslation.LoadIdentity
                Call pT4x4OfTranslation.Translate(pVectorOfTranslation)

            Else                        'Door construction for SimpleDoorAsm
                Dim TopFrameLength As Double
                Dim TopFrameDepth As Double
                Dim TopFrameWidth As Double
                Dim LowerFrameLength As Double
                Dim LowerFrameDepth As Double
                Dim LowerFrameWidth As Double
                Dim LeftFrameLength As Double
                Dim LeftFrameDepth As Double
                Dim LeftFrameWidth As Double
                Dim RightFrameLength As Double
                Dim RightFrameDepth As Double
                Dim RightFrameWidth As Double
    
                Let dOpeningRatio = arrayOfInputs(3) ' 0
                Let Kinematics = arrayOfInputs(4) ' 0
                Let lPush = arrayOfInputs(5) ' 0
                Let dHeight = arrayOfInputs(6) ' 800
                Let dWidth = arrayOfInputs(7) '400
                
                Let TopFrameLength = arrayOfInputs(8) '400
                Let TopFrameDepth = arrayOfInputs(9) '150
                Let TopFrameWidth = arrayOfInputs(10) '150
                
                Let LowerFrameLength = arrayOfInputs(11) ' 400
                Let LowerFrameDepth = arrayOfInputs(12) ' 50
                Let LowerFrameWidth = arrayOfInputs(13) ' 10
                
                Let LeftFrameLength = arrayOfInputs(14) ' 800
                Let LeftFrameDepth = arrayOfInputs(15) ' 50
                Let LeftFrameWidth = arrayOfInputs(16) '10
                
                Let RightFrameLength = arrayOfInputs(17) ' 800
                Let RightFrameDepth = arrayOfInputs(18) ' 50
                Let RightFrameWidth = arrayOfInputs(19) '10
                
                Dim dPanelThickness As Double
                
                Let dPanelThickness = arrayOfInputs(20) ' 2
                
                Let dDoorXPosition = arrayOfInputs(21) '0
                Let dDoorYPosition = arrayOfInputs(23) 'dHeight / 2
                Let dDoorZPosition = -arrayOfInputs(22)
                
                ' make the left Lower corner of the mating surface coincident with the origin of the coordinate system
                Call pVectorOfTranslation.Set(-(0 - dWidth / 2), -(0 - dHeight / 2), -(0))
                Call Vector_Permute(pVectorOfTranslation)
                Call pT4x4OfTranslation.LoadIdentity
                Call pT4x4OfTranslation.Translate(pVectorOfTranslation)
    
            End If
            
            Dim iOutput As Integer
            iOutput = 0

            'Cutout wireframe
            iOutput = iOutput + 1
            

            ' The coordinate system was initially at the center of the door (pos at width/2 height/2), it has then been translated to the
            ' Lower left corner.
            '       Call pOutputCollection.AddOutput(arrayOfOutputs(iOutput), _
            '                .ComplexStrings3d.CreateByCurves(pObjectOfOutputCollection.ResourceManager, _
            '                     CurvesAsElements( _
            '                       .Lines3d.CreateBy2Points(Nothing, _
            '1                           dDoorXPosition - dWidth / 2, dDoorYPosition - dHeight / 2, dDoorZPosition, _
            '2                           dDoorXPosition + dWidth / 2, dDoorYPosition - dHeight / 2, dDoorZPosition), _
            '                         .Lines3d.CreateBy2Points(Nothing, _
            '3                           dDoorXPosition + dWidth / 2, dDoorYPosition - dHeight / 2, dDoorZPosition, _
            '4                           dDoorXPosition + dWidth / 2, dDoorYPosition + dHeight / 2, dDoorZPosition), _
            '                         .Lines3d.CreateBy2Points(Nothing, _
            '5                           dDoorXPosition + dWidth / 2, dDoorYPosition + dHeight / 2, dDoorZPosition, _
            '6                           dDoorXPosition - dWidth / 2, dDoorYPosition + dHeight / 2, dDoorZPosition), _
            '                        .Lines3d.CreateBy2Points(Nothing, _
            '7                           dDoorXPosition - dWidth / 2, dDoorYPosition + dHeight / 2, dDoorZPosition, _
            '8                           dDoorXPosition - dWidth / 2, dDoorYPosition - dHeight / 2, dDoorZPosition))))

            ' reverse the order of the points 8-7, 6-5, 4-3, 2-1
            Call pOutputCollection.AddOutput(arrayOfOutputs(iOutput), _
                Geometry_Transform( _
                    .ComplexStrings3d.CreateByCurves(pObjectOfOutputCollection.ResourceManager, _
                        CurvesAsElements( _
                            Line_Permute( _
                                .Lines3d.CreateBy2Points(Nothing, _
                                    dDoorXPosition - dWidth / 2, dDoorYPosition - dHeight / 2, dDoorZPosition, _
                                    dDoorXPosition - dWidth / 2, dDoorYPosition + dHeight / 2, dDoorZPosition)), _
                            Line_Permute( _
                                .Lines3d.CreateBy2Points(Nothing, _
                                    dDoorXPosition - dWidth / 2, dDoorYPosition + dHeight / 2, dDoorZPosition, _
                                    dDoorXPosition + dWidth / 2, dDoorYPosition + dHeight / 2, dDoorZPosition)), _
                            Line_Permute( _
                                .Lines3d.CreateBy2Points(Nothing, _
                                    dDoorXPosition + dWidth / 2, dDoorYPosition + dHeight / 2, dDoorZPosition, _
                                    dDoorXPosition + dWidth / 2, dDoorYPosition - dHeight / 2, dDoorZPosition)), _
                            Line_Permute( _
                                .Lines3d.CreateBy2Points(Nothing, _
                                    dDoorXPosition + dWidth / 2, dDoorYPosition - dHeight / 2, dDoorZPosition, _
                                    dDoorXPosition - dWidth / 2, dDoorYPosition - dHeight / 2, dDoorZPosition)))), pT4x4OfTranslation))

            ' OperationalEnvelope1 output (Marc Fournier 04-Jun-2010, v2011)
            ' Output computed at the center of the door (pos at width/2 height/2) then translated
            '2: 'Swing along vertical left axis
            '3: 'Swing along vertical right axis
            '7: 'Double swing
            
            If ((lPush = 1) Or (lPush = -1)) And ((Kinematics = 2) Or (Kinematics = 3) Or (Kinematics = 7)) Then

                Dim ComplexString(2) As ComplexString3d
                iOutput = iOutput + 1

                If (Kinematics = 2) Or (Kinematics = 3) Then    'Single Swinging Door/Window.
                                                                '     CP         SP
                    Dim CP(3) As Double ' Swinging corner point         .-------.
                    Dim SP(3) As Double ' Opposite corner point         |
                    Dim EP(3) As Double ' 90 degrees point              |
                                                                   'EP  .

                    If Kinematics = 2 Then
                        CP(1) = dDoorXPosition - dWidth / 2
                        CP(2) = dDoorYPosition - dHeight / 2
                        CP(3) = dDoorZPosition
                        SP(1) = dDoorXPosition + dWidth / 2
                        SP(2) = dDoorYPosition - dHeight / 2
                        SP(3) = dDoorZPosition
                        EP(1) = dDoorXPosition - dWidth / 2
                        EP(2) = dDoorYPosition - dHeight / 2
                        EP(3) = dDoorZPosition - lPush * dWidth
                    ElseIf Kinematics = 3 Then
                        CP(1) = dDoorXPosition + dWidth / 2
                        CP(2) = dDoorYPosition - dHeight / 2
                        CP(3) = dDoorZPosition
                        SP(1) = dDoorXPosition - dWidth / 2
                        SP(2) = dDoorYPosition - dHeight / 2
                        SP(3) = dDoorZPosition
                        EP(1) = dDoorXPosition + dWidth / 2
                        EP(2) = dDoorYPosition - dHeight / 2
                        EP(3) = dDoorZPosition + lPush * dWidth
                    End If

                    ' Lower curve (2 lines + the arc at lower elevation).
                    Set ComplexString(1) = .ComplexStrings3d.CreateByCurves(Nothing, _
                                                CurvesAsElements( _
                                                    .Lines3d.CreateBy2Points(Nothing, SP(1), -SP(3), SP(2), CP(1), -CP(3), CP(2)), _
                                                    .Lines3d.CreateBy2Points(Nothing, CP(1), -CP(3), CP(2), EP(1), -EP(3), EP(2)), _
                                                    .Arcs3d.CreateByCenterStartEnd(Nothing, CP(1), -CP(3), CP(2), EP(1), -EP(3), EP(2), SP(1), -SP(3), SP(2))))
                    ' Upper curve (2 lines + the arc at upper elevation).
                    Set ComplexString(2) = .ComplexStrings3d.CreateByCurves(Nothing, _
                                                CurvesAsElements( _
                                                    .Lines3d.CreateBy2Points(Nothing, SP(1), -SP(3), SP(2) + dHeight, CP(1), -CP(3), CP(2) + dHeight), _
                                                    .Lines3d.CreateBy2Points(Nothing, CP(1), -CP(3), CP(2) + dHeight, EP(1), -EP(3), EP(2) + dHeight), _
                                                    .Arcs3d.CreateByCenterStartEnd(Nothing, CP(1), -CP(3), CP(2) + dHeight, EP(1), -EP(3), EP(2) + dHeight, SP(1), -SP(3), SP(2) + dHeight)))

                ElseIf (Kinematics = 7) Then                    'Double Swinging Door/Window.

                    ' 1st Corner point
                    Dim CP1(3) As Double
                    CP1(1) = dDoorXPosition - dWidth / 2
                    CP1(2) = dDoorYPosition - dHeight / 2
                    CP1(3) = dDoorZPosition
                                                            '       CP1       MP      CP2
                    ' Middle point                                      '.-----.-----.
                    Dim MP(3) As Double                                 '|           |
                    MP(1) = dDoorXPosition                              '|           |
                    MP(2) = dDoorYPosition - dHeight / 2           'EP1  .           . EP2
                    MP(3) = dDoorZPosition

                    ' 1st End point
                    Dim EP1(3) As Double
                    EP1(1) = dDoorXPosition - dWidth / 2
                    EP1(2) = dDoorYPosition - dHeight / 2
                    EP1(3) = dDoorZPosition - lPush * dWidth / 2

                    ' 2nd Corner point
                    Dim CP2(3) As Double
                    CP2(1) = dDoorXPosition + dWidth / 2
                    CP2(2) = dDoorYPosition - dHeight / 2
                    CP2(3) = dDoorZPosition

                    ' 2nd End point
                    Dim EP2(3) As Double
                    EP2(1) = dDoorXPosition + dWidth / 2
                    EP2(2) = dDoorYPosition - dHeight / 2
                    EP2(3) = dDoorZPosition - lPush * dWidth / 2

                    ' Lower curve (3 lines + 2 arcs at lower elevation).
                    Set ComplexString(1) = .ComplexStrings3d.CreateByCurves(Nothing, _
                                                CurvesAsElements( _
                                                    .Lines3d.CreateBy2Points(Nothing, CP1(1), -CP1(3), CP1(2), EP1(1), -EP1(3), EP1(2)), _
                                                    .Arcs3d.CreateByCenterStartEnd(Nothing, CP1(1), -CP1(3), CP1(2), EP1(1), -EP1(3), EP1(2), MP(1), -MP(3), MP(2)), _
                                                    .Arcs3d.CreateByCenterStartEnd(Nothing, CP2(1), -CP2(3), CP2(2), MP(1), -MP(3), MP(2), EP2(1), -EP2(3), EP2(2)), _
                                                    .Lines3d.CreateBy2Points(Nothing, EP2(1), -EP2(3), EP2(2), CP2(1), -CP2(3), CP2(2)), _
                                                    .Lines3d.CreateBy2Points(Nothing, CP2(1), -CP2(3), CP2(2), CP1(1), -CP1(3), CP1(2))))
                    ' Upper curve (3 lines + 2 arcs at upper elevation).
                    Set ComplexString(2) = .ComplexStrings3d.CreateByCurves(Nothing, _
                                                CurvesAsElements( _
                                                    .Lines3d.CreateBy2Points(Nothing, CP1(1), -CP1(3), CP1(2) + dHeight, EP1(1), -EP1(3), EP1(2) + dHeight), _
                                                    .Arcs3d.CreateByCenterStartEnd(Nothing, CP1(1), -CP1(3), CP1(2) + dHeight, EP1(1), -EP1(3), EP1(2) + dHeight, MP(1), -MP(3), MP(2) + dHeight), _
                                                    .Arcs3d.CreateByCenterStartEnd(Nothing, CP2(1), -CP2(3), CP2(2) + dHeight, MP(1), -MP(3), MP(2) + dHeight, EP2(1), -EP2(3), EP2(2) + dHeight), _
                                                    .Lines3d.CreateBy2Points(Nothing, EP2(1), -EP2(3), EP2(2) + dHeight, CP2(1), -CP2(3), CP2(2) + dHeight), _
                                                    .Lines3d.CreateBy2Points(Nothing, CP2(1), -CP2(3), CP2(2) + dHeight, CP1(1), -CP1(3), CP1(2) + dHeight)))
                End If

                ' Create OperationalEnvelope1 output, a ruled surface from lower to upper curve (capped).
                Call pOutputCollection.AddOutput(arrayOfOutputs(iOutput), Geometry_Transform(.RuledSurfaces3d.CreateByCurves(pObjectOfOutputCollection.ResourceManager, ComplexString(1), ComplexString(2), True), pT4x4OfTranslation))

            End If

        End With

        Exit Sub
ErrorHandler:
    ReportUnanticipatedError MODULE, METHOD

    End Sub

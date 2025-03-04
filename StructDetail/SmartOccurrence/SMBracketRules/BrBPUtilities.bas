Attribute VB_Name = "BrBPUtilities"
Option Explicit

Public Const BrBPToDoMsgCodelist = "IntellishipMsgs"
Public Const TDL_BRACKET_BY_PLANE_INVALID_SUPPORTS = 4
                            '"Corner gusset connection has invalid inputs. Delete and replace."
Public Const E_FAIL = -2147467259
Public Const m_sProjectName As String = CUSTOMERID + "BracketRules"
Public Const m_sProjectPath As String = "S:\StructDetail\Data\SmartOccurrence\" + m_sProjectName + "\"


Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\" + CUSTOMERID + "BracketRules\BrBPUtilities.bas"

'*************************************************************************
'Function
'ToDoErrorNotify
'
'Abstract
' Called to notify the SmartOccurrence of a ToDo error that occurred during a
' smart occurrence custom evaluate
'
'***************************************************************************
Public Sub ToDoErrorNotify(strCodelistTable As String, nToDoListErrorNum As Long, oObjectInError As Object, oObjectToUpdate As Object)
    Const METHOD = "ToDoErrorNotify"
    On Error GoTo ErrHandler:
    Dim oToDoListHelper As IJToDoListHelper

    Set oToDoListHelper = oObjectInError ' Set ToDoListHelper = pointer to the CAO Object
    If Not oToDoListHelper Is Nothing Then
        If oObjectToUpdate Is Nothing Then
          oToDoListHelper.SetErrorInfo strCodelistTable, nToDoListErrorNum
        Else
          oToDoListHelper.SetErrorInfo strCodelistTable, nToDoListErrorNum, oObjectToUpdate
        End If
    End If
    
    Exit Sub
    
ErrHandler:
    ' Report the error but do not return the error since this call maybe in their
    '   error handler
    Err.Raise LogError(Err, MODULE, METHOD).Number
    Err.Clear
End Sub
'*************************************************************************
'Function
'GetIntersection
'
'Abstract
' Gets the intersection object between the two input objects
'
'***************************************************************************
Public Function GetIntersection(pObject1 As Object, pObject2 As Object) As Object
    Const METHOD = "GetIntersection"
    On Error GoTo ErrorHandler
 
    Dim oStructIntersector As IMSModelGeomOps.DGeomOpsIntersect
    Dim oIntersectedUnknown As IUnknown
    Dim pAgtorUnk As IUnknown
    Dim NullObject As Object
    
    On Error Resume Next
    Set oStructIntersector = New IMSModelGeomOps.DGeomOpsIntersect
    oStructIntersector.PlaceIntersectionObject NullObject, pObject1, pObject2, pAgtorUnk, oIntersectedUnknown
    If Err.Number <> 0 Then
        Err.Clear
    End If
    On Error GoTo ErrorHandler
   
    Set GetIntersection = oIntersectedUnknown
    
    Set oStructIntersector = Nothing
    Set oIntersectedUnknown = Nothing
    Set pAgtorUnk = Nothing
    
    Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
    Err.Clear
End Function

'*************************************************************************
'Function
'CheckValidityOfSupports
'
'Abstract
' Gets the supports from the input support collection.
' Validates the supports by finding the intersection between the supports
' as per the requirements of the symbol used for the bracket creation with
' this supports.
' Returns a boolean as True or False
'***************************************************************************
Public Function CheckValidityOfSupports(oSupports As IJElements) As Boolean

'Initialize the return value to True
CheckValidityOfSupports = True

If (oSupports Is Nothing) Then
    'If the support count is 0 or 1 then they are invalid.
    CheckValidityOfSupports = False
    Exit Function
End If




Dim oSupport() As Object
Dim i As Integer
Dim nSuppCount As Integer
i = 0
'Get the Support count
nSuppCount = oSupports.Count

ReDim oSupport(nSuppCount - 1)

'Get the Support objects
For i = 1 To nSuppCount
    Set oSupport(i - 1) = oSupports.Item(i)
Next i

Dim oIntersector As Object
If nSuppCount = 2 Then
    ' this is a single support bracket.
    Dim pHelper As StructDetailObjects.Helper
    Set pHelper = New StructDetailObjects.Helper
    If TypeOf oSupport(0) Is ISPSMemberSystem Then
        CheckValidityOfSupports = True
        Exit Function
    End If
    
    If pHelper.ObjectType(oSupport(0)) = SDOBJECT_STIFFENERSYSTEM Then
        Dim oProfileObject As StructDetailObjects.ProfileSystem
        Set oProfileObject = New StructDetailObjects.ProfileSystem
        With oProfileObject
            ' use the affected leaf system as the cross section can change
            Set .object = oSupport(1)
            'if the profile has a flange, it can be used for a "1S"
            ' bracket
            CheckValidityOfSupports = (.FlangeLength > 0#)
        End With
    'ElseIf pHelper.ObjectType(oSupport(0)) = SDOBJECT_MEMBERSYSTEM Then
        ' will need member system struct detail objects See:
        ' CR-CP·116730  Allow member as support for BracketByPlane
    
    Else
        CheckValidityOfSupports = False
    End If


'Check for intersection of Supports as per symbol requirements.
ElseIf nSuppCount = 4 Then
    If TypeOf oSupport(0) Is ISPSMemberSystem Or TypeOf oSupport(1) Is ISPSMemberSystem Then
        CheckValidityOfSupports = True
        Exit Function
    End If
    ' The two supports need to intersect each other else invalid.
     Set oIntersector = GetIntersection(oSupport(0), oSupport(1))
    If oIntersector Is Nothing Then
        CheckValidityOfSupports = False
    End If
ElseIf nSuppCount = 6 Then
     If TypeOf oSupport(0) Is ISPSMemberSystem Or TypeOf oSupport(1) Is ISPSMemberSystem Or TypeOf oSupport(2) Is ISPSMemberSystem Then
        CheckValidityOfSupports = True
        Exit Function
    End If
    ' Support1 should intersect with Support2 and Support3
    Set oIntersector = GetIntersection(oSupport(0), oSupport(1))
    If oIntersector Is Nothing Then
        CheckValidityOfSupports = False
        Exit Function
    End If
    Set oIntersector = Nothing
    Set oIntersector = GetIntersection(oSupport(0), oSupport(2))
    If oIntersector Is Nothing Then
        CheckValidityOfSupports = False
        Exit Function
    End If
ElseIf nSuppCount = 8 Then
    ' Support1 should intersect with Support2 and Support3
    ' Support2 should intersect with Support4
    Set oIntersector = GetIntersection(oSupport(0), oSupport(1))
    If oIntersector Is Nothing Then
        CheckValidityOfSupports = False
        Exit Function
    End If
    
    Set oIntersector = Nothing
    Set oIntersector = GetIntersection(oSupport(0), oSupport(2))
    If oIntersector Is Nothing Then
        CheckValidityOfSupports = False
        Exit Function
    End If
    
    Set oIntersector = Nothing
    Set oIntersector = GetIntersection(oSupport(1), oSupport(3))
    If oIntersector Is Nothing Then
        CheckValidityOfSupports = False
        Exit Function
    End If
ElseIf nSuppCount = 10 Then
    ' Support1 should intersect with Support2 and Support3
    ' Support2 should intersect with Support4
    ' Support3 should intersect with Support5
    Set oIntersector = GetIntersection(oSupport(0), oSupport(1))
    If oIntersector Is Nothing Then
        CheckValidityOfSupports = False
        Exit Function
    End If
    
    Set oIntersector = Nothing
    Set oIntersector = GetIntersection(oSupport(0), oSupport(2))
    If oIntersector Is Nothing Then
        CheckValidityOfSupports = False
        Exit Function
    End If
    
    Set oIntersector = Nothing
    Set oIntersector = GetIntersection(oSupport(1), oSupport(3))
    If oIntersector Is Nothing Then
        CheckValidityOfSupports = False
        Exit Function
    End If

    Set oIntersector = Nothing
    Set oIntersector = GetIntersection(oSupport(2), oSupport(4))
    If oIntersector Is Nothing Then
        CheckValidityOfSupports = False
        Exit Function
    End If
End If

Set oIntersector = Nothing
For i = LBound(oSupport) To UBound(oSupport)
    Set oSupport(i) = Nothing
Next i
End Function
'*************************************************************************
'Function
'IsFlangeIn
'
'Purpose:
'   Determines is the flange for a giving support is in the direction of
'   the bracket.
'Returns:
'   True   - if Flange is IN (orientation is towards the direction of the bracket.)
'   False - if Flange is OUT (orientation is away from the direction of the bracket.)
'*********************************************************************************************
Public Function IsFlangeTowardsBracket(oProfileSupport As Object, oBracketByPlaneSO As IJSmartOccurrence, uBracketPoint As IJPoint, vBracketPoint As IJPoint) As Boolean

'Assume true...
IsFlangeTowardsBracket = True
                 
'*********************************************************************************************
 '1. Get interstion point between supports and bracket plate
 Dim oBracketUtils As GSCADCreateModifyUtilities.IJBracketAttributes
Set oBracketUtils = New GSCADCreateModifyUtilities.PlateUtils
Dim oBracketPlate As IJPlate
Set oBracketPlate = oBracketUtils.GetBracketByPlaneFromBracketContour(oBracketByPlaneSO)
Dim oIntersectUtil As IIntersect
Set oIntersectUtil = New Intersect
Dim oCommonBody As IUnknown
Dim oStructGeometry As Object
'
'Debug.Print "TypeOf oBracketPlate Is IJDModelBody = " & TypeOf oBracketPlate Is IJDModelBody
'Debug.Print "TypeOf oBracketPlate Is IJPlate = " & TypeOf oBracketPlate Is IJPlate
'Debug.Print "TypeOf oBracketPlate Is IJDGeometry = " & TypeOf oBracketPlate Is IJDGeometry

'During recursion due, need to check if the object is a struct geometry type...

On Error Resume Next
If TypeOf oBracketPlate Is IJDModelBody Then
    ' Just query
    Set oStructGeometry = oBracketPlate
    oIntersectUtil.GetCommonGeometry oProfileSupport, oStructGeometry, oCommonBody, 0
End If
If oCommonBody Is Nothing Then
    Set oStructGeometry = oBracketUtils.GetLimitedBracketPlane(oBracketPlate)
    oIntersectUtil.GetCommonGeometry oProfileSupport, oStructGeometry, oCommonBody, 0
End If
On Error GoTo ErrorHandler:

'Debug.Print ""

'oIntersectUtil.GetCommonGeometry oProfileSupport, oStructGeometry, oCommonBody, 0

Dim oFlangeOrientVec As IJDVector
Dim oWebOrientVec As IJDVector
Dim oPointsGraphBody As IJPointsGraphBody
Dim oPntGraphUtils As New SGOPointsGraphUtilities
Set oPointsGraphBody = oCommonBody
Dim oPointColl As Collection
Set oPointColl = oPntGraphUtils.GetPositionsFromPointsGraph(oPointsGraphBody)
Dim oOrientationPosition As IJDPosition
Set oOrientationPosition = oPointColl.Item(1)
                 
Set oBracketUtils = Nothing

'*********************************************************************************************
'2. Get the direction vector for first and seondary orientation of profile sent...
Dim oProfilehelper As IJProfileAttributes
Set oProfilehelper = New ProfileUtils
oProfilehelper.GetProfileOrientation oProfileSupport, oOrientationPosition, oFlangeOrientVec, oWebOrientVec
 
'*********************************************************************************************
'3. Found out what support was sent and choice either the U or V vector from the bracket...
Dim oBracketDir As IJDVector
Dim oS1 As Object
Dim oS2 As Object
Dim oS3 As Object
Dim oS4 As Object
Dim oS5 As Object
Dim oSupports As IJElements
Dim nNumSupports As Long
Set oBracketUtils = New GSCADCreateModifyUtilities.PlateUtils

oBracketUtils.GetSupportsFromBracketContourSO oBracketByPlaneSO, oSupports, nNumSupports

Dim i As Long
For i = 1 To oSupports.Count
    Select Case i
        Case 1
            Set oS1 = oSupports.Item(i)
        Case 2
            Set oS2 = oSupports.Item(i)
        Case 3
            Set oS3 = oSupports.Item(i)
        Case 4
            Set oS4 = oSupports.Item(i)
        Case 5
            Set oS5 = oSupports.Item(i)
    End Select
Next i

Dim xDirPoint As Double
Dim yDirPoint As Double
Dim zDirPoint As Double

If oProfileSupport Is oS1 Then
    ' Bracket is attached to flange, therefore no need to determine flange direction...
    Exit Function
ElseIf (oProfileSupport Is oS2) Or (oProfileSupport Is oS3) Then
    
    uBracketPoint.GetPoint xDirPoint, yDirPoint, zDirPoint
    Set oBracketDir = New DVector
    
    oBracketDir.x = xDirPoint
    oBracketDir.y = yDirPoint
    oBracketDir.z = zDirPoint

ElseIf (oProfileSupport Is oS4) Or (oProfileSupport Is oS5) Then

    vBracketPoint.GetPoint xDirPoint, yDirPoint, zDirPoint
    Set oBracketDir = New DVector
    
    oBracketDir.x = xDirPoint
    oBracketDir.y = yDirPoint
    oBracketDir.z = zDirPoint
    
End If

'*********************************************************************************************
'4. Determine if the direction between the flange direction is in the same direction as the
'    the bracket...
Dim fDot As Double

fDot = oBracketDir.x * oFlangeOrientVec.x + _
           oBracketDir.y * oFlangeOrientVec.y + _
           oBracketDir.z * oFlangeOrientVec.z

If fDot < 0 Then
    IsFlangeTowardsBracket = True
Else
    IsFlangeTowardsBracket = False
End If

Exit Function
'Clean up...
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "IsFlangeTowardsBracket").Number
    Err.Clear
End Function

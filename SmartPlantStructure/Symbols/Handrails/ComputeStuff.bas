Attribute VB_Name = "ComputeStuff"
' ComputeStuff
Option Explicit

'******************************************************************
'  Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'  File               ComputeStuff.bas
'  ProgID             SPSHandrailMacros.ComputeStuff
'  Author             Satish N Kota
'  Creation Date      [May 08, 2003]
'  Description        Computes the common Handrails Attributes vb bas file
'
'  Change History:
'  6th Feb 004   mkaveti   Added new public method to validate user keyed in
'                           for the attributes Height,ToTopOfMidrailDistance,
'                           NoOfMidrails and MidRailSpacing
'
'  10 Sept 2004 Manish Dhotre
' Skiniing option added which can be used to create handrail surface based on requirement
' Skinning option is used to set "brkcrv " argument of CreateBySingleSweep() function and also
' through CreateSurface wrapper provided in CrossSectionService.bas. So proper argument need to be sent
' to CreateSurface() through CreateProjectionFromCSProfile(). For more details of
' appropriate skiining  option see definition of CreateProjectionFromCSProfile() & documentation of CreateBySingleSweep
' Skinning option will not work for cases like toprail section radius is high and handrail path bend radius is
' too small for given radius. So validation code added to check if CreateBySingleSweep works
' for given path & cross section. If not then skinning option will be changed. This check is done for
' Physical  representations. Can also be added for operational representation if based on requirement
'
' 6 Jun 05      JS      Went to correct a problem where the Error handlers never had m_Errors
'                       declared (TR#79298) & discovered other issues -- the
'                       judiscious use of "On Error Resume Next" was masking many problems. To
'                       use this function correctly one should immediately reset back to the error
'                       handler after the offending statement; otherwise, other coding problems later
'                       in the code will be masked (which was found in several instances).
'                       Also, resume next is not necessary to determine if an object supports
'                       an interface "Typeof" should be used instead.
'
'   05-oct-06   SS      TR#107161 - commented out msgbox as we should not be raising any msgboxes from
'                       symbol could. If such error condition happens during sync, ifc it could
'                       stop the process.
'
'   11-Oct-08   SS      CR#38818 - add handrail sketched path as its output, locatable but not displayable
'
'   16-Apr-08   SS      TR#137322 - Infinite routine on Line will make the line infinite in both directions
'                       this will be an issue for collinear path segments.This is not required.
'
'   06-Aug-08   SS      TR#142601 - changed path control flags to NO_DRAW instead of NO_DISPLAY as it hinders
'                       locating handrail path.
'   07-Aug-08   WR      TR-CP-131048 - Added new methods to GetTopRailRadius() and GetTopRailInfo().  These
'                       methods are used by the symbols to retrieve the radius of the top rail.
'
'   02-Oct-08   SS      CR#148021 - Add Centerline representation to support Drawings for CR112776.
'                       DLL version is bumped and will require a synchronize.
'
'   02-Oct-08   SS      CR#141607 - Removed Detailed Physical from handrail representations as it is same
'                       as SimplePhysical and redundant.
'   12-Aug-09   GG      DM-169707  - Fixed the weight issues and some COG issues
'   25-Aug-09   GG      TR-168826 and TR-168827  - Fixed the weight issues and COG issues for MDR and multi copy/paste
'                       Always use output collection from data source
'   04-Sept-09  GG      DM-171219 Changed the way to calculate weight and COG for handrails. Now the WCG calculation
'                       is independent to the Physical evaluation. The main changes in this module are:
'                       1. Moved CalcWCG and CalcWCG_A1 into related class modules with modifications
'                       2. Moved CreatePosts subroutine back to TypeATopEmbedded class module
'                       3. Created subroutines AddPostVolumeInfo, AddTreatmentVolumeInfo, and CalcRailsVolumeCG
'                       4. In AddPostVolumeInfo, the effective length for circular treatment compare to the rectangular treatment is H*(PI+2)/6=0.857*H, assume the cir radius is H/6
'                       5. Removed function GetOutputCollFromSource
'   20-Nov-09 GG        DM-175208  Circular end treatment adapts to the post section size
'                       1. Always use toprail cross section for circular treatment;
'                       2. Changed path direction to be consistent with toprail;
'                       3. Fixed the mirror direction for circular treatment;
'                       4. If the assumption for circular radius is not a good one, a formular is used to get the radius.
'   14-Dec-09   MJ      TR-169906 Horizontal Offsets are not applied when figuring COG of handrail
'                       The fix for this TR is to ensure that the symbol's Calculate Weight and CG
'                       follows the same logic that is used by it's physical representation calculation.
'                       This involved significant code refactoring so that both custom methods are able
'                       to call the same code
'   10-Feb-10   GG      Fixed the following issues:
'                       1.  Use top-rail cross section for the rectangular treatment (TR-173639);
'                       2.  Corrected the orientation of the rectangular end treatment or the last post if there is no end treatment(DM -176290);
'                       3.  Fixed the centerline graphics for treatments (TR-169532, TR-169001);
'                       4.  Made the pad length and pad width consistent to the legend (TR-161195);
'                       5.  The Spacing between posts is always greater than the minimum clearance if "With Post at Turn" is false (CR-58237).
'  17-Feb-10    MH      BuildHandrailOutput for ConvertHandrail
'   14-May-10   GG      DM 182326. Allow the open end treatment
'                       DM 182314. The horizontal offset direction is dependent on the direction of the curvature for curved supporting member.
'   17-Aug-10   GG      DM 186125. Circular end treatment of handrail is not correct if the path is sloped
'   23-Nov-10   GG      DI 177040. Added subroutine SetAttributeValue and reuse saved PlateDimensions proxy
'   29-Dec-10   GG      TR 191191  The handrail computing fails for some path patterns. Also removed some obsolete subs
'
'*******************************************************************

Private Const MODULE = "ComputeStuff::"

Public Const E_FAIL = -2147467259
Private oLocalizer As IJLocalizer
Public Const PI = 3.14159265358979

'see CreateProjectionFromCSProfile() in CrossSectionService.bas for more details of skin option
Public SkinOption As Long

Public m_complex As IngrGeom3D.ComplexString3d
Public m_transform As IngrGeom3D.IJDT4x4
Public m_GeomFactory As IngrGeom3D.GeometryFactory
Public m_oCatResMgr As IUnknown

Public Enum eSEGMENTTYPE
    SEGMENT_BEGIN_TYPE = 0
    SEGMENT_END_TYPE = 1
    SEGMENT_MID_TYPE = 2
End Enum

Public Enum InputIndex
    PART_INDEX = 1
    SKETCHOBJ_INDEX
    HEIGHT_INDEX
    WITHTOEPLATE_INDEX
    MIDRAILSNO_INDEX
    HORIZONTAL_OFFSET_TYPE_INDEX
    HORIZONTAL_OFFSET_DIM_INDEX
    ORIENTATION_INDEX
    MAXSPACING_INDEX
    SLOPE_MAXSPACING_INDEX  '10
    TOPOFTOEPLATE_DIM_INDEX
    TOPOFMIDRAIL_DIM_INDEX
    MIDRAIL_SPACING_DIM_INDEX
    POST_AT_TURN_INDEX
    BEGIN_TREAT_INDEX
    BEGIN_EXT_DIM_INDEX
    END_TREAT_INDEX
    END_EXT_DIM_INDEX
    ISASSEMBLY_INDEX
    ISSYSTEM_INDEX          '20
    TOPRAIL_CSNAME_INDEX
    TOPRAIL_CSSTD_INDEX
    TOPRAIL_CSCP_INDEX
    TOPRAIL_CSANGLE_INDEX
    MIDRAIL_CSNAME_INDEX
    MIDRAIL_CSSTD_INDEX
    MIDRAIL_CSCP_INDEX
    MIDRAIL_CSANGLE_INDEX
    TOEPLATE_CSNAME_INDEX
    TOEPLATE_CSSTD_INDEX    '30
    TOEPLATE_CSCP_INDEX
    TOEPLATE_CSANGLE_INDEX
    POST_CSNAME_INDEX
    POST_CSSTD_INDEX
    POST_CSCP_INDEX
    POST_CSANGLE_INDEX
    MATERIAL_INDEX
    GRADE_INDEX
    MIN_CLEARANCE_AT_POST_TURN_INDEX
    MAX_CLEARANCE_AT_POST_TURN_INDEX '40
End Enum

Public Enum eRepresentationType
    SimpleRep = 1
    CenterLineRep
End Enum

Public Enum eHandrailType
    TypeA = 1
    TypeASideMntToMem
    TypeASideMount
    TypeATopEmbedded
    TypeATopMounted
End Enum
Const strIJDOutputCollectionUUID = "{15916CAF-6CB5-11D1-A655-00A0C98D7F13}"
Const strOutputName = "toOutputs"

Public Const dtol = 0.00001

Public Const MemberType_Post = 602
Public Const MemberType_TopRail = 603
Public Const MemberType_MidRail = 604
Public Const MemberType_ToePlate = 605
Public Const MemberType_EndTreatment = 606


Public Function GetAttributeCollection(oBO As Object, attrInterface As String) As Object
    Const METHOD = "GetAttributeCollection"
    On Error GoTo ErrorHandler
    Dim pIJAttrbs As IJDAttributes
    
    If Not oBO Is Nothing Then
        Set pIJAttrbs = oBO
        On Error Resume Next
        Set GetAttributeCollection = pIJAttrbs.CollectionOfAttributes(attrInterface)
        Err.Clear
    End If
    Exit Function

ErrorHandler:
    HandleError MODULE, METHOD
End Function

Public Function GetAttributeValue(oAttrCollection As CollectionProxy, attrName As String) As Variant
    Const METHOD = "GetAttributeValue"
    On Error GoTo ErrorHandler
    
    Dim attrValueEmpty As Variant
    Dim oAttr As IJDAttribute

    GetAttributeValue = attrValueEmpty

    If Not oAttrCollection Is Nothing Then
        On Error Resume Next
        Set oAttr = oAttrCollection.Item(attrName)
        If Err.Number = 0 Then
            GetAttributeValue = oAttr.Value
        Else
            Err.Clear
        End If
    End If
    
    Exit Function

ErrorHandler:
    HandleError MODULE, METHOD
End Function


Public Sub SetAttributeValue(oAttrCollection As CollectionProxy, attrName As String, attrValue As Variant)
    Const METHOD = "SetAttributeValue"
    On Error GoTo ErrorHandler
    
    'Dim attrValueEmpty As Variant
    Dim oAttr As IJDAttribute

    'GetAttributeValue = attrValueEmpty

    If Not oAttrCollection Is Nothing Then
        On Error Resume Next
        Set oAttr = oAttrCollection.Item(attrName)
        If Err.Number = 0 Then
            oAttr.Value = attrValue
        Else
            Err.Clear
        End If
    End If
    
    Exit Sub

ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub BuildHandrailMemberSystem(ByVal oHandrail As ISPSHandrail, ByVal oTraceCurve As Object, ByVal oSectionDef As Object, _
                        ByVal lMemberType As Long, ByVal lCardinalPoint As Long, ByVal dAngle As Double, ByVal bMirror As Boolean)

' Use oHandrail to pass to MemberFactory for the POM
' Use oHandrail for PG, and the system owner
' Set Priority to tertiary.
' Add the newly created MemberSystem to the handrail's cached list
    On Error GoTo ErrorHandler
    Const METHOD = "BuildHandrailMemberSystem"

    Dim oAlign As ISPSMemberSystemAlignment
    Dim oMemberFactory As ISPSMemberFactory
    Dim oMemberSystem As ISPSMemberSystem
    Dim oMemberPart As ISPSMemberPartPrismatic
    Dim oObjectPart As iJDObject, oObjectHandrail As iJDObject  ' for setting PG
    Dim oSystem As IJDesignParent  ' for adding system as a child.
    Dim iCurve As IJCurve
    
    Dim oAttrColl As Object
    Dim varMaterial As Variant
    Dim varGrade As Variant
    Dim sX As Double, sY As Double, sZ As Double, eX As Double, eY As Double, eZ As Double

    Set oMemberFactory = New SPSMemberFactory
    Set iCurve = oTraceCurve
        
    If TypeOf oTraceCurve Is IJLine Then
        iCurve.EndPoints sX, sY, sZ, eX, eY, eZ
        
        Set oMemberSystem = oMemberFactory.CreateMemberSystemPrismaticLinear(oHandrail)
        
        oMemberSystem.LogicalAxis.SetLogicalStartPoint sX, sY, sZ
        oMemberSystem.LogicalAxis.SetLogicalEndPoint eX, eY, eZ

        Set oAlign = oMemberSystem
        oAlign.Align = True
        Set oAlign = Nothing
    
    Else
        Set oMemberSystem = oMemberFactory.CreateMemberSystemCurve(oHandrail)
        Dim oSketchFactory As Sketch3dFactory
        Set oSketchFactory = New Sketch3dFactory
        Dim oHRObj As iJDObject
        Set oHRObj = oHandrail
        
        Dim oSketch As Sketch3d
        Set oSketch = oSketchFactory.CreateFromCurve(oHRObj.ResourceManager, _
                                                     oTraceCurve)
        Set oMemberSystem.LogicalAxis.CurveObject = oSketch
    End If

    oMemberSystem.MemberType.TypeCategory = SPSMembers.SPSMemberTypeCategory.SPSMemberTypeCategory_Handrail
    oMemberSystem.MemberType.Type = lMemberType
    oMemberSystem.MemberType.Priority = 3  'tertiary

    Set oMemberPart = oMemberSystem.DesignPartAtEnd(SPSMemberAxisStart)

    Set oMemberPart.CrossSection.definition = oSectionDef
    oMemberPart.CrossSection.CardinalPoint = lCardinalPoint
    oMemberPart.Rotation.BetaAngle = dAngle
    oMemberPart.Rotation.Mirror = bMirror

    If oHandrail.ConvertHelper.MaterialProxy Is Nothing Then
        Set oAttrColl = GetAttributeCollection(oHandrail, "IJUAHRTypeAProps")
        varMaterial = GetAttributeValue(oAttrColl, "Primary_SPSMaterial")
        varGrade = GetAttributeValue(oAttrColl, "Primary_SPSGrade")
        Set oHandrail.ConvertHelper.MaterialProxy = GetMaterialProxyFromNames(oMemberPart, varMaterial, varGrade)
    End If
    
    Set oMemberPart.MaterialDefinition = oHandrail.ConvertHelper.MaterialProxy
    
    Set oObjectPart = oMemberPart
    Set oObjectHandrail = oHandrail
    oObjectPart.PermissionGroup = oObjectHandrail.PermissionGroup

    ' the same memberType name rules apply to linear and curved
    SetEntityNameRule oMemberSystem, "SPSMembers.SPSMemberSystemLinear", "MemberSystemTypeNameRule"
    SetEntityNameRule oMemberPart, "SPSMembers.SPSMemberPartPrismatic", "MemberPartTypeNameRule"

    Set oSystem = oHandrail
    oSystem.AddChild oMemberSystem

    ' cache the MemberSystem for use by the ConnectMembers method.
    If lMemberType = MemberType_Post Then
        oHandrail.ConvertHelper.Posts.Add oMemberSystem
    ElseIf lMemberType = MemberType_TopRail Then
        oHandrail.ConvertHelper.TopRails.Add oMemberSystem
    ElseIf lMemberType = MemberType_MidRail Then
        oHandrail.ConvertHelper.MidRails.Add oMemberSystem
    ElseIf lMemberType = MemberType_ToePlate Then
        oHandrail.ConvertHelper.ToePlates.Add oMemberSystem

    ' need to determine whether this is a Begin or End treatment
    ElseIf lMemberType = MemberType_EndTreatment Then
        Dim oHandrailSketch As IJDSketch3d
        Dim oSketchCurve As IJCurve
        Dim distStart As Double, distEnd As Double
        Dim sketchStartX As Double, sketchStartY As Double, sketchStartZ As Double
        Dim sketchEndX As Double, sketchEndY As Double, sketchEndZ As Double
        Dim oPos As DPosition
        Set oPos = New DPosition
        
        Set oHandrailSketch = oHandrail.SketchPath
        Set oSketchCurve = oHandrailSketch.GetComplexString
        oSketchCurve.EndPoints sketchStartX, sketchStartY, sketchStartZ, sketchEndX, sketchEndY, sketchEndZ
        
        ' distStart is distance from start point of sketchPath to the endTreatment curve
        ' distEnd is distance from end point of sketchPath to the endTreatment curve
        
        oPos.Set sketchStartX, sketchStartY, sketchStartZ
        'oLine.DefineBy2Points sketchStartX, sketchStartY, sketchStartZ, sketchStartX, sketchStartY, sketchStartZ
        iCurve.DistanceBetween oPos, distStart, sX, sY, sZ, eX, eY, eZ
        oPos.Set sketchEndX, sketchEndY, sketchEndZ
        'oLine.DefineBy2Points sketchEndX, sketchEndY, sketchEndZ, sketchEndX, sketchEndY, sketchEndZ
        iCurve.DistanceBetween oPos, distEnd, sX, sY, sZ, eX, eY, eZ
    
        If distStart < distEnd Then     ' endTreatment is closer to sketchPath start
            Set oHandrail.ConvertHelper.BeginTreatment = oMemberSystem
        Else
            Set oHandrail.ConvertHelper.EndTreatment = oMemberSystem
        End If
    End If
                
    Set oMemberSystem = Nothing

    Exit Sub

ErrorHandler:
    Dim oErrors As New IMSErrorLog.JServerErrors
    oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
End Sub
                        
Public Sub BuildHandrailOutput(oOutputParent As Object, oTraceCurve As Object, _
                    oSection As Object, lCP As Integer, dAngle As Double, bMirror As Boolean, _
                    lSkinOption As Long, strOutputPrefix As String, lMemberType As Long, _
                    ByRef pCountCreated As Long)
                    
    On Error GoTo ErrorHandler
    Const METHOD = "BuildHandrailOutput"

    Dim ii As Long, ncount As Long
    Dim elesObjects As IJElements
    Dim oObj As Object
    Dim iComplexString As IJComplexString
    Dim startx As Double, starty As Double, startz As Double
    Dim Endx As Double, Endy As Double, Endz As Double

    If TypeOf oOutputParent Is IJDOutputCollection Then

        Dim strOutName As String
        Dim oOutputCollection As IJDOutputCollection
        Dim iMirror As Integer

        Set oOutputCollection = oOutputParent
        If bMirror Then
            iMirror = 1
        Else
            iMirror = 0
        End If
        
        If Not oSection Is Nothing Then
        
            If lMemberType = MemberType_ToePlate Then   ' no twist.
                Set elesObjects = CreateProjections(oOutputCollection, oTraceCurve, oSection, lCP, dAngle, iMirror, lSkinOption)
            Else
                Set elesObjects = CreateProjectionFromCSProfile(Nothing, oTraceCurve, oSection, lCP, dAngle, iMirror, Nothing, Nothing, lSkinOption)
            End If
        End If

        If Not elesObjects Is Nothing Then

            ncount = elesObjects.Count
            For ii = 1 To ncount
                Set oObj = elesObjects(ii)
                If pCountCreated = 0 Then
                    strOutName = strOutputPrefix
                Else
                    strOutName = strOutputPrefix & pCountCreated
                End If
                InitNewOutput oOutputCollection, strOutName
                oOutputCollection.AddOutput strOutName, oObj
                Set oObj = Nothing
                pCountCreated = pCountCreated + 1
            Next ii
        
            Set elesObjects = Nothing
            ' if CreateProjectionFromCSProfile produced something

        Else
            
            ' add the trace curve to sym output to show that skinning failed or no cross section, or Centerline rep.
            Dim oOutTraceCurve As Object
            Dim oGeomFactory As IngrGeom3D.GeometryFactory
            Set oGeomFactory = New GeometryFactory
            
            If TypeOf oTraceCurve Is IJComplexString Then
                Set iComplexString = oTraceCurve            ' oTraceCurve is not persistent.  so read it to make a persistent one.
                iComplexString.GetCurves elesObjects
                Set oOutTraceCurve = oGeomFactory.ComplexStrings3d.CreateByCurves(Nothing, elesObjects)

            ElseIf TypeOf oTraceCurve Is Arc3d Then
                Dim iArc As Arc3d
                Dim normalx As Double, normaly As Double, normalz As Double
                Dim centerx As Double, centery As Double, centerz As Double
                
                Set iArc = oTraceCurve
                iArc.GetStartPoint startx, starty, startz
                iArc.GetEndPoint Endx, Endy, Endz
                iArc.GetNormal normalx, normaly, normalz
                iArc.GetCenterPoint centerx, centery, centerz

                Set oOutTraceCurve = oGeomFactory.Arcs3d.CreateByCtrNormStartEnd(Nothing, centerx, centery, centerz, _
                                        normalx, normaly, normalz, startx, starty, startz, Endx, Endy, Endz)
            
            ElseIf TypeOf oTraceCurve Is Line3d Then
                Dim iLine As Line3d
                Set iLine = oTraceCurve
                iLine.GetStartPoint startx, starty, startz
                iLine.GetEndPoint Endx, Endy, Endz
            
                Set oOutTraceCurve = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                        startx, starty, startz, Endx, Endy, Endz)
            End If
    
            If Not oOutTraceCurve Is Nothing Then
                If pCountCreated = 0 Then
                    strOutName = strOutputPrefix
                Else
                    strOutName = strOutputPrefix & pCountCreated
                End If
    
                InitNewOutput oOutputCollection, strOutName
                oOutputCollection.AddOutput strOutName, oOutTraceCurve
                Set elesObjects = Nothing
                Set oOutTraceCurve = Nothing
                
                pCountCreated = pCountCreated + 1
            End If
        End If  ' else if using the trace curve as output.
        
    Else    ' else it is a Handrail type.  ( Used by the Convert Handrail process. )
        
        ' NOTE:
        ' This is less than desirable as long as we do not support curved members.  The middle segment of
        ' a circular end treatment does not have an orientation that is consistent with the curved section
        ' ( that we do not create. )  Once curved members are supported, then this should be changed to
        ' check whether the complex string has arcs.  If it does, and it is an endTreatment, then the whole
        ' complex string should be created as a single curved memberSystem.

        If TypeOf oTraceCurve Is IJComplexString Then
            Dim iCurve As IJCurve
            Set iComplexString = oTraceCurve
            iComplexString.GetCurves elesObjects
            Dim bCreateAsSinglMember As Boolean
            bCreateAsSinglMember = False
            ncount = elesObjects.Count
        
            If MemberType_EndTreatment = lMemberType Then
                For ii = 1 To ncount
                    If Not TypeOf elesObjects(ii) Is Line3d Then
                        bCreateAsSinglMember = True
                        Exit For
                    End If
                Next ii
            End If
            If bCreateAsSinglMember = False Then
                For ii = 1 To ncount
            
                    Set iCurve = elesObjects(ii)
    
                    BuildHandrailOutput oOutputParent, iCurve, oSection, lCP, dAngle, bMirror, lSkinOption, strOutputPrefix, lMemberType, pCountCreated
                
                Next ii
            Else
                BuildHandrailMemberSystem oOutputParent, oTraceCurve, oSection, lMemberType, lCP, dAngle, bMirror
            
                pCountCreated = pCountCreated + 1
            End If
            Set elesObjects = Nothing

        Else
        
            BuildHandrailMemberSystem oOutputParent, oTraceCurve, oSection, lMemberType, lCP, dAngle, bMirror
            
            pCountCreated = pCountCreated + 1
        
        End If
        
    End If

    Exit Sub

ErrorHandler:
    Dim oErrors As New IMSErrorLog.JServerErrors
    oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
End Sub

Public Sub SetEntityNameRule(oEntity As Object, strEntityProgId As String, strNameRuleName As String)
    Const METHOD = "SetEntityNameRule"
    On Error GoTo ErrHandler

    Dim elesNamingRules As IJElements
    Dim oNameRuleHolder As GSCADGenericNamingRulesFacelets.IJDNameRuleHolder
    Dim oNameRuleHlpr As GSCADNameRuleSemantics.IJDNamingRulesHelper
    Dim oNameRuleAE As GSCADGenNameRuleAE.IJNameRuleAE

    Set oNameRuleHlpr = New GSCADNameRuleHlpr.NamingRulesHelper
    Call oNameRuleHlpr.GetEntityNamingRulesGivenProgID(strEntityProgId, elesNamingRules)
    
    Dim ncount As Long, ii As Long
    
    ncount = elesNamingRules.Count
    Set oNameRuleHolder = Nothing
    ' look for the strNameRuleName
    For ii = 1 To ncount
        Set oNameRuleHolder = elesNamingRules.Item(ii)
        If strNameRuleName = oNameRuleHolder.name Then

            Exit For
        Else
            Set oNameRuleHolder = Nothing
        End If
    Next ii
    
    ' if not found, then use "DefaultNameRule"
    If oNameRuleHolder Is Nothing Then
        strNameRuleName = "DefaultNameRule"
        For ii = 1 To ncount
            Set oNameRuleHolder = elesNamingRules.Item(ii)
            If strNameRuleName = oNameRuleHolder.name Then
                Exit For
            Else
                Set oNameRuleHolder = Nothing
            End If
        Next ii
    End If
    
    ' if not found, "DefaultNameRule" and there are items in elesNamingRules, then use the first item
    If oNameRuleHolder Is Nothing And ncount > 0 Then
        Set oNameRuleHolder = elesNamingRules.Item(1)
    End If
    
    If Not oNameRuleHolder Is Nothing Then
        Call oNameRuleHlpr.AddNamingRelations(oEntity, oNameRuleHolder, oNameRuleAE)
    End If

    Set oNameRuleHolder = Nothing
    Set oNameRuleHlpr = Nothing
    Set oNameRuleHolder = Nothing
    Set oNameRuleAE = Nothing
    Exit Sub
ErrHandler:
    Dim oErrors As New IMSErrorLog.JServerErrors
    oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
End Sub

Private Function CreateProjections(pOC As IJDOutputCollection, _
                                    pTraceCurve As Object, _
                                    oCSOcc As Object, _
                                    m_cp As Integer, _
                                    m_rotation As Double, _
                                    m_mirror As Integer, _
                                    skinningOption As Long) As IJElements

Const METHOD = "CreateProjections"
On Error GoTo ErrorHandler

    Dim pProfiles As IJElements
    Dim trans4x4 As IJDT4x4
    Dim m_cpx As Double
    Dim m_cpy As Double

    Dim m_xService As SP3DStructGenericTools.CrossSectionServices
    Set m_xService = New SP3DStructGenericTools.CrossSectionServices

    m_xService.GetProfiles oCSOcc, "SimplePhysical", pProfiles
    m_xService.GetCardinalPoint oCSOcc, m_cp, m_cpx, m_cpy

    Set trans4x4 = New DT4x4
    m_xService.ComputeTransformForProjection pTraceCurve, m_mirror, m_rotation, m_cpx, m_cpy, trans4x4
   
    Dim i As Integer
    Dim sectionprof As SectionProfile
    For i = 1 To pProfiles.Count
        Set sectionprof = pProfiles.Item(i)
        sectionprof.Holes = Nothing
    Next i

    Dim elesProjections As IJElements, elesOutProjections As IJElements
    Dim numCaps As Long
    Dim sNorm() As Double, eNorm() As Double
    Dim secProfile As SectionProfile
    Dim outCurve As Object
    Dim oTransform As IJDGeometry
    Dim vec As Variant

    For i = 1 To pProfiles.Count

        Set secProfile = pProfiles.Item(i)
        If Not secProfile Is Nothing Then
            m_xService.CreateCurveCopyCCW secProfile, Nothing, outCurve
        Else
            GoTo ErrorHandler
        End If
        
        If Not outCurve Is Nothing Then
            Set oTransform = outCurve
        Else
            GoTo ErrorHandler
        End If
        
        oTransform.DTransform trans4x4

        Set elesProjections = m_GeomFactory.GeometryServices.CreateBySingleSweepWCapsOpts(Nothing, _
                                                                        pTraceCurve, outCurve, _
                                                                        SkinningCornerOptions.AverageCorner, _
                                                                        SkinningBreakOptions.BreakPathAndCrossSection, _
                                                                        SkinningCrossSectionStart.StartAtTraceBeg, _
                                                                        SkinningCrossSectionOrientation.ZOrientation, _
                                                                        sNorm, eNorm, True, numCaps)
        If elesOutProjections Is Nothing Then
            Set elesOutProjections = elesProjections
        Else
            elesOutProjections.AddElements elesProjections
        End If
        
        Set elesProjections = Nothing
        Set outCurve = Nothing
        Set oTransform = Nothing
        
    Next i

    Set trans4x4 = Nothing
    Set m_xService = Nothing
    Set pProfiles = Nothing
    Set oTransform = Nothing
    Set secProfile = Nothing
    
    Set CreateProjections = elesOutProjections
    Exit Function
                                    
ErrorHandler:
    Dim oErrors As New IMSErrorLog.JServerErrors
    oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
End Function

Public Function GetAttribute(OccAttrs As IJDAttributes, strAttribute As String, InfosColl As IJDInfosCol)    'this is collection of all interfaces
Const METHOD = "GetAttribute"
On Error GoTo ErrHandler
    
    Dim oAttr As IJDAttribute
    Dim Attrcol As IJDAttributesCol
    Dim AttrInfo As IJDAttributeInfo
    Dim oAttributeMetaData As IJDAttributeMetaData
    
    Dim oDBTypeConfig As IJDBTypeConfiguration
    Dim jContext As IJContext
    Dim oConnectMiddle As IJDAccessMiddle
    Dim strModelDBID As String
     
    Set jContext = GetJContext()
    Set oDBTypeConfig = jContext.GetService("DBTypeConfiguration")
    Set oConnectMiddle = jContext.GetService("ConnectMiddle")
    
    strModelDBID = oDBTypeConfig.get_DataBaseFromDBType("Model")
    Set oAttributeMetaData = oConnectMiddle.GetResourceManager(strModelDBID)
    
    Dim icounter As Integer
    Dim iInterfaceCount As Integer
        
    If Not InfosColl Is Nothing Then
        iInterfaceCount = InfosColl.Count
        If iInterfaceCount > 0 Then
        For icounter = 1 To iInterfaceCount
            On Error Resume Next
            Set AttrInfo = oAttributeMetaData.AttributeInfo(InfosColl.Item(icounter).Type, strAttribute)
            On Error GoTo ErrHandler
            If Not AttrInfo Is Nothing And Err.Number = 0 Then
                Set Attrcol = OccAttrs.CollectionOfAttributes(InfosColl.Item(icounter).Type)
                Set oAttr = Attrcol.Item(strAttribute)
                GetAttribute = oAttr.Value
                Exit For
            End If
            Set AttrInfo = Nothing
        Next icounter
        End If
    End If
    If oAttr Is Nothing Then
        Debug.Print "Could not get " & strAttribute
    End If
    Set oAttr = Nothing
    Set Attrcol = Nothing
    Set AttrInfo = Nothing
    Set oAttributeMetaData = Nothing
    Set oDBTypeConfig = Nothing
    Set jContext = Nothing
    Set oConnectMiddle = Nothing
    Exit Function
    
ErrHandler:
    Dim oErrors As New IMSErrorLog.JServerErrors
    oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
End Function

Public Sub GetWeightFromCatalog(SecName As String, SecStandard As String, ByRef NominalWt As Double)
Const METHOD = "GetWeightFromCatalog"
On Error GoTo ErrorHandler

    Dim pProfiles               As IJElements
    Dim oTrans4x4               As IJDT4x4
    
    Dim strCatlogDB         As String
    Dim CatalogDef          As Object

    Dim oDBTypeConfig As IJDBTypeConfiguration
    Dim jContext As IJContext
    Dim oConnectMiddle As IJDAccessMiddle
    Dim oCatResMgr As IUnknown
    
    Set jContext = GetJContext()
    Set oDBTypeConfig = jContext.GetService("DBTypeConfiguration")
    Set oConnectMiddle = jContext.GetService("ConnectMiddle")
    
    strCatlogDB = oDBTypeConfig.get_DataBaseFromDBType("Catalog")
    Set oCatResMgr = oConnectMiddle.GetResourceManager(strCatlogDB)

    Dim m_xService As SP3DStructGenericTools.CrossSectionServices
    Set m_xService = New SP3DStructGenericTools.CrossSectionServices

    Dim sectype As String
    m_xService.GetStructureCrossSectionDefinition oCatResMgr, _
                                                  SecStandard, sectype, _
                                                  SecName, CatalogDef
    Dim wt As Variant
    wt = 0#
    On Error Resume Next
    m_xService.GetCrossSectionAttributeValue CatalogDef, "UnitWeight", wt
    On Error GoTo ErrorHandler
    NominalWt = wt
    
    Set oDBTypeConfig = Nothing
    Set jContext = Nothing
    Set oConnectMiddle = Nothing
    Set oCatResMgr = Nothing
    Set m_xService = Nothing
    Set CatalogDef = Nothing
    
    Exit Sub
ErrorHandler:
    Dim oErrors As New IMSErrorLog.JServerErrors
    oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
End Sub

'***************************************************************************
'
' GetCenOfHRObjectWRTCP -- This method reads the CrossSection object and
'retrieves the Center point of action of the given Crosssection which can be
'used to form the cylinder for the Top Rail.
'
'***************************************************************************
Public Sub GetCenOfHRObjectWRTCP(ByVal strCSStandard As String, ByVal strCSName As String, ByVal nCSCP As Integer, ByVal dCSAngle As Double, ByRef dCenterX As Double, ByRef dCenterY As Double, ByRef dWidth As Double, ByRef dDepth As Double)
    Const METHOD = "GetCenOfHRObjectWRTCP"
    On Error GoTo ErrorHandler
    Dim w As Double, h As Double
    Dim CGx As Double, CGy As Double
    Dim xp As Double, yp As Double
    Dim x(1 To 4) As Double
    Dim y(1 To 4) As Double
    Dim nX(1 To 4) As Double
    Dim nY(1 To 4) As Double
    'Get the Width and Depth of Given CrossSection.
    GetCrossSecData strCSName, strCSStandard, w, h
    dDepth = h
    dWidth = w
    'Get the Centroid of the Given cross Section. This will be used only when
    'User Selects the Cardinal Points 10, 11, 12, 13 and 14
    CGx = GetCSAttribData(strCSName, strCSStandard, "ISTRUCTCrossSectionDesignProperties", "CentroidX")
    CGy = GetCSAttribData(strCSName, strCSStandard, "ISTRUCTCrossSectionDesignProperties", "CentroidY")
    
    'Get the Shear Center of the Given CrossSection. This will be used only when
    'User selects the Cardinal Point 15
    
    xp = GetCSAttribData(strCSName, strCSStandard, "IJUAL", "xp")
    yp = GetCSAttribData(strCSName, strCSStandard, "IJUAL", "yp")
    
    'Check weather the data is got from the CustomAttributes, if not then
    'Assume the points lie at distance half of either width or depth or both.
    If IsEmpty(CGx) Or CGx = 0# Then
        CGx = w / 2
    End If
    
    If IsEmpty(CGy) Or CGy = 0# Then
        CGy = h / 2
    End If
    
    If IsEmpty(xp) Or xp = 0# Then
        xp = w / 2
    End If
    
    If IsEmpty(yp) Or yp = 0# Then
        yp = h / 2
    End If
    
    'Based on the Cardinal Point set the 4 end points which shall be rotated
    If nCSCP = 1 Then 'Bottom Left
        x(1) = w / 2
        y(1) = h / 2
    ElseIf nCSCP = 2 Then 'Bottom Center
        x(1) = 0
        y(1) = h / 2
    ElseIf nCSCP = 3 Then 'Bottom Right
        x(1) = -w / 2
        y(1) = h / 2
    ElseIf nCSCP = 4 Then 'Center Left
        x(1) = w / 2
        y(1) = 0
    ElseIf nCSCP = 5 Then 'Center
        x(1) = 0
        y(1) = 0
    ElseIf nCSCP = 6 Then 'Center Right
        x(1) = -w / 2
        y(1) = 0
    ElseIf nCSCP = 7 Then 'Top Left
        x(1) = w / 2
        y(1) = -h / 2
    ElseIf nCSCP = 8 Then 'Top Center
        x(1) = 0
        y(1) = -h / 2
    ElseIf nCSCP = 9 Then 'Top Right
        x(1) = -w / 2
        y(1) = -h / 2
    ElseIf nCSCP = 10 Then 'Centroid
        x(1) = w / 2 - CGx
        y(1) = h / 2 - CGy
    ElseIf nCSCP = 11 Then 'Centroid Bottom
        x(1) = w / 2 - CGx
        y(1) = h / 2
    ElseIf nCSCP = 12 Then 'Centroid Left
        x(1) = w / 2
        y(1) = h / 2 - CGy
    ElseIf nCSCP = 13 Then 'Centroid Right
        x(1) = -w / 2
        y(1) = h / 2 - CGy
    ElseIf nCSCP = 14 Then 'Centroid Top
        x(1) = w / 2 - CGx
        y(1) = -h / 2
    ElseIf nCSCP = 15 Then 'Shear Center
        x(1) = w / 2 - xp
        y(1) = h / 2 - yp
    End If
    
    nX(1) = x(1) * Cos(dCSAngle) + y(1) * Sin(dCSAngle)
    nY(1) = y(1) * Cos(dCSAngle) - x(1) * Sin(dCSAngle)
    
    dCenterX = nX(1)
    dCenterY = nY(1)
    Exit Sub
    
ErrorHandler:
    Dim oErrors As New IMSErrorLog.JServerErrors
    oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
End Sub

Public Function GetSectionAngleWRTCurveParameterObject(oCurve As IJCurve, dParam As Double, dOldSectionAngle As Double) As Double
    Const METHOD = "GetSectionAngleWRTCurveParameterObject"
    On Error GoTo ErrorHandler
    
    Dim dStartRange As Double, dEndRange As Double
    Dim dX As Double, dY As Double, dZ As Double
    Dim dTanx As Double, dTany As Double, dTanz As Double
    Dim dTan2x As Double, dTan2y As Double, dTan2z As Double
    Dim oXAxisVector As IJDVector
    Dim oTangentVec As IJDVector
    Dim oNormalVec As IJDVector
    
    oCurve.ParamRange dStartRange, dEndRange
    
    oCurve.Evaluate dParam, dX, dY, dZ, dTanx, dTany, dTanz, dTan2x, dTan2y, dTan2z
    
    Set oXAxisVector = New DVector
    oXAxisVector.Set 1, 0, 0
    Set oTangentVec = New DVector
    oTangentVec.Set dTanx, dTany, 0
    Set oNormalVec = New DVector
    oNormalVec.Set 0, 0, 1#
    GetSectionAngleWRTCurveParameterObject = dOldSectionAngle + oXAxisVector.Angle(oTangentVec, oNormalVec)
    
    Exit Function
    
ErrorHandler:
    Dim oErrors As New IMSErrorLog.JServerErrors
    oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
End Function

'Create new output for the whole Curve.
Public Sub InitOperationOutput(pOC As IJDOutputCollection, name As String)
Const METHOD = "InitNewOutput"
    On Error GoTo ErrHandler
    
    Dim oRep As IMSSymbolEntities.IJDRepresentation
    Dim oOutputs As IMSSymbolEntities.IJDOutputs
    Dim oOutput As IMSSymbolEntities.IJDOutput
    
    Set oOutput = New DOutput
    'Use the Operation Aspect
    Set oRep = pOC.definition.IJDRepresentations.GetRepresentationByName("OperationRepresentation")
    Set oOutputs = oRep
    
    'Add Data to the Output
    oOutput.name = name
    oOutput.Description = name
    oOutputs.SetOutput oOutput
    oOutput.Reset
    Exit Sub
    
ErrHandler:
    Dim oErrors As New IMSErrorLog.JServerErrors
    oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
End Sub

Public Sub GetCrossSecData(SectionName As String, _
                            SectionStandard As String, _
                            ByRef sectionWidth As Double, _
                            ByRef sectionDepth As Double)
Const METHOD = "GetCrossSecData"
On Error GoTo ErrorHandler
                     
    sectionDepth = GetCSAttribData(SectionName, SectionStandard, "ISTRUCTCrossSectionDimensions", "Depth")
    sectionWidth = GetCSAttribData(SectionName, SectionStandard, "ISTRUCTCrossSectionDimensions", "Width")
                                     
    Exit Sub
    
ErrorHandler:
    Dim oErrors As New IMSErrorLog.JServerErrors
    oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
End Sub

'Calcualtes the SegmentLength to which the posts will be in active,
'No. of Intermediate posts that have to be placed, Actual clearance
'of the posts when post at turn is set to false.-- Satish N Kota [April 24, 2003]
Public Sub CalculateSegmentDetails(ByRef SegLength As Double, ByRef NoIntermediatePosts As Integer, ByVal IsPostAtTurn As Boolean, ByVal IsSloped As Boolean, ByVal SlopedSegmentMaxSpacing As Double, ByVal SegmentMaxSpacing As Double, ByVal MaxTurnClearanceDist As Double, ByVal MinTurnClearanceDist As Double, ByRef ActualClearance As Double, ByVal eSegType As eSEGMENTTYPE, Optional ByVal bnextprevslopeCol As Collection)
    Const METHOD = "CalculateSegmentDetails"
    Const INCHES = 39.3701
    Dim NewSegLen As Double
    Dim TempSegLength As Double
    Dim TempSlopedSegMaxSpacing As Double
    Dim TempSegMaxSpacing As Double
    Dim TempMinClearance As Double
    Dim TempMaxClearance As Double
    Dim lTempPosts As Integer
    Dim dSpacingValue As Double
    Dim bNoBeginPostClearance As Boolean
    Dim bNoEndPostClearance As Boolean
    On Error GoTo ErrorHandler
    
    Dim MaxSegLength As Double
    Dim dMaxSpacingValue As Double
    Dim SegMaxSpacing As Double
    
    If Not bnextprevslopeCol Is Nothing Then
        bNoBeginPostClearance = bnextprevslopeCol.Item(1)
        bNoEndPostClearance = bnextprevslopeCol.Item(2)
    End If
    
    NewSegLen = SegLength
    'Convert all the values into Inches, This is done to take care of
    'Creating the spacing to the nearest 1" value and since the symbol
    'units are always in metric units conversion is required.
    TempSegLength = Round(SegLength * INCHES, 2)
    TempSlopedSegMaxSpacing = Round(SlopedSegmentMaxSpacing * INCHES, 2)
    TempSegMaxSpacing = Round(SegmentMaxSpacing * INCHES, 2)
    TempMinClearance = Round(MinTurnClearanceDist * INCHES, 2)
    TempMaxClearance = Round(MaxTurnClearanceDist * INCHES, 2)
             
    'Check if the post is at turn or not
    If IsPostAtTurn = False Then
        'Place posts at a distance which is between the MAX and MIN clearance
        'limits.
        If eSegType = SEGMENT_BEGIN_TYPE Or eSegType = SEGMENT_END_TYPE Then
            MaxSegLength = TempSegLength - TempMinClearance
            TempSegLength = TempSegLength - TempMaxClearance
        Else
            'not end segments, hence take clearance on both ends of the segment.
            
            If bNoBeginPostClearance = True Or bNoEndPostClearance = True Then
                MaxSegLength = TempSegLength - TempMinClearance
                TempSegLength = TempSegLength - (TempMaxClearance)
            Else
                MaxSegLength = TempSegLength - 2 * TempMinClearance
                TempSegLength = TempSegLength - (2 * TempMaxClearance)
            End If
        End If
    End If
    
    'Check for weather the segment is in a sloped line or not...and calculate
    'the number of posts on that basis.
    If IsSloped Then
        SegMaxSpacing = TempSlopedSegMaxSpacing
        If TempSegLength > TempSlopedSegMaxSpacing Then
            If Abs(TempSegLength / TempSlopedSegMaxSpacing) > Round(Abs(TempSegLength / TempSlopedSegMaxSpacing), 0) Then
                lTempPosts = Round(Abs(TempSegLength / TempSlopedSegMaxSpacing)) + 1  ' need to verify this
            Else
                lTempPosts = Round(Abs(TempSegLength / TempSlopedSegMaxSpacing))  ' need to verify this
            End If
        End If
    Else
        SegMaxSpacing = TempSegMaxSpacing
        If TempSegLength > TempSegMaxSpacing Then
            If Abs(TempSegLength / TempSegMaxSpacing) > Round(Abs(TempSegLength / TempSegMaxSpacing), 0) Then
                lTempPosts = Round(Abs(TempSegLength / TempSegMaxSpacing)) + 1  ' need to verify this
            Else
                lTempPosts = Round(Abs(TempSegLength / TempSegMaxSpacing))  ' need to verify this
            End If
        End If
    End If
    
    'if the post at turn is not allowed then calculate the actual spacing
    'between the posts and decrease the spacing of first or last post from the
    'segment ends accordingly....
    If IsPostAtTurn = False Then
        If lTempPosts <> 0 Then
            dSpacingValue = TempSegLength / lTempPosts
            If dSpacingValue > Round(dSpacingValue, 0) Then
                dSpacingValue = Round(dSpacingValue, 0) + 1
            Else
                dSpacingValue = Round(dSpacingValue, 0)
             End If
            dMaxSpacingValue = Round(MaxSegLength / lTempPosts, 2)
            'The clearance between the first or last post from the turn.
            'honor the minimum clearance if possible
            If dMaxSpacingValue > TempSegMaxSpacing Then
                ActualClearance = NewSegLen - (TempSegMaxSpacing * lTempPosts / INCHES)
                NewSegLen = TempSegMaxSpacing * lTempPosts / INCHES
            Else
                ActualClearance = NewSegLen - (dMaxSpacingValue * lTempPosts / INCHES)
                NewSegLen = dMaxSpacingValue * lTempPosts / INCHES
            End If
                    
            'If the actualclearance calculated is less than the Minimun
            'allowable limit, then set the minimum allowable limit to be the
            'clearance and adjust the spacing accordingly
            If eSegType = SEGMENT_MID_TYPE Then
                If ActualClearance < 2 * MinTurnClearanceDist Then
                    ActualClearance = 2 * MinTurnClearanceDist
                    NewSegLen = SegLength - ActualClearance
                End If
            Else
                If ActualClearance < MinTurnClearanceDist Then
                    ActualClearance = MinTurnClearanceDist
                    NewSegLen = SegLength - ActualClearance
                End If
            End If
            'Hence calculate the actual segment length between which
            'the posts will be placed.
'            segLength = NewSegLen  'TR#55631
        Else
            'This happens when the number of posts from the max spacing is 0
            'i.e. the segment length is less than the Max segment spacing length.
            'honor the minimum clearance if possible
            If MaxSegLength > SegMaxSpacing Then
                ActualClearance = SegLength - (SegMaxSpacing / INCHES)
            Else
                ActualClearance = SegLength - (MaxSegLength / INCHES)
            End If
'            segLength = TempSegLength / INCHES 'TR#55631
        End If
    End If
    
    If ActualClearance >= SegLength Then
        ActualClearance = 0
    End If
    
    NoIntermediatePosts = lTempPosts
    Exit Sub
ErrorHandler:
    Dim oErrors As New IMSErrorLog.JServerErrors
    oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
End Sub

Public Sub InitNewOutput(pOC As IJDOutputCollection, name As String)
Const METHOD = "InitNewOutput"
    On Error GoTo ErrHandler
    
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
    
ErrHandler:
    Dim oErrors As New IMSErrorLog.JServerErrors
    oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
End Sub
'TR#51383-changed input vector from SegVec to vecOv so that for sloped segments
'arc points will be calculated accordingly. Also modified so that it works correctly for
' orientation="Perpendicular to.." case with new input vector
Public Function CreateCirEndTreatment(SegCurve As IJCurve, _
                                       ResourceManager As Object, _
                                       PostStartX As Double, _
                                       PostStartY As Double, _
                                       PostStartZ As Double, _
                                       MidRailDist As Double, _
                                       TotalHt As Double, _
                                       PostVec As DVector, _
                                       bIsbegin As Boolean, dToprailDepth As Double, Optional dRadius As Double) As ComplexString3d
                                       
Const METHOD = "CreateCirEndTreatment"
On Error GoTo ErrorHandler
  
    Set CreateCirEndTreatment = Nothing
    Dim StartPar As Double, EndPar As Double
    Dim xPt As Double, yPt As Double, zPt As Double
    Dim vtanx As Double, vtany As Double, vtanz As Double
    Dim vTan2X As Double, vTan2Y As Double, vTan2Z As Double
    SegCurve.ParamRange StartPar, EndPar
    If bIsbegin Then
        SegCurve.Evaluate StartPar, xPt, yPt, zPt, vtanx, vtany, vtanz, vTan2X, vTan2Y, vTan2Z
    Else
        SegCurve.Evaluate EndPar, xPt, yPt, zPt, vtanx, vtany, vtanz, vTan2X, vTan2Y, vTan2Z
        
    End If
    'The implementation for circular treatment on sloped path is only applicable for "Always Vertical"
    If Abs(vtanz) > dtol And Abs(PostVec.x) < dtol And Abs(PostVec.y) < dtol Then
        Set CreateCirEndTreatment = CreateSlopedCirEndTreatment(SegCurve, ResourceManager, PostStartX, PostStartY, PostStartZ, _
                                           MidRailDist, TotalHt, PostVec, bIsbegin, dToprailDepth, dRadius)
        Exit Function
    End If
    
    Dim Pt1 As DPosition
    Dim postPt As DPosition
    Dim arcPt As DPosition
    Dim vecOV As DVector
    Set vecOV = New DVector
    vecOV.Set PostVec.x, PostVec.y, PostVec.z
    Set Pt1 = New DPosition
    Set postPt = New DPosition
    Set arcPt = New DPosition
    postPt.Set PostStartX, PostStartY, PostStartZ
    
    Dim ht1 As Double
    If dRadius > dtol Then
        ht1 = dRadius
    Else
        ht1 = (TotalHt - MidRailDist) / 6   ' assumption 1/6 of the total ht = arc radius
    
        'If the above assuption is not a good one, use the following formular to get arc radius
        'This value should be consistent with the value of dCirTOffset in PhysicalRepresentation
        If dToprailDepth > (TotalHt - MidRailDist) Then
            ht1 = (TotalHt - MidRailDist) / 2
        ElseIf ht1 < dToprailDepth / 2 Then
            ht1 = dToprailDepth / 4 + (TotalHt - MidRailDist) / 4
        End If
    End If
    
    Dim ParamOnSegment As Double
    
'    ParamOnSegment = StartPar + ht1 '*  (EndPar - StartPar) / SegCurve.length = 1 don't need to do actually
        ParamOnSegment = (EndPar - StartPar) * ht1 / SegCurve.Length ' TR#70016
    
    Dim Arc1X1 As Double, Arc1Y1 As Double, Arc1Z1 As Double
    Dim Arc1X2 As Double, Arc1Y2 As Double, Arc1Z2 As Double
    Dim Arc1X3 As Double, Arc1Y3 As Double, Arc1Z3 As Double
    
    Dim Arc2X1 As Double, Arc2Y1 As Double, Arc2Z1 As Double
    Dim Arc2X2 As Double, Arc2Y2 As Double, Arc2Z2 As Double
    Dim Arc2X3 As Double, Arc2Y3 As Double, Arc2Z3 As Double
    
    If m_GeomFactory Is Nothing Then
        Set m_GeomFactory = New GeometryFactory
    End If
    
    If bIsbegin = True Then
        SegCurve.Position StartPar + ParamOnSegment, xPt, yPt, zPt 'TR#70016
    Else
        SegCurve.Position (EndPar - ParamOnSegment), xPt, yPt, zPt
    End If
    
    Pt1.Set xPt, yPt, zPt 'TR#51383
    
        vecOV.Length = MidRailDist
    Set arcPt = Pt1.Offset(vecOV)
    Arc1X1 = arcPt.x      'xPt
    Arc1Y1 = arcPt.y      'yPt
    Arc1Z1 = arcPt.z      'zPt + MidRailDist
    
        vecOV.Length = MidRailDist + ht1
    Set arcPt = Pt1.Offset(vecOV)
    Arc1X2 = arcPt.x      'Arc1X1
    Arc1Y2 = arcPt.y      'Arc1Y1
    Arc1Z2 = arcPt.z      ' Arc1Z1 + ht1
    
        vecOV.Length = MidRailDist + ht1
    Set arcPt = postPt.Offset(vecOV)
    
    Arc1X3 = arcPt.x      'PostStartX
    Arc1Y3 = arcPt.y      'PostStartY
    Arc1Z3 = arcPt.z      'PostStartZ + MidRailDist + ht1
    
        vecOV.Length = TotalHt
    Set arcPt = Pt1.Offset(vecOV)
    Arc2X1 = arcPt.x        'xPt
    Arc2Y1 = arcPt.y        'yPt
    Arc2Z1 = arcPt.z        'zPt + TotalHt
    
        vecOV.Length = TotalHt - ht1
    Set arcPt = Pt1.Offset(vecOV)
    Arc2X2 = arcPt.x        'Arc2X1
    Arc2Y2 = arcPt.y        'Arc2Y1
    Arc2Z2 = arcPt.z        'Arc2Z1 - ht1
    
        vecOV.Length = TotalHt - ht1
    Set arcPt = postPt.Offset(vecOV)
    Arc2X3 = arcPt.x        'PostStartX
    Arc2Y3 = arcPt.y        'PostStartY
    Arc2Z3 = arcPt.z       'PostStartZ + TotalHt - ht1

    Dim iElements As IJElements
    Set iElements = New JObjectCollection ' IMSElements.DynElements
    
    Dim oLine As IngrGeom3D.Line3d
    Dim oarc1 As IngrGeom3D.Arc3d, oArc2 As IngrGeom3D.Arc3d
    
    ' create an arc with arc2pt1 as start so that the cross section direction will be consistent with the toprail
    Set oarc1 = m_GeomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, _
                                                        Arc2X2, Arc2Y2, Arc2Z2, _
                                                        Arc2X1, Arc2Y1, Arc2Z1, _
                                                        Arc2X3, Arc2Y3, Arc2Z3)
    If Not oarc1 Is Nothing Then
        iElements.Add oarc1
        oarc1.GetEndPoint Arc2X3, Arc2Y3, Arc2Z3
    End If

    Set oLine = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                                    Arc2X3, Arc2Y3, Arc2Z3, _
                                                    Arc1X3, Arc1Y3, Arc1Z3)
    If Not oLine Is Nothing Then iElements.Add oLine
    Set oArc2 = m_GeomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, _
                                                        Arc1X2, Arc1Y2, Arc1Z2, _
                                                        Arc1X3, Arc1Y3, Arc1Z3, _
                                                        Arc1X1, Arc1Y1, Arc1Z1)
    If Not oArc2 Is Nothing Then iElements.Add oArc2
    
    Dim oCmpx As ComplexString3d
    
    Set oCmpx = m_GeomFactory.ComplexStrings3d.CreateByCurves(Nothing, iElements)

    Set CreateCirEndTreatment = oCmpx
    
    Set oarc1 = Nothing
    Set oArc2 = Nothing
    Set oLine = Nothing
    Set oCmpx = Nothing
    Set iElements = Nothing
    Set Pt1 = Nothing
    Set postPt = Nothing
    Set arcPt = Nothing
    
    Exit Function
    
ErrorHandler:
    Dim oErrors As New IMSErrorLog.JServerErrors
    oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
End Function

'This function will return projected handrail curve. Input is handaril path curve. This can be used to get curve at topraul/midrail/toe
' plate location by sending appropriate height value. This will handle both orientations of handrail (always vert. & perpendicular to path)
'Currently it doesn't work for orientation=perpendicular to path. It is not implemented for
'mid-rail, top-rail of any type of handrail. It has to be implemented at all the places including this function
'TR#51383- modified to handle "Perpendicular to slope" case
Public Sub CreateProjectedHRCurve(ByRef pCurve, _
                                ByVal pComplex As ComplexString3d, _
                                Orientation As Integer, _
                                Height As Double)
                                
Const METHOD = "CreateProjectedHRCurve"
On Error GoTo ErrorHandler

    Dim oTrans4x4       As IJDT4x4
    Dim oVector         As IJDVector
    Dim oSegments       As IJElements

    If m_GeomFactory Is Nothing Then
        Set m_GeomFactory = New GeometryFactory
    End If
    
    If Orientation And ComplexstringcontainsArc(pComplex) = False Then
        Set pCurve = GetPathByOffset(pComplex, Height, 0) 'TR#51383- workaround for GeometryServices.CreateByOffset as it doesn't work for non-planar path
    Else
        Set oVector = New DVector
        oVector.Set 0, 0, Height
        
        Set oTrans4x4 = New DT4x4
        oTrans4x4.LoadIdentity
        oTrans4x4.Translate oVector
    
        pComplex.GetCurves oSegments
        
        Set pCurve = New ComplexString3d
        Set pCurve = m_GeomFactory.ComplexStrings3d.CreateByCurves(Nothing, oSegments)
        
        pCurve.Transform oTrans4x4
    End If
    
        Set oSegments = Nothing
        Set oTrans4x4 = Nothing
        Set oVector = Nothing
    
    Exit Sub
    
ErrorHandler:
    Dim oErrors As New IMSErrorLog.JServerErrors
    oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
End Sub

' This is similar to CreateProjectedHRCurve(). Difference is that here it takes input as vector
' and offsets the path curve in the direction of vector by distance equal to vector length.
' Even though it is similar above function is written separately to handle HR with orientation
' always vertical & perpendicular to path.
Public Sub CreateProjectedTopRailCurve(ByRef pCurve, _
                                ByVal pComplex As ComplexString3d, _
                                oVector As DVector)
                                
Const METHOD = "CreateProjectedTopRailCurve"
On Error GoTo ErrorHandler

    Dim oTrans4x4       As IJDT4x4
    Dim oSegments       As IJElements
    
    
    If m_GeomFactory Is Nothing Then
        Set m_GeomFactory = New GeometryFactory
    End If
    
    Set oTrans4x4 = New DT4x4
    oTrans4x4.LoadIdentity
    oTrans4x4.Translate oVector

    pComplex.GetCurves oSegments
    
    Set pCurve = New ComplexString3d
    Set pCurve = m_GeomFactory.ComplexStrings3d.CreateByCurves(Nothing, oSegments)
    
    pCurve.Transform oTrans4x4
 
    Set oSegments = Nothing
    Set oTrans4x4 = Nothing
    
    Exit Sub
    
ErrorHandler:
    Dim oErrors As New IMSErrorLog.JServerErrors
    oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
End Sub

'***************************************************************************
'
' TR45745:
' GethDeltaForTopRail -- This method gets the hDelta value for toprail based on
' CP & orientation of top rail. This will be used to place top rail such that its
' top always matches the height of handrail. This works only for orientation
' 0, 90, 180, 270 which are most common. For other orientations. more detailed
' calculations are required.
'
'***************************************************************************
Public Sub GethDeltaForTopRail(ByVal strCSStandard As String, ByVal strCSName As String, ByVal nCSCP As Integer, ByVal dCSAngle As Double, ByRef dhDelta As Double, ByRef w As Double, ByRef h As Double)
    Const METHOD = "GethDeltaForTopRail"
    Dim CGx As Double, CGy As Double
    Dim xp As Double, yp As Double
    Dim x1 As Double, x2 As Double
    Dim y1 As Double, y2 As Double
    Dim nY As Double, ny1 As Double
    
    'Get the Centroid of the Given cross Section. This will be used only when
    'User Selects the Cardinal Points 10, 11, 12, 13 and 14
    CGx = GetCSAttribData(strCSName, strCSStandard, "ISTRUCTCrossSectionDesignProperties", "CentroidX")
    CGy = GetCSAttribData(strCSName, strCSStandard, "ISTRUCTCrossSectionDesignProperties", "CentroidY")
    
    'Get the Shear Center of the Given CrossSection. This will be used only when
    'User selects the Cardinal Point 15
    
    xp = GetCSAttribData(strCSName, strCSStandard, "IJUAL", "xp")
    yp = GetCSAttribData(strCSName, strCSStandard, "IJUAL", "yp")
    
    'Check weather the data is got from the CustomAttributes, if not then
    'Assume the points lie at distance half of either width or depth or both.
    If IsEmpty(CGx) Or CGx = 0# Then
        CGx = w / 2
    End If
    
    If IsEmpty(CGy) Or CGy = 0# Then
        CGy = h / 2
    End If
    
    If IsEmpty(xp) Or xp = 0# Then
        xp = w / 2
    End If
    
    If IsEmpty(yp) Or yp = 0# Then
        yp = h / 2
    End If
    
    
    If nCSCP = 1 Then 'Bottom Left
        x1 = w
        x2 = 0#
        y1 = h
        y2 = 0#
    ElseIf nCSCP = 2 Then 'Bottom Center
        x1 = w / 2
        x2 = -w / 2
        y1 = h
        y2 = 0#
    ElseIf nCSCP = 3 Then 'Bottom Right
        x1 = 0#
        x2 = -w
        y1 = h
        y2 = 0#
    ElseIf nCSCP = 4 Then 'Center Left
        x1 = w
        x2 = 0#
        y1 = h / 2
        y2 = -h / 2
    ElseIf nCSCP = 5 Then 'Center
        x1 = w / 2
        x2 = -w / 2
        y1 = h / 2
        y2 = -h / 2
    ElseIf nCSCP = 6 Then 'Center Right
        x1 = 0#
        x2 = -w
        y1 = h / 2
        y2 = -h / 2
    ElseIf nCSCP = 7 Then 'Top Left
        x1 = w
        x2 = 0#
        y1 = 0#
        y2 = -h
    ElseIf nCSCP = 8 Then 'Top Center
        x1 = w / 2
        x2 = -w / 2
        y1 = 0#
        y2 = -h
    ElseIf nCSCP = 9 Then 'Top Right
        x1 = 0#
        x2 = -w
        y1 = 0#
        y2 = -h
    ElseIf nCSCP = 10 Then 'Centroid
        x1 = w - CGx
        x2 = -CGx
        y1 = h - CGy
        y2 = -CGy
    ElseIf nCSCP = 11 Then 'Centroid Bottom
        x1 = w - CGx
        x2 = -CGx
        y1 = h
        y2 = 0#
    ElseIf nCSCP = 12 Then 'Centroid Left
        x1 = w
        x2 = 0#
        y1 = h - CGy
        y2 = -CGy
    ElseIf nCSCP = 13 Then 'Centroid Right
        x1 = 0#
        x2 = -w
        y1 = h - CGy
        y2 = -CGy
    ElseIf nCSCP = 14 Then 'Centroid Top
        x1 = w - CGx
        x2 = -CGx
        y1 = 0#
        y2 = -h
    ElseIf nCSCP = 15 Then 'Shear Center
        x1 = w - xp
        x2 = -xp
        y1 = h - yp
        y2 = -yp
    End If
    
    nY = y1 * Cos(dCSAngle) - x1 * Sin(dCSAngle)
    ny1 = y2 * Cos(dCSAngle) - x2 * Sin(dCSAngle)
    If nY < ny1 Then
        dhDelta = ny1
    Else
        dhDelta = nY
    End If

End Sub

'TR#51383- Copied this from GenericSpanAndOrigin.bas as it is
Public Function GetIntersectPoint(m_pTopEdge As Object, m_pRefEdge As Object) As DPosition
Const METHOD = "GetIntersectPoint"
On Error GoTo ErrorHandler
    
    Set GetIntersectPoint = New DPosition
    Dim x2 As Double, y2 As Double, z2 As Double
    Dim xSt As Double, ySt As Double, zSt As Double
    Dim xEn As Double, yEn As Double, zEn As Double
    Dim ClosestParam As Double
    Dim ParmU As Double, ParmV As Double
    Dim oTopface As IJPlane
    Dim oTopSurf As IJSurface
    Dim oTopCurve As IJCurve
    Dim oTopEdge As IJLine
    Dim oRefCurve As IJCurve
    Dim oRefEdge As IJLine
    Dim oRefFace As IJPlane
    Dim oRefSurf As IJSurface
    

    If TypeOf m_pTopEdge Is IJLine Then
        Set oTopEdge = m_pTopEdge
    End If
    If TypeOf m_pTopEdge Is IJCurve Then
        Set oTopCurve = m_pTopEdge
    End If
    If oTopCurve Is Nothing Then
        Set oTopCurve = oTopEdge
    End If
    If TypeOf m_pTopEdge Is IJPlane Then
        Set oTopface = m_pTopEdge
    End If
    If TypeOf m_pTopEdge Is IJSurface Then
        Set oTopSurf = m_pTopEdge
    End If
    If oTopSurf Is Nothing Then
        Set oTopSurf = oTopface
    End If
    
    If TypeOf m_pRefEdge Is IJLine Then
        Set oRefEdge = m_pRefEdge
    End If
    If TypeOf m_pRefEdge Is IJCurve Then
        Set oRefCurve = m_pRefEdge
    End If
    If oRefCurve Is Nothing Then
        Set oRefCurve = oRefEdge
    End If
    If TypeOf m_pRefEdge Is IJPlane Then
        Set oRefFace = m_pRefEdge
    End If
    If TypeOf m_pRefEdge Is IJSurface Then
        Set oRefSurf = m_pRefEdge
    End If
    If oRefSurf Is Nothing Then
        Set oRefSurf = oRefFace
    End If
        
    If Not oTopSurf Is Nothing And Not oRefSurf Is Nothing Then
        GoTo ErrorHandler
    End If
    
    If m_pRefEdge Is Nothing Then
        If Not oTopCurve Is Nothing Then
            oTopCurve.EndPoints xSt, ySt, zSt, xEn, yEn, zEn
            GetIntersectPoint.Set xSt, ySt, zSt
            Exit Function
        End If
    Else
        If oRefSurf Is Nothing And oRefCurve Is Nothing Then   ' i know Top will be curve ' kludge as gridplanes dont support IJsurface
            Set GetIntersectPoint = FindIntersection(oRefFace, oTopCurve)
            Exit Function
        End If
    End If
    
    If Not oRefCurve Is Nothing Or oTopSurf Is Nothing And oRefSurf Is Nothing Then
        oRefCurve.EndPoints xSt, ySt, zSt, xEn, yEn, zEn
        If Not oTopCurve Is Nothing Then
            Dim x1 As Double, y1 As Double, z1 As Double, Mindist As Double
            oTopCurve.DistanceBetween oRefCurve, Mindist, x1, y1, z1, x2, y2, z2
            oTopCurve.Parameter x2, y2, z2, ClosestParam
            oTopCurve.Position ClosestParam, x2, y2, z2
        ElseIf Not oTopSurf Is Nothing Then
            oTopSurf.Parameter xEn, yEn, zEn, ParmU, ParmV
            oTopSurf.Position ParmU, ParmV, x2, y2, z2
        End If
    ElseIf Not oRefSurf Is Nothing Then
        oTopCurve.EndPoints xSt, ySt, zSt, xEn, yEn, zEn
        oRefSurf.Parameter xSt, ySt, zSt, ParmU, ParmV
        oRefSurf.Position ParmU, ParmV, x2, y2, z2
    End If
    GetIntersectPoint.Set x2, y2, z2
    
    Exit Function
    
ErrorHandler:
    Dim oErrors As New IMSErrorLog.JServerErrors
    oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
End Function
'TR#51383- Copied this from GenericSpanAndOrigin.bas as it is
Private Function FindIntersection(oPlane As IJPlane, oLine As IJCurve) As DPosition
    On Error GoTo ErrHandler
    Dim v1 As DVector
    Dim v2 As DVector
    Dim StartPoint As DPosition
    Dim EndPoint As DPosition
    Dim PlanePoint As DPosition
    Dim Dot_prod As Double


    Set v1 = New DVector
    Set v2 = New DVector
    Set StartPoint = New DPosition
    Set EndPoint = New DPosition
    Set PlanePoint = New DPosition
    Dim x As Double, y As Double, z As Double
    Dim x1 As Double, y1 As Double, z1 As Double
    
    oLine.EndPoints x, y, z, x1, y1, z1
    StartPoint.Set x, y, z
    EndPoint.Set x1, y1, z1
    oPlane.GetRootPoint x, y, z
    PlanePoint.Set x, y, z
    
    v1.Set EndPoint.x - StartPoint.x, EndPoint.y - StartPoint.y, EndPoint.z - StartPoint.z
    v2.Set PlanePoint.x - StartPoint.x, PlanePoint.y - StartPoint.y, PlanePoint.z - StartPoint.z

        v1.Length = 1
    Dot_prod = v1.Dot(v2)

    Dim Intersection As DPosition
    Set Intersection = New DPosition

    Intersection.x = StartPoint.x + (Dot_prod * v1.x)
    Intersection.y = StartPoint.y + (Dot_prod * v1.y)
    Intersection.z = StartPoint.z + (Dot_prod * v1.z)
    Set FindIntersection = Intersection


    Set v1 = Nothing
    Set v2 = Nothing
    Set StartPoint = Nothing
    Set EndPoint = Nothing
    Set PlanePoint = Nothing
    Set Intersection = Nothing

    Exit Function
ErrHandler:
    Dim oErrors As New IMSErrorLog.JServerErrors
    oErrors.Add Err.Number, "FindIntersection", Err.Description
    Err.Raise E_FAIL
End Function

'TR#51383- This function is written as workaround for GeometryServices.CreateByOffset
'because it doesn't work correctly for non-planar path
'TR#53722- Last argument added which indicates the direction in which offset is to be applied
' 0- In vertical direction. Required to get toprail location for "Perpendicular to path" condition
'1 - Horizontal direction normal to handrail path segment. Required while applying horizontal offset to HR
' This function works only when path doesn't contain arcs

Public Function GetPathByOffset(pCurve As ComplexString3d, Offset As Double, offsetDir As Integer) As ComplexString3d
Const METHOD = "GetPathByOffset"
On Error GoTo ErrorHandler

    Dim vZ As DVector, vTan As DVector, vNorm As DVector, vOrient As DVector
    Dim pSegments As IJElements
    Dim pNewSegment As IJElements
    Dim segNo As Integer
    Dim pline1 As IJLine, pline2 As IJLine
    Dim sPt As DPosition, ePt As DPosition
    Dim sX As Double, sY As Double, sZ As Double
    Dim eX As Double, eY As Double, eZ As Double
    Dim stIntPtX As Double, stIntPtY As Double, stIntPtZ As Double
    Dim intPt As DPosition
    Dim dirVec As DVector ' TR#53722
    
    If m_GeomFactory Is Nothing Then
        Set m_GeomFactory = New GeometryFactory
    End If

    Set pNewSegment = New JObjectCollection
    Set intPt = New DPosition
    Set sPt = New DPosition
    Set ePt = New DPosition
    Set vZ = New DVector
    Set vTan = New DVector
    Set vNorm = New DVector
    Set vOrient = New DVector
    Set dirVec = New DVector
    
    vZ.Set 0#, 0#, 1#
    If offsetDir = 0 Then ' TR#53722- offset curve in vertical direction used for "Perpendicular to slope" case
        dirVec.Set 0#, 0#, 1#
    End If
    
    pCurve.GetCurves pSegments
    
    For segNo = 1 To pSegments.Count
        ' get current segment line and then offset it with given value and based on direction input
        
        Set pline1 = pSegments(segNo)
        pline1.GetStartPoint sX, sY, sZ
        pline1.GetEndPoint eX, eY, eZ
        
        vTan.Set (eX - sX), (eY - sY), (eZ - sZ)
        If offsetDir = 1 Then ' get direction vector which is normal to segment and in horizontal plane. Used for applying hz offset
            Set dirVec = vZ.Cross(vTan)
        End If
        Set vNorm = vTan.Cross(dirVec)
        Set vOrient = vNorm.Cross(vTan)
            vOrient.Length = Offset
    
        sPt.Set sX, sY, sZ
        ePt.Set eX, eY, eZ
        Set sPt = sPt.Offset(vOrient)
        Set ePt = ePt.Offset(vOrient)
        sX = sPt.x
        sY = sPt.y
        sZ = sPt.z
        eX = ePt.x
        eY = ePt.y
        eZ = ePt.z
        
        ' for 1st segment this is start point for subsequest segments it will be intersetion point with previous segment
        If segNo = 1 Then
            stIntPtX = sX
            stIntPtY = sY
            stIntPtZ = sZ
        End If
        
        Set pline1 = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, stIntPtX, stIntPtY, stIntPtZ, eX, eY, eZ)
        
        ' get next segment line and then offset it with given value and based on direction input.
        ' except for last segment
        
        If segNo <> pSegments.Count Then
            If TypeOf pSegments(segNo + 1) Is IJLine Then
                Set pline2 = pSegments(segNo + 1)
                pline2.GetStartPoint sX, sY, sZ
                pline2.GetEndPoint eX, eY, eZ
                Set pline2 = Nothing
                
                vTan.Set (eX - sX), (eY - sY), (eZ - sZ)
                
                If offsetDir = 1 Then ' get direction vector which is normal to segment and in horizontal plane
                    Set dirVec = vZ.Cross(vTan)
                End If
                
                Set vNorm = vTan.Cross(dirVec)
                Set vOrient = vNorm.Cross(vTan)
                    vOrient.Length = Offset
            
                sPt.Set sX, sY, sZ
                ePt.Set eX, eY, eZ
                Set sPt = sPt.Offset(vOrient)
                Set ePt = ePt.Offset(vOrient)
                sX = sPt.x
                sY = sPt.y
                sZ = sPt.z
                eX = ePt.x
                eY = ePt.y
                eZ = ePt.z
                Set pline2 = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, sX, sY, sZ, eX, eY, eZ)
                
                ' now we have two lines corresponding to segNo & segNo+1. Find its intersection point
                'and reconstruct line corresponding to segNo from start point to intersection point
                Set intPt = GetIntersectPoint(pline1, pline2)
                
                eX = intPt.x
                eY = intPt.y
                eZ = intPt.z
            End If
        End If
        
        Set pline1 = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, stIntPtX, stIntPtY, stIntPtZ, eX, eY, eZ)
        
        ' intersection point is going to be as starting point for line for next segNo
        stIntPtX = eX
        stIntPtY = eY
        stIntPtZ = eZ
        
        pNewSegment.Add pline1
    Next segNo

    'create new complexstring3d based on new segment collection
    Set GetPathByOffset = m_GeomFactory.ComplexStrings3d.CreateByCurves(Nothing, pNewSegment)
    
    Set pNewSegment = Nothing
    Set vZ = Nothing
    Set vTan = Nothing
    Set vNorm = Nothing
    Set vOrient = Nothing
    Set pline1 = Nothing
    Set pline2 = Nothing
    Set intPt = Nothing
    Set ePt = Nothing
    Set sPt = Nothing
    Set dirVec = Nothing
    
Exit Function
    
ErrorHandler:
    Dim oErrors As New IMSErrorLog.JServerErrors
    oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
End Function


'TR#51383- This routine will get modified handrail segments by comparing with top rail segment
'this is required when orientation is "Perpendicular to slope". For this case post location are
' not decided just by original handrail segment but based on top rail segment. Refer to specs for details
Public Sub GetModifiedHRSegments(ByRef pHRsegments As IJElements, pTopSegments As IJElements, Height As Double)
    Const METHOD = "GetModifiedHRSegments"
    On Error GoTo ErrorHandler
    Dim SegIndex As Integer
    Dim pTopLine As IJLine
    Dim pBotLine As IJLine
    Dim pArc As IJCurve
    Dim tSt As DPosition
    Dim tEnd As DPosition
    Dim bSt As DPosition
    Dim bEnd As DPosition
    Dim bInterPt As DPosition
    Dim vZ As DVector, vTan As DVector, vNorm As DVector, vOrient As DVector
    Dim x As Double, y As Double, z As Double
    Dim pComplex As ComplexString3d
    Dim newHRsegments As IJElements
    Dim sX As Double, sY As Double, sZ As Double
    Dim eX As Double, eY As Double, eZ As Double
    
    Set pComplex = New ComplexString3d
    
    If m_GeomFactory Is Nothing Then
        Set m_GeomFactory = New GeometryFactory
    End If
    
    For SegIndex = 1 To pHRsegments.Count
        Set pBotLine = pHRsegments(SegIndex)
    
        If pBotLine Is Nothing Then  'If arc then don't modify it. Currently arcs are not handled for "Perpend.." case
            Set pArc = pHRsegments(SegIndex)
            If newHRsegments Is Nothing Then
                pComplex.AddCurve pArc, False
                pComplex.GetCurves newHRsegments
            Else
                newHRsegments.Add pArc
            End If
            Set pArc = Nothing
        Else
            Set pTopLine = pTopSegments(SegIndex)
        
            Set tSt = New DPosition
            Set tEnd = New DPosition
            Set bSt = New DPosition
            Set bEnd = New DPosition
            Set bInterPt = New DPosition
        
            pTopLine.GetStartPoint x, y, z  'start point of top line
            tSt.Set x, y, z
            pTopLine.GetEndPoint x, y, z    'End point of top line
            tEnd.Set x, y, z
            pBotLine.GetStartPoint x, y, z  'start point of Bottom line
            bSt.Set x, y, z
            pBotLine.GetEndPoint x, y, z    'End point of Bottom line
            bEnd.Set x, y, z
        
            Set vZ = New DVector
            Set vTan = New DVector
            Set vNorm = New DVector
            Set vOrient = New DVector
    
            vZ.Set 0#, 0#, 1#
            vTan.Set (bEnd.x - bSt.x), (bEnd.y - bSt.y), (bEnd.z - bSt.z)
            Set vNorm = vZ.Cross(vTan)
            Set vOrient = vNorm.Cross(vTan)
                vOrient.Length = Height
        
            'Project start point of top line on bottom. If it is within exisiting bottom line then set this as star point
            Set bInterPt = tSt.Offset(vOrient)
            If bInterPt.DistPt(bEnd) < bSt.DistPt(bEnd) Then
                sX = bInterPt.x
                sY = bInterPt.y
                sZ = bInterPt.z
            Else
                sX = bSt.x
                sY = bSt.y
                sZ = bSt.z
            End If
        
            'Project end point of top line on bottom. If it is within exisiting bottom line then set this as end point
            Set bInterPt = tEnd.Offset(vOrient)
            If bInterPt.DistPt(bSt) < bEnd.DistPt(bSt) Then
                eX = bInterPt.x
                eY = bInterPt.y
                eZ = bInterPt.z
            Else
                eX = bEnd.x
                eY = bEnd.y
                eZ = bEnd.z
            End If
        
            Set pBotLine = Nothing
            Set pBotLine = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, sX, sY, sZ, eX, eY, eZ)
            
            If newHRsegments Is Nothing Then
                pComplex.AddCurve pBotLine, False
                pComplex.GetCurves newHRsegments
            Else
                newHRsegments.Add pBotLine
            End If
            
            Set pBotLine = Nothing
            Set pTopLine = Nothing
            Set tSt = Nothing
            Set tEnd = Nothing
            Set bSt = Nothing
            Set bEnd = Nothing
            Set bInterPt = Nothing
            Set pComplex = Nothing
            Set vZ = Nothing
            Set vTan = Nothing
            Set vNorm = Nothing
            Set vOrient = Nothing
        End If  'end if it is pLine
    Next SegIndex  'next line or arc in handrail path
 
    Set pHRsegments = Nothing
    Set pHRsegments = newHRsegments
    
    Exit Sub
ErrorHandler:
    Dim oErrors As New IMSErrorLog.JServerErrors
    oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
End Sub

'TR#51383- checks whether complex string contains any arc segment
Public Function ComplexstringcontainsArc(pComplex As ComplexString3d) As Boolean
Const METHOD = "ComplexstringcontainsArc"
On Error GoTo ErrorHandler

    Dim pSegments As IJElements
    Dim i As Integer

    ComplexstringcontainsArc = False
    
    pComplex.GetCurves pSegments
    If Not pSegments Is Nothing Then
        For i = 1 To pSegments.Count
            If Not (TypeOf pSegments(i) Is IJLine) Then
                ComplexstringcontainsArc = True
                Exit For
            End If
        Next i
    End If

    Set pSegments = Nothing
Exit Function

ErrorHandler:
    Dim oErrors As New IMSErrorLog.JServerErrors
    oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise Err.Number
End Function

'#TR-CP·52075 - mkaveti
'This public method currently validates for the user keyed in inputs of Height, ToTopOfMidRailDistance, NoOfMidRails and MidRailSpacing.
Public Function ValidateHandrailAttributes(ByVal pIJDAttrs As SP3DStructInterfaces.IJDAttributes, sInterfaceName As String, strAttribValid As Boolean) As String
    
    Const METHOD = "ValidateHandrailAttributes"
    On Error GoTo ErrorHandler
    
    Dim dHRHeight As IJDAttribute
    Dim dHRMRdistance As IJDAttribute
    Dim dNoOfMidrails As IJDAttribute
    Dim dMidRailSpacing As IJDAttribute
    
    strAttribValid = True
    
    'Get the Handrail Attributes to validate -mkaveti
    Set dHRHeight = pIJDAttrs.CollectionOfAttributes(sInterfaceName).Item("Height")
    Set dHRMRdistance = pIJDAttrs.CollectionOfAttributes(sInterfaceName).Item("TopOfMidRailDim")
    Set dNoOfMidrails = pIJDAttrs.CollectionOfAttributes(sInterfaceName).Item("NoOfMidRails")
    Set dMidRailSpacing = pIJDAttrs.CollectionOfAttributes(sInterfaceName).Item("MidRailSpacing")
    
    'Check whether Top of Midrail distance is more than Handrai Height -mkaveti
    If dHRHeight.Value <= dHRMRdistance.Value Then
        strAttribValid = False
        Exit Function
    End If
    
    'Check whether Number of MidRails are keyed in with a valid MidRail spacing or vice versa -mkaveti
    If ((dNoOfMidrails.Value - 1) * (dMidRailSpacing.Value)) >= (dHRMRdistance.Value) Then
        strAttribValid = False
        Exit Function
    End If

Exit Function

ErrorHandler:
    Dim oErrors As New IMSErrorLog.JServerErrors
    oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
End Function
'TR#55627- Function to get COG based on curve for toprail, moidrail, toe plate.
' It will first get the curve for top/midrail/toe plate based on hr path curve , height & orientation
' We are geting start point & end point of toprail & midrail. This will be useful in COG calculations of begin/end treatment posts
Public Function GetCOGForTopMidRail(HRpathCurve As ComplexString3d, Height As Double, _
                    Orientation As Integer, Optional ByRef sPt As DPosition, Optional ByRef ePt As DPosition) As DPosition
Const METHOD = "GetCOGForTopMidRail"
On Error GoTo ErrorHandler
    
    Dim pComplex As ComplexString3d
    Dim oVector As DVector
    Dim oTrans4x4  As IJDT4x4
    Dim oSegments  As IJElements
    Dim oCurve As IJCurve
    Dim totalLen As Double
    Dim tmpPar As Double
    Dim i As Integer
    Dim x As Double, y As Double, z As Double
    Dim x1 As Double, y1 As Double, z1 As Double
    
    If m_GeomFactory Is Nothing Then
        Set m_GeomFactory = New GeometryFactory
    End If

    Set oVector = New DVector
    Set pComplex = New ComplexString3d
    Set GetCOGForTopMidRail = New DPosition
    
    If Orientation And ComplexstringcontainsArc(HRpathCurve) = False Then
        Set pComplex = GetPathByOffset(HRpathCurve, Height, 0)
    Else
        HRpathCurve.GetCurves oSegments
        Set pComplex = m_GeomFactory.ComplexStrings3d.CreateByCurves(Nothing, oSegments)
        
        oVector.Set 0, 0, Height
        Set oTrans4x4 = New DT4x4
        oTrans4x4.LoadIdentity
        oTrans4x4.Translate oVector
        pComplex.Transform oTrans4x4
    End If

    pComplex.GetCurves oSegments
    
    totalLen = 0#
    If Not oSegments Is Nothing Then
        For i = 1 To oSegments.Count
            If TypeOf oSegments(i) Is IJCurve Then
                Set oCurve = oSegments(i)
            End If
            If Not oCurve Is Nothing Then
                oCurve.Centroid x, y, z
                    totalLen = totalLen + oCurve.Length
                    x = GetCOGForTopMidRail.x + x * oCurve.Length
                    y = GetCOGForTopMidRail.y + y * oCurve.Length
                    z = GetCOGForTopMidRail.z + z * oCurve.Length
                GetCOGForTopMidRail.Set x, y, z
            End If
        Next i
        
        x = GetCOGForTopMidRail.x / totalLen
        y = GetCOGForTopMidRail.y / totalLen
        z = GetCOGForTopMidRail.z / totalLen
        GetCOGForTopMidRail.Set x, y, z
    End If
    
    ' Get start & end points which are by calling function for COG calculations of begin/end treatment
    If Not sPt Is Nothing And Not ePt Is Nothing Then
        Set oCurve = Nothing
        If TypeOf oSegments(1) Is IJCurve Then
            Set oCurve = oSegments(1)
        End If
        If Not oCurve Is Nothing Then
            'tmpPar = 0#
            oCurve.EndPoints x, y, z, x1, y1, z1
            sPt.Set x, y, z
        End If
        If TypeOf oSegments(oSegments.Count) Is IJCurve Then
            Set oCurve = oSegments(oSegments.Count)
        End If
        If Not oCurve Is Nothing Then
            'tmpPar = oCurve.Length
            oCurve.EndPoints x, y, z, x1, y1, z1
            ePt.Set x1, y1, z1
        End If
    End If
   
    Set pComplex = Nothing
    Set oVector = Nothing
    Set oSegments = Nothing
    Set oTrans4x4 = Nothing
    Set oCurve = Nothing

Exit Function
ErrorHandler:
    Dim oErrors As New IMSErrorLog.JServerErrors
    oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
End Function


'for given path and top/mid rail cross section try to use CreateBySingleSweep() with brkcrv argument =3
' If it is successful then it is valid. CreateBySingleSweep() fails hadrail path contains arcs whose radius is
' too large for given cross section. In such case brkcrv argument value to be changed to 2 or 1 or 1 which works for
' given handrail path

Public Sub ValidateInputForSkinning(pComplex As ComplexString3d, ByRef SkinOption As Long, maxDia As Double)
    Const METHOD = "ValidateInputPath"
    On Error GoTo ErrorHandler
    
    Dim xS As Double, yS As Double, zS As Double
    Dim tX As Double, tY As Double, tZ As Double
    Dim t2x As Double, t2y As Double, t2z As Double
    Dim pSegments As IJElements
    Dim pCurve      As IJCurve
    Dim SParam As Double, EParam As Double
    Dim tmp1() As Double, tmp2() As Double
    Dim pProjections    As IJElements
    Dim bcapped As Long
    Dim brkcrv As Long
        
    If m_GeomFactory Is Nothing Then
        Set m_GeomFactory = New GeometryFactory
    End If

    pComplex.GetCurves pSegments
    If Not pSegments Is Nothing Then
        If TypeOf pSegments(1) Is IJCurve Then
            Set pCurve = pSegments(1)
        End If
    End If
    
    If Not pCurve Is Nothing Then
        pCurve.ParamRange SParam, EParam
        pCurve.Evaluate SParam, xS, yS, zS, tX, tY, tZ, t2x, t2y, t2z
    
        Dim pCircle As IJCircle
        Set pCircle = New Circle3d
           
        pCircle.DefineByCenterNormalRadius xS, yS, zS, tX, tY, tZ, maxDia / 2
        
        'get appropriate bcapped & brkcrv argument values based on skinoption. see CreateProjectionFromCSProfile()
        ' and CreateBySingleSweep documentation for details
        Call getSkinningoptions(SkinOption, bcapped, brkcrv)
        brkcrv = 1
        Set pProjections = m_GeomFactory.GeometryServices.CreateBySingleSweep(Nothing, _
                                                    pComplex, pCircle, 0, brkcrv, _
                                                    tmp1, tmp2, bcapped)
        If pProjections.Count < 1 Then
           'Earlier skin option failed so change the brkcrv option
            If pProjections.Count < 1 And (SkinOption = 6 Or SkinOption = 7) Then
               SkinOption = SkinOption - 2
               Call getSkinningoptions(SkinOption, bcapped, brkcrv)
               Set pProjections = m_GeomFactory.GeometryServices.CreateBySingleSweep(Nothing, _
                                                    pComplex, pCircle, 0, brkcrv, _
                                                    tmp1, tmp2, bcapped)
            End If
            
            If pProjections.Count < 1 And (SkinOption = 4 Or SkinOption = 5) Then
               SkinOption = SkinOption - 2
               Call getSkinningoptions(SkinOption, bcapped, brkcrv)
               Set pProjections = m_GeomFactory.GeometryServices.CreateBySingleSweep(Nothing, _
                                                    pComplex, pCircle, 0, brkcrv, _
                                                    tmp1, tmp2, bcapped)
            End If
            
            If pProjections.Count < 1 And (SkinOption = 2 Or SkinOption = 3) Then
               SkinOption = SkinOption - 2
               Call getSkinningoptions(SkinOption, bcapped, brkcrv)
               Set pProjections = m_GeomFactory.GeometryServices.CreateBySingleSweep(Nothing, _
                                                    pComplex, pCircle, 0, brkcrv, _
                                                    tmp1, tmp2, bcapped)
            End If
            
        End If
    End If
        
cleanup:
    Set pSegments = Nothing
    Set pCurve = Nothing
    Set pCircle = Nothing
    Set pProjections = Nothing
        
    Exit Sub

ErrorHandler:
    Dim oErrors As New IMSErrorLog.JServerErrors
    oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
End Sub

'For details refer documentation of CreateSurface documentation in StructGenerictools service

Public Sub getSkinningoptions(SkinOption As Long, ByRef bcapped As Long, ByRef brkcrv As Long)

    Select Case SkinOption
        Case 1, 3, 5, 7:
            bcapped = 1
        Case Else:
            bcapped = 0
    End Select
    Select Case SkinOption
        Case 2, 3:
            brkcrv = 1
        Case 4, 5:
            brkcrv = 2
        Case 6, 7:
            brkcrv = 3
        Case Else:
            brkcrv = 0
    End Select
End Sub


Public Function GetCrossSectionDiamension(ToprailSec As String, toprailStd As String, MidrailSec As String, midrailStd As String, _
                        toePltSec As String, toePltStd As String) As Double
    Const METHOD = "GetCrossSectionDiamension"
    On Error GoTo ErrorHandler
    Dim width As Double
    Dim depth As Double
 
    GetCrossSectionDiamension = 0#
     
    GetCrossSecData ToprailSec, toprailStd, width, depth
    If GetCrossSectionDiamension < width Then GetCrossSectionDiamension = width
    If GetCrossSectionDiamension < depth Then GetCrossSectionDiamension = depth
    
    GetCrossSecData MidrailSec, midrailStd, width, depth
    If GetCrossSectionDiamension < width Then GetCrossSectionDiamension = width
    If GetCrossSectionDiamension < depth Then GetCrossSectionDiamension = depth
    
    GetCrossSecData toePltSec, toePltStd, width, depth
    If GetCrossSectionDiamension < width Then GetCrossSectionDiamension = width
    If GetCrossSectionDiamension < depth Then GetCrossSectionDiamension = depth

    Exit Function
    
ErrorHandler:
    Dim oErrors As New IMSErrorLog.JServerErrors
    oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
End Function

'Add the path as one of the outputs, but make it only locatable and not displayable.
'This is to enable smartsketch to locate points off the path
Public Sub AddPathAsOutputOfHandrail(pOC As IJDOutputCollection, _
                                     ByVal pComplex As ComplexString3d)
Const METHOD = "AddPathAsOutputOfHandrail"
On Error GoTo ErrorHandler

    Dim i               As Integer
    Dim j               As Integer
    Dim dSx             As Double
    Dim dSy             As Double
    Dim dSz             As Double
    Dim dEx             As Double
    Dim dEy             As Double
    Dim dEz             As Double
    Dim OutStr          As String
    
    Dim oCurve          As IJCurve
    Dim oLine           As IJLine
    Dim oArc            As IJArc
    Dim oEllipArc       As EllipticalArc3d
    Dim oCurveElems     As IJElements
    Dim oCtlFlags       As IJControlFlags
    Dim oTempPoint      As IngrGeom3D.Point3d
    Dim oOutputComplexString As IJComplexString
    
    i = 0
    If m_GeomFactory Is Nothing Then
        Set m_GeomFactory = New GeometryFactory
    End If
    
    pComplex.GetCurves oCurveElems
    
    Dim oline1 As Line3d
    Dim oarc1 As Arc3d
    Dim oEllipArc1 As EllipticalArc3d
    
    For j = 1 To oCurveElems.Count
    
        Set oCurve = oCurveElems.Item(j)
        oCurve.EndPoints dSx, dSy, dSz, dEx, dEy, dEz
        
        'Add the curve to the output
        If TypeOf oCurve Is IJLine Then
                
            Set oLine = oCurve
            Dim sX As Double, sY As Double, sZ As Double
            Dim eX As Double, eY As Double, eZ As Double
            oLine.GetStartPoint sX, sY, sZ
            oLine.GetEndPoint eX, eY, eZ
            
            Set oline1 = New Line3d
            Set oline1 = m_GeomFactory.Lines3d.CreateBy2Points(pOC.ResourceManager, sX, sY, sZ, eX, eY, eZ)

            Set oCtlFlags = oline1
            oCtlFlags.ControlFlags(CTL_FLAG_NO_DRAW) = CTL_FLAG_NO_DRAW
            OutStr = "Handrail Path Line"
            InitNewOutput pOC, OutStr & j
            pOC.AddOutput (OutStr & Trim$(Str$(j))), oline1
            Set oLine = Nothing
            Set oline1 = Nothing
            
        ElseIf TypeOf oCurve Is EllipticalArc3d Then
            
            Set oEllipArc = oCurve
            Dim elipCx As Double, elipCy As Double, elipCz As Double
            Dim elipNorX As Double, elipNorY As Double, elipNorZ As Double
            Dim elipMjrX As Double, elipMjrY As Double, elipMjrZ As Double
            Dim elipMgrRadius As Double
            Dim elimMMRatio As Double, elipStaAngle As Double, elipSwpAngle As Double
            
            oEllipArc.GetCenterPoint elipCx, elipCy, elipCz
            oEllipArc.GetNormal elipNorX, elipNorY, elipNorZ
            oEllipArc.GetMajorAxis elipMjrX, elipMjrY, elipMjrZ
            elipMgrRadius = oEllipArc.MajorRadius
            elimMMRatio = oEllipArc.MinorMajorRatio
            elipStaAngle = oEllipArc.StartAngle
            elipSwpAngle = oEllipArc.SweepAngle
            
            Set oEllipArc1 = New EllipticalArc3d
            Set oEllipArc1 = m_GeomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(pOC.ResourceManager, _
                                    elipCx, elipCy, elipCz, elipNorX, elipNorY, elipNorZ, _
                                    elipMgrRadius * elipMjrX, elipMgrRadius * elipMjrY, elipMgrRadius * elipMjrZ, _
                                    elimMMRatio, elipStaAngle, elipSwpAngle)
            
            Set oCtlFlags = oEllipArc1
            oCtlFlags.ControlFlags(CTL_FLAG_NO_DRAW) = CTL_FLAG_NO_DRAW
            OutStr = "Handrail Path Elliptical Arc"
            InitNewOutput pOC, OutStr & j
            pOC.AddOutput (OutStr & Trim$(Str$(j))), oEllipArc1
            Set oEllipArc = Nothing
            Set oEllipArc1 = Nothing
            
        
        ElseIf TypeOf oCurve Is IJArc Then
                
            Set oArc = oCurve
            Dim arcSx As Double, arcSy As Double, arcSz As Double
            Dim arcEx As Double, arcEy As Double, arcEz As Double
            Dim arcCx As Double, arcCy As Double, arcCz As Double
            Dim arcNx As Double, arcNy As Double, arcNz As Double
            
            oArc.GetStartPoint arcSx, arcSy, arcSz
            oArc.GetEndPoint arcEx, arcEy, arcEz
            oArc.GetCenterPoint arcCx, arcCy, arcCz
            oArc.GetNormal arcNx, arcNy, arcNz
            
            Set oarc1 = New Arc3d
            Set oarc1 = m_GeomFactory.Arcs3d.CreateByCtrNormStartEnd(pOC.ResourceManager, arcCx, arcCy, arcCz, _
                                                    arcNx, arcNy, arcNz, arcSx, arcSy, arcSz, arcEx, arcEy, arcEz)
                                                        
            Set oCtlFlags = oarc1
            oCtlFlags.ControlFlags(CTL_FLAG_NO_DRAW) = CTL_FLAG_NO_DRAW
            OutStr = "Handrail Path Arc"
            InitNewOutput pOC, OutStr & j
            pOC.AddOutput (OutStr & Trim$(Str$(j))), oarc1
            Set oArc = Nothing
            Set oarc1 = Nothing
        
        End If
        
        OutStr = "Handrail Path Point "
            
        If j = 1 Then
                                       
            'Add the first point as output
            Set oTempPoint = m_GeomFactory.Points3d.CreateByPoint(pOC.ResourceManager, dSx, dSy, dSz)
            Set oCtlFlags = oTempPoint
            oCtlFlags.ControlFlags(CTL_FLAG_NO_DRAW) = CTL_FLAG_NO_DRAW
            i = i + 1
            InitNewOutput pOC, OutStr & i
            pOC.AddOutput (OutStr & Trim$(Str$(i))), oTempPoint
            Set oCtlFlags = Nothing
            Set oTempPoint = Nothing
                                    
            'Add the second point as output
            Set oTempPoint = m_GeomFactory.Points3d.CreateByPoint(pOC.ResourceManager, dEx, dEy, dEz)
            Set oCtlFlags = oTempPoint
            oCtlFlags.ControlFlags(CTL_FLAG_NO_DRAW) = CTL_FLAG_NO_DRAW
            i = i + 1
            InitNewOutput pOC, OutStr & i
            pOC.AddOutput (OutStr & Trim$(Str$(i))), oTempPoint
            Set oCtlFlags = Nothing
            Set oTempPoint = Nothing

        Else
            'the first point should have already got added by the previous curve, only add the second point
            Set oTempPoint = m_GeomFactory.Points3d.CreateByPoint(pOC.ResourceManager, dEx, dEy, dEz)
            Set oCtlFlags = oTempPoint
            oCtlFlags.ControlFlags(CTL_FLAG_NO_DRAW) = CTL_FLAG_NO_DRAW
            i = i + 1
            InitNewOutput pOC, OutStr & i
            pOC.AddOutput (OutStr & Trim$(Str$(i))), oTempPoint
            Set oCtlFlags = Nothing
            Set oTempPoint = Nothing
        End If
        
        Set oCurve = Nothing
    Next j
    
    Set oCurveElems = Nothing
    
    Exit Sub
    
ErrorHandler:
    Dim oErrors As New IMSErrorLog.JServerErrors
    oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
    
End Sub

Public Sub GetTopRailInfo(ByRef oPart As IJDPart, ByRef TopRailSection As String, ByRef TopRailStandard As String, _
                           ByRef TopRailCP As Integer, ByRef TopRailAngle As Double)

Const METHOD = "GetTopRailInfo"
On Error GoTo ErrorHandler

    Dim defAttrs            As IJDAttributes
    Dim Attrcol             As IJDAttributesCol
    
    Set defAttrs = oPart
    
    If Not defAttrs Is Nothing Then
        Set Attrcol = defAttrs.CollectionOfAttributes("IJUAHRTypeAProps")
        If Not Attrcol Is Nothing Then
            TopRailSection = Attrcol.Item("TopRail_SPSSectionName").Value
            TopRailStandard = Attrcol.Item("TopRail_SPSSectionRefStandard").Value
            TopRailCP = Attrcol.Item("TopRailSectionCP").Value
            TopRailAngle = Attrcol.Item("TopRailSectionAngle").Value
        End If
    End If
    
    'release objects
    Set defAttrs = Nothing
    Set Attrcol = Nothing
    
    Exit Sub
    
ErrorHandler:
    ''release objects
    Set defAttrs = Nothing
    Set Attrcol = Nothing
    
    Dim oErrors As New IMSErrorLog.JServerErrors
    oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
End Sub

Public Function GetTopRailRadius(ByVal TopRailWidth As Double, ByVal TopRailDepth) As Double

Const METHOD = "GetTopRailRadius"
On Error GoTo ErrorHandler

    GetTopRailRadius = (Sqr((TopRailWidth * TopRailWidth) + (TopRailDepth * TopRailDepth))) / 2 + 3 * 0.0254
    
    Exit Function
    
ErrorHandler:
    
    Dim oErrors As New IMSErrorLog.JServerErrors
    oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
End Function

Public Sub CalcRailsVolumeCG(oHandrail As ISPSHandrail, _
                      ByVal PartOccInfoCol As IJDInfosCol, _
                      HandrailType As eHandrailType, _
                      ByRef volume As Double, _
                      ByRef VCogX As Double, ByRef VCogY As Double, ByRef VCogZ As Double)
    Const METHOD = "CalcRailsVolumeCG"
    On Error GoTo ErrHandler
    
    Dim i As Integer
    
    Dim OccAttrs As IJDAttributes
    
    Dim Length  As Double
    Dim toeplateht As Double
    Dim railwt As Double, totalVolume As Double
    Dim ToprailSec As String, TopRailStandard As String
    Dim MidrailSec As String, MidRailStandard As String
    Dim ToePlateSec As String, ToePlateStandard As String
    Dim PostSec As String, PostStandard As String
    Dim noofmidrails As Integer, midrailno As Integer
    Dim material As Variant, grade As Variant
    Dim TopOfMidRailDim As Double
    Dim MidRailSpacing As Double
    Dim SPSHRBeginTreatment As Integer
    Dim SPSHREndTreatment As Integer
    Dim treatmentPostLength As Double
    Dim Height As Double
    Dim outPutString As String
    Dim PostCSArea As Double
    Dim CSArea As Double
    Dim totalCOG As DPosition
    Dim partCOG As DPosition
    Dim HRpathCurve As ComplexString3d
    Dim HRpath As Sketch3d
    Dim Orientation As Integer
    Dim WithToePlate As Boolean
    
    Set partCOG = New DPosition
    Set totalCOG = New DPosition
    
    Set OccAttrs = oHandrail
     
    Set HRpath = oHandrail.SketchPath
    Dim strOffsetTypeAttrName As String
    Dim strOffsetValAttrName As String
    strOffsetTypeAttrName = "SPSHRPathHorizontalOffsetType"
    strOffsetValAttrName = "SPSHRPathHorizontalOffset"
   
    If HandrailType = TypeA Then
        strOffsetTypeAttrName = "HorizontalOffset"
        strOffsetValAttrName = "HorizontalOffsetDim"
    End If
    
   
    Dim iOffsetType As Integer
    iOffsetType = GetAttribute(OccAttrs, strOffsetTypeAttrName, PartOccInfoCol)
    Dim dOffsetValue As Double
    dOffsetValue = GetAttribute(OccAttrs, strOffsetValAttrName, PartOccInfoCol)
    Set HRpathCurve = GetOffsetCurve(HRpath.GetComplexString, _
                               iOffsetType, _
                               dOffsetValue)
    Set HRpath = Nothing
    Orientation = GetAttribute(OccAttrs, "HandrailOrientation", PartOccInfoCol) 'TR#55627
    
    'The length attribute is not currently being correctly set by the handrail semantic
    'It simply uses the length of the sketch path which may not be the same as the actual length
    'Planning has not decided yet if this should be fixed for the current handrail symbols
    'or if it is simply a requirement of the pending "Traffic Items as Parts" CR
'    If HandrailType = TypeA Then
'        Length = GetAttribute(OccAttrs, "TotalLength", PartOccInfoCol)
'    Else
'        Length = GetAttribute(OccAttrs, "Length", PartOccInfoCol)
'    End If
    Dim oOffsetpathCurve As IJCurve
    Set oOffsetpathCurve = HRpathCurve
        Length = oOffsetpathCurve.Length
    
    toeplateht = GetAttribute(OccAttrs, "TopOfToePlateDim", PartOccInfoCol)
    
    ToprailSec = GetAttribute(OccAttrs, "TopRail_SPSSectionName", PartOccInfoCol)
    
    TopRailStandard = GetAttribute(OccAttrs, "TopRail_SPSSectionRefStandard", PartOccInfoCol)
    
    MidrailSec = GetAttribute(OccAttrs, "MidRail_SPSSectionName", PartOccInfoCol)
    
    MidRailStandard = GetAttribute(OccAttrs, "MidRail_SPSSectionRefStandard", PartOccInfoCol)
    
    noofmidrails = GetAttribute(OccAttrs, "NoOfMidRails", PartOccInfoCol)
    
    ToePlateSec = GetAttribute(OccAttrs, "ToePlate_SPSSectionName", PartOccInfoCol)
    
    ToePlateStandard = GetAttribute(OccAttrs, "ToePlate_SPSSectionRefStandard", PartOccInfoCol)
   
    TopOfMidRailDim = GetAttribute(OccAttrs, "TopOfMidRailDim", PartOccInfoCol)
        
    MidRailSpacing = GetAttribute(OccAttrs, "MidRailSpacing", PartOccInfoCol)
        
    Height = GetAttribute(OccAttrs, "Height", PartOccInfoCol)
    
    WithToePlate = GetAttribute(OccAttrs, "WithToePlate", PartOccInfoCol)
    
    CSArea = 0#
    totalVolume = 0#
    totalCOG.Set 0#, 0#, 0#
        
    'TopRail
    CSArea = GetCSAttribData(ToprailSec, TopRailStandard, "ISTRUCTCrossSectionDimensions", "Area")
        totalVolume = totalVolume + CSArea * Length
    Set partCOG = GetCOGForTopMidRail(HRpathCurve, Height, Orientation)
        totalCOG.x = totalCOG.x + partCOG.x * CSArea * Length
        totalCOG.y = totalCOG.y + partCOG.y * CSArea * Length
        totalCOG.z = totalCOG.z + partCOG.z * CSArea * Length
    
    'MidRail
    CSArea = GetCSAttribData(MidrailSec, MidRailStandard, "ISTRUCTCrossSectionDimensions", "Area")
        totalVolume = totalVolume + CSArea * Length * noofmidrails
    For i = 1 To noofmidrails
        Set partCOG = GetCOGForTopMidRail(HRpathCurve, (TopOfMidRailDim - (i - 1) * MidRailSpacing), Orientation)
            totalCOG.x = totalCOG.x + partCOG.x * CSArea * Length
            totalCOG.y = totalCOG.y + partCOG.y * CSArea * Length
            totalCOG.z = totalCOG.z + partCOG.z * CSArea * Length
    Next i
    
    'ToePlate
    If WithToePlate Then
        CSArea = GetCSAttribData(ToePlateSec, ToePlateStandard, "ISTRUCTCrossSectionDimensions", "Area")
            totalVolume = totalVolume + CSArea * Length
        
        Set partCOG = GetCOGForTopMidRail(HRpathCurve, toeplateht, Orientation)
            totalCOG.x = totalCOG.x + partCOG.x * CSArea * Length
            totalCOG.y = totalCOG.y + partCOG.y * CSArea * Length
            totalCOG.z = totalCOG.z + partCOG.z * CSArea * Length
    End If
    
    
    volume = totalVolume
    
    VCogX = totalCOG.x
    VCogY = totalCOG.y
    VCogZ = totalCOG.z
    
    ' TR#55627
    Set partCOG = Nothing
    Set totalCOG = Nothing
    Set HRpathCurve = Nothing

Exit Sub
ErrHandler:
    Dim oErrors As New IMSErrorLog.JServerErrors
    oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
End Sub

Public Sub AddPostVolumeInfo(PostCSArea As Double, PostPosition As DPosition, vOV As DVector, dLength As Double, ByRef volume, ByRef VCogX As Double, ByRef VCogY As Double, ByRef VCogZ As Double)
    Const METHOD = "AddPostVolumeInfo"
    On Error GoTo ErrHandler
    Dim localVolume As Double
    localVolume = PostCSArea * dLength
    volume = volume + localVolume
    VCogX = VCogX + localVolume * (PostPosition.x + vOV.x * dLength / 2)
    VCogY = VCogY + localVolume * (PostPosition.y + vOV.y * dLength / 2)
    VCogZ = VCogZ + localVolume * (PostPosition.z + vOV.z * dLength / 2)
Exit Sub
ErrHandler:
    Dim oErrors As New IMSErrorLog.JServerErrors
    oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL

End Sub

Public Sub AddTreatmentVolumeInfo(PostCSArea As Double, PostPosition As DPosition, vOV As DVector, dLength As Double, TreatmentType As Integer, ByRef volume, ByRef VCogX As Double, ByRef VCogY As Double, ByRef VCogZ As Double)
    Const METHOD = "AddTreatmentVolumeInfo"
    On Error GoTo ErrHandler
    Dim localVolume As Double
    localVolume = PostCSArea * dLength
    If TreatmentType = 5 Then
        localVolume = localVolume * 0.857 ' based on arc radius = (PostHeight - midrailHt)/6
    End If
    volume = volume + localVolume
    VCogX = VCogX + localVolume * (PostPosition.x + vOV.x * dLength / 2)
    VCogY = VCogY + localVolume * (PostPosition.y + vOV.y * dLength / 2)
    VCogZ = VCogZ + localVolume * (PostPosition.z + vOV.z * dLength / 2)
Exit Sub
ErrHandler:
    Dim oErrors As New IMSErrorLog.JServerErrors
    oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL

End Sub


Public Function GetOffsetCurve(pathCurve As ComplexString3d, _
                               iOffsetType As Integer, _
                               dOffsetValue As Double) As ComplexString3d
                               
    Const METHOD = "GetOffsetCurve"
    On Error GoTo ErrHandler
    Dim pSegments As IJElements

            
    Dim pCurve      As IJCurve

    ' offset the sketch 3d curve by horizontal offset
    'or if offset distnace is very very small some of default values are zero
    If iOffsetType = 15 Or Abs(dOffsetValue) < dtol Then
        Set GetOffsetCurve = pathCurve
    Else
        If m_GeomFactory Is Nothing Then
            Set m_GeomFactory = New GeometryFactory
        End If
        Set pCurve = pathCurve
        If pCurve.Scope = CURVE_SCOPE_PLANAR Then
            Dim hintPoint As IJDPosition
            Dim pointOnRight As Boolean
            pointOnRight = (iOffsetType = 5)
            Set hintPoint = GetHintPoint(pCurve, pointOnRight)
            Set GetOffsetCurve = m_GeomFactory.GeometryServices.CreateByOffset(Nothing, _
                                                      pathCurve, _
                                                      hintPoint.x, hintPoint.y, hintPoint.z, dOffsetValue, 0)
                                                      
        ElseIf ComplexstringcontainsArc(pathCurve) = False Then
            If iOffsetType = 10 Then    'left
                Set GetOffsetCurve = GetPathByOffset(pathCurve, dOffsetValue, 1)
            Else 'right
                Set GetOffsetCurve = GetPathByOffset(pathCurve, -dOffsetValue, 1)
            End If
        End If
        
 
        If GetOffsetCurve Is Nothing Then Set GetOffsetCurve = pathCurve ' default curve for any failure
        
        Set pCurve = Nothing
       
        
    End If
Exit Function
ErrHandler:
    Dim oErrors As New IMSErrorLog.JServerErrors
    oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
End Function

'The angle adjustment for regular Treatments
Public Sub CalcStartEndTreatmentOrientationAdjustments(pSegments As IJElements, ByRef StartAngle As Double, ByRef StartXSign As Integer, ByRef StartYSign As Integer, ByRef EndAngle As Double, ByRef EndXSign As Integer, ByRef EndYSign As Integer)
    Const METHOD = "CalcStartEndTreatmentOrientationAdjustments"
    On Error GoTo ErrHandler
    Dim pLine                   As IJLine
    Dim pCurve                  As IJCurve
    Dim sX As Double, sY As Double, sZ As Double
    Dim eX As Double, eY As Double, eZ As Double
    Dim VecX As Double, VecY As Double, VecZ As Double
    
    Dim SParam                  As Double
    Dim EParam                  As Double
    Dim xa As Double, ya As Double, za As Double
    Dim t2x As Double, t2y As Double, t2z As Double
    
    StartXSign = 1
    StartYSign = 1
    EndXSign = 1
    EndYSign = 1
    
    If TypeOf pSegments(1) Is IJLine Then
        Set pLine = pSegments(1)
        pLine.GetStartPoint sX, sY, sZ
        pLine.GetEndPoint eX, eY, eZ
        VecX = sX - eX
        VecY = sY - eY
        VecZ = sZ - eZ
    Else
        Set pCurve = pSegments(1)
        pCurve.ParamRange SParam, EParam
        pCurve.Evaluate SParam, xa, ya, za, VecX, VecY, VecZ, t2x, t2y, t2z
        VecX = -VecX
        VecY = -VecY
        VecZ = -VecZ
    End If
    If Abs(VecY) > 0 Then
        StartYSign = Sgn(VecY)
    End If
    If Abs(VecX) > 0 Then
        StartXSign = Sgn(VecX)
        StartAngle = Atn(VecY / VecX)
    Else
        StartAngle = StartYSign * PI / 2
    End If
    
    If TypeOf pSegments(pSegments.Count) Is IJLine Then
        Set pLine = pSegments(pSegments.Count)
        pLine.GetStartPoint sX, sY, sZ
        pLine.GetEndPoint eX, eY, eZ
        VecX = sX - eX
        VecY = sY - eY
        VecZ = sZ - eZ
    Else
        Set pCurve = pSegments(pSegments.Count)
        pCurve.ParamRange SParam, EParam
        pCurve.Evaluate EParam, xa, ya, za, VecX, VecY, VecZ, t2x, t2y, t2z
        VecX = -VecX
        VecY = -VecY
        VecZ = -VecZ
    End If
    If Abs(VecY) > 0 Then
        EndYSign = Sgn(VecY)
    End If
    If Abs(VecX) > 0 Then
        EndXSign = Sgn(VecX)
        EndAngle = Atn(VecY / VecX)
    Else
        EndAngle = EndYSign * PI / 2
    End If

Exit Sub
ErrHandler:
    Dim oErrors As New IMSErrorLog.JServerErrors
    oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
End Sub

'The post and treatment positions are calculated and stored in the Posts collection, which can be used to create posts and treatments
Public Sub CalcPostPositions(Posts As Collection, _
                        pSegments As IJElements, _
                        BeginExtensionLength As Double, _
                        BeginTreatType As Integer, _
                        EndExtensionLength As Double, _
                        EndTreatType As Integer, _
                        SegmentMaxSpacing As Double, _
                        SlopedSegmentMaxSpacing As Double, _
                        bPostAtEveryTurn As Boolean, _
                        MinClearenceAtPostTurn As Double, _
                        MaxClearenceAtPostTurn As Double, ToprailSectionAngle As Double, PostSectionAngle As Double, Orientation As Integer, Mirror As Boolean, DeltaAngle As Double)
    Const METHOD = "CalcPostPositions"
    On Error GoTo ErrHandler
    Dim SegIndex As Integer
    Dim SegLength As Double
    Dim post As HandrailPost
    Dim StartPosition As DPosition
    Dim EndPosition As DPosition
    Dim StartPathVec As DVector
    Dim EndPathVec As DVector
    Dim PostPosition As DPosition
    Dim PostDirection As DVector
    Dim PathVec As DVector
    Dim SectionAngle As Double
    
    Dim pLine                   As IJLine
    Dim pCurve                  As IJCurve
    Dim curvescope As Geom3dCurveScopeConstants
    
    Dim sX As Double, sY As Double, sZ As Double
    Dim eX As Double, eY As Double, eZ As Double
    Dim nX As Double, nY As Double, nZ As Double
    Dim px As Double, py As Double, pz As Double
    Dim xa As Double, xb As Double, ya As Double, yb As Double, za As Double, zb As Double
    Dim curveParam     As Double
    Dim curveDist     As Double
    Dim NoOfIntermediatePosts     As Integer
    Dim SParam As Double, EParam As Double
    Dim SegVec As DVector
    Dim bslopecol As Collection
    Dim IsSloped As Boolean
    Dim bPostAtEveryTurn1 As Boolean
    bPostAtEveryTurn1 = bPostAtEveryTurn
    Dim PostDistBeforeTurnPercent As Double, PostDistAfterTurnPercent As Double
    
    Dim dStartAngle As Double
    Dim iStartXSign As Integer
    Dim iStartYSign As Integer
    Dim dEndAngle As Double
    Dim iEndXSign As Integer
    Dim iEndYSign As Integer
    'Post type: 2 for post, 5 for circular treatment, 10 for regular treatment
    'Get angle adjustments for treatments if necessary
    If (BeginExtensionLength > 0 And BeginTreatType <> 2) Or (EndExtensionLength > 0 And EndTreatType <> 2) Then
        CalcStartEndTreatmentOrientationAdjustments pSegments, dStartAngle, iStartXSign, iStartYSign, dEndAngle, iEndXSign, iEndYSign
    End If
    
    Set bslopecol = New Collection
    
    If Posts Is Nothing Then
        Set Posts = New Collection
    Else
        While Posts.Count > 0
            Posts.Remove Posts.Count
        Wend
    End If
    
    Set SegVec = New DVector
    Set StartPosition = New DPosition
    Set EndPosition = New DPosition
    Set PostPosition = New DPosition
    Dim postcount As Integer
    For SegIndex = 1 To pSegments.Count
               
        sX = 0#
        sY = 0#
        sZ = 0#
        eX = 0#
        eY = 0#
        eZ = 0#
        SegLength = 0#
        NoOfIntermediatePosts = 0#
        curveParam = 0#
        curveDist = 0#
        While bslopecol.Count > 0
            bslopecol.Remove bslopecol.Count
        Wend
        bslopecol.Add NeedNoClearance(pSegments, SegIndex, True)
        bslopecol.Add NeedNoClearance(pSegments, SegIndex, False)
        bPostAtEveryTurn = bPostAtEveryTurn1
        
        Set StartPathVec = New DVector
        Set EndPathVec = New DVector
        If TypeOf pSegments(SegIndex) Is IJLine Then
            Set pLine = pSegments(SegIndex)
            pLine.GetStartPoint sX, sY, sZ
            If SegIndex = 1 Then StartPosition.Set sX, sY, sZ
            pLine.GetEndPoint eX, eY, eZ
            If SegIndex = pSegments.Count Then EndPosition.Set eX, eY, eZ
            
            SegVec.Set eX - sX, eY - sY, eZ - sZ
                SegVec.Length = 1
            If SegIndex = 1 Then Set StartPathVec = SegVec
            If SegIndex = pSegments.Count Then Set EndPathVec = SegVec
            Set pCurve = pLine
        Else
            Set pCurve = pSegments(SegIndex)
            pCurve.ParamRange SParam, EParam
                    
            Dim tX As Double, tY As Double, tZ As Double
            Dim t2x As Double, t2y As Double, t2z As Double
        
            pCurve.Evaluate SParam, xa, ya, za, tX, tY, tZ, t2x, t2y, t2z
            
            SegVec.Set tX, tY, tZ
            If SegIndex = 1 Then
                StartPosition.Set xa, ya, za
                Set StartPathVec = SegVec
            End If
            pCurve.Evaluate EParam, xa, ya, za, tX, tY, tZ, t2x, t2y, t2z
            If SegIndex = pSegments.Count Then
                EndPosition.Set xa, ya, za
                EndPathVec.Set tX, tY, tZ
                    EndPathVec.Length = 1
            End If
        End If
        
        Set PostDirection = New DVector
        
        'only one direction for each segment (only lines and planar curves are supported now)
        If Orientation Then
            GetPerpendicularDirection SegVec, tX, tY, tZ 'TR#51383- Orientation= Perpendicular to Slope
            PostDirection.Set tX, tY, tZ
        Else
            PostDirection.Set 0#, 0#, 1# 'TR#51383- Orienation= Always vertical
        End If
        
        
        'add start treatment position if applicable
        If SegIndex = 1 And BeginExtensionLength > 0 And BeginTreatType <> 2 Then
            If BeginTreatType = 10 Then
                AddPost Posts, StartPosition, PostDirection, StartPathVec, BeginTreatType, ToprailSectionAngle + dStartAngle - (iStartXSign - 1) * PI / 2
            Else
                AddPost Posts, StartPosition, PostDirection, StartPathVec, BeginTreatType, ToprailSectionAngle
            End If
        End If
        
        pCurve.EndPoints sX, sY, sZ, eX, eY, eZ
            SegLength = pCurve.Length
           
        pCurve.ParamRange SParam, EParam
        
        IsSloped = SegmentIsSloped(pSegments, SegIndex)
        If IsSloped Then
            bPostAtEveryTurn = True
        End If
        If bslopecol.Item(1) And bslopecol.Item(2) Then
            bPostAtEveryTurn = True
        End If
        
        Dim sP  As Double, eP As Double
        sP = SParam
        eP = EParam
        'leave space for begin treatment extension
        If SegIndex = 1 And BeginExtensionLength > 0 And BeginTreatType <> 2 Then
                sP = BeginExtensionLength * (EParam - SParam) / pCurve.Length + SParam
            SegLength = SegLength - BeginExtensionLength
        End If
        'leave space for end treatment extension
        If SegIndex = pSegments.Count And EndExtensionLength > 0 And EndTreatType <> 2 Then  'TR#52838- only if end tratment is applied
 
            SegLength = SegLength - EndExtensionLength
                eP = (pCurve.Length - EndExtensionLength) * (EParam - SParam) / pCurve.Length + SParam
        End If
          
        Dim dPostClearance As Double
        Dim MaxPostDistClearance As Double
        Dim MinPostDistClearance As Double
        MaxPostDistClearance = MaxClearenceAtPostTurn
        MinPostDistClearance = MinClearenceAtPostTurn
        
        If SegIndex = 1 Then
            CalculateSegmentDetails SegLength, NoOfIntermediatePosts, bPostAtEveryTurn, IsSloped, SlopedSegmentMaxSpacing, SegmentMaxSpacing, MaxPostDistClearance, MinPostDistClearance, dPostClearance, SEGMENT_BEGIN_TYPE, bslopecol
        ElseIf SegIndex = pSegments.Count Then
            CalculateSegmentDetails SegLength, NoOfIntermediatePosts, bPostAtEveryTurn, IsSloped, SlopedSegmentMaxSpacing, SegmentMaxSpacing, MaxPostDistClearance, MinPostDistClearance, dPostClearance, SEGMENT_END_TYPE, bslopecol
        Else
            CalculateSegmentDetails SegLength, NoOfIntermediatePosts, bPostAtEveryTurn, IsSloped, SlopedSegmentMaxSpacing, SegmentMaxSpacing, MaxPostDistClearance, MinPostDistClearance, dPostClearance, SEGMENT_MID_TYPE, bslopecol
        End If
        
        CalcPostDistancePercentAtTurn pSegments.Count, SegIndex, SegLength, MaxPostDistClearance, MinPostDistClearance, dPostClearance, bPostAtEveryTurn, PostDistBeforeTurnPercent, PostDistAfterTurnPercent, NoOfIntermediatePosts, bslopecol
        
        If NoOfIntermediatePosts > 0 Then curveDist = (eP - sP) * (1 - PostDistBeforeTurnPercent - PostDistAfterTurnPercent) / NoOfIntermediatePosts
        
        Dim j As Integer
        
        For j = 0 To NoOfIntermediatePosts

            If j = NoOfIntermediatePosts And SegIndex < pSegments.Count And bPostAtEveryTurn Then
                Exit For
            End If
            curveParam = j * curveDist + sP + (eP - sP) * PostDistAfterTurnPercent 'TR#104699
            SectionAngle = GetSectionAngleWRTCurveParameterObject(pCurve, curveParam, PostSectionAngle)
            pCurve.Evaluate curveParam, px, py, pz, tX, tY, tZ, t2x, t2y, t2z
            Set PostPosition = New DPosition
            PostPosition.Set px, py, pz
            Set PathVec = New DVector
            PathVec.Set tX, tY, tZ
            'reverse the angle for the last post
            If SegIndex = pSegments.Count And j = NoOfIntermediatePosts And EndTreatType = 2 Then
                If Mirror Then
                    SectionAngle = SectionAngle - DeltaAngle
                Else
                    SectionAngle = SectionAngle + DeltaAngle
                End If
            End If
            AddPost Posts, PostPosition, PostDirection, PathVec, 2, SectionAngle
            Set PostPosition = Nothing
            
        Next j
        'add end treatment position if applicable
        If SegIndex = pSegments.Count And EndExtensionLength > 0 And EndTreatType <> 2 Then  'TR#52838- only if end tratment is applied
            'add a post as treatment
            If EndTreatType = 10 Then
                AddPost Posts, EndPosition, PostDirection, EndPathVec, EndTreatType, ToprailSectionAngle - dEndAngle + (iEndXSign + 1) * PI / 2
            Else
                AddPost Posts, EndPosition, PostDirection, EndPathVec, EndTreatType, ToprailSectionAngle
            End If
        End If
        
        Set pLine = Nothing
        Set pCurve = Nothing
    Next SegIndex
Exit Sub
ErrHandler:
    Dim oErrors As New IMSErrorLog.JServerErrors
    oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
End Sub

'segvec should not be vertical, actually this case is ruled out in client
Private Sub GetPerpendicularDirection(SegVec As DVector, ByRef vX As Double, ByRef vY As Double, ByRef vZ As Double)
    Const METHOD = "GetPerpendicularDirection"
    On Error GoTo ErrHandler
    Dim vecAxis As DVector
    Dim vecPDir As DVector
    
    Set vecAxis = New DVector
    Set vecPDir = New DVector
    
    vecAxis.Set 0, 0, 1#
    Set vecPDir = vecAxis.Cross(SegVec)
    vecPDir.Length = 1#
    Set vecPDir = SegVec.Cross(vecPDir)
    vecPDir.Length = 1#
    vecPDir.Get vX, vY, vZ
    Set vecAxis = Nothing
    Set vecPDir = Nothing

Exit Sub
ErrHandler:
    Dim oErrors As New IMSErrorLog.JServerErrors
    oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
End Sub
'Create a post at the provided location
Public Sub CreatePost(oOutputParentObject As Object, post As HandrailPost, pCSProfileObj As Object, CSectionCP As Integer, PostHeight As Double, dPadOffset As Double, Mirror As Boolean, SkinOption As Long, bNoCSSymbol As Boolean, ByRef outCount As Long, DeltaAngle As Double)

    Const METHOD = "CreatePost"
    On Error GoTo ErrHandler

    Dim ptmpLine    As IJLine
    Dim OutStr      As String
    Dim pProjectionEles As IJElements
    Dim px As Double, py As Double, pz As Double
    Dim vX As Double, vY As Double, vZ As Double

    OutStr = "HRPost"
    post.BasePos.Get px, py, pz
    post.DirectionVec.Get vX, vY, vZ
    
    If m_GeomFactory Is Nothing Then
        Set m_GeomFactory = New GeometryFactory
    End If

    Set ptmpLine = m_GeomFactory.Lines3d.CreateByPtVectLength(Nothing, _
                                        px + dPadOffset * vX, py + dPadOffset * vY, pz + dPadOffset * vZ, vX, vY, vZ, _
                                        PostHeight - dPadOffset)

    BuildHandrailOutput oOutputParentObject, ptmpLine, pCSProfileObj, CSectionCP, post.SectionAngle + DeltaAngle, Mirror, SkinOption, OutStr, MemberType_Post, outCount

    Exit Sub

ErrHandler:
    Dim oErrors As New IMSErrorLog.JServerErrors
    oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
End Sub

'Create a treatment at the provided location
Public Sub CreateTreatment(oOutputParentObject As Object, post As HandrailPost, pSegments As IJElements, bNoCSSymbol As Boolean, pResourceManage As Object, PostHeight As Double, midrailHt As Double, _
                            pTopRailCSProfileObj As Object, TopRailSectionCP As Integer, Mirror As Boolean, SkinOption As Long, TopRailSectionDepth As Double, TreatmentRadius As Double)
    Const METHOD = "CreateTreatment"
    On Error GoTo ErrHandler

    Dim ptmpLine    As IJLine
    Dim pCurve As IJCurve
    Dim OutStr As String
    Dim OutType As Long
    Dim px As Double, py As Double, pz As Double
    Dim vX As Double, vY As Double, vZ As Double
    Dim Index As Integer
    Dim outCount As Long
    
    Dim pOC As IJDOutputCollection

    post.BasePos.Get px, py, pz
    post.DirectionVec.Get vX, vY, vZ
    If m_GeomFactory Is Nothing Then
        Set m_GeomFactory = New GeometryFactory
    End If

    If post.Index > 1 Then
        Set pCurve = pSegments.Item(pSegments.Count)
        OutStr = "HREndTreatment"
        OutType = MemberType_EndTreatment
    Else
        Set pCurve = pSegments.Item(1)
        OutStr = "HRBeginTreatment"
        OutType = MemberType_EndTreatment
    End If

    If post.PostType = 10 Then
        If post.Index > 1 Then
            Set ptmpLine = m_GeomFactory.Lines3d.CreateByPtVectLength(Nothing, _
                                        px + PostHeight * vX, py + PostHeight * vY, pz + PostHeight * vZ, -vX, -vY, -vZ, PostHeight - midrailHt)
        Else
            Set ptmpLine = m_GeomFactory.Lines3d.CreateByPtVectLength(Nothing, _
                                        px + midrailHt * vX, py + midrailHt * vY, pz + midrailHt * vZ, vX, vY, vZ, PostHeight - midrailHt)
        End If
        BuildHandrailOutput oOutputParentObject, ptmpLine, pTopRailCSProfileObj, TopRailSectionCP, post.SectionAngle, Mirror, SkinOption, OutStr, OutType, outCount
    
    ElseIf post.PostType = 5 Then
        If post.Index > 1 Then Mirror = Not Mirror
        Set m_complex = CreateCirEndTreatment(pCurve, Nothing, px, py, pz, midrailHt, PostHeight, post.DirectionVec, post.Index = 1, TopRailSectionDepth, TreatmentRadius) 'TR#51383-changed input vector from segVec to vecOV
        
        BuildHandrailOutput oOutputParentObject, m_complex, pTopRailCSProfileObj, TopRailSectionCP, post.SectionAngle, Not Mirror, SkinOption, OutStr, OutType, outCount
        
    End If
    Exit Sub

ErrHandler:
    Dim oErrors As New IMSErrorLog.JServerErrors
    oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
End Sub

'Check if the segment needs clearance at the specified end
'If the previous segment is sloped then there is no clearance at the beginning side
'If the next segment is sloped then there is no clearance at the end side
Public Function NeedNoClearance(pSegments As IJElements, SegIndex As Integer, IsStart As Boolean) As Boolean
      
    Const METHOD = "NeedNoClearance"
    On Error GoTo ErrHandler
    If SegIndex = 1 And IsStart Then
        NeedNoClearance = True
        Exit Function
    End If
    If SegIndex = pSegments.Count And Not IsStart Then
        NeedNoClearance = True
        Exit Function
    End If
    
    Dim oSlopedCurve As IJCurve
    Dim o1x As Double, o1y As Double, o1z As Double
    Dim o2x As Double, o2y As Double, o2z As Double
    Dim bNoBeginPostClearance As Boolean
    Dim bNoEndPostClearance As Boolean
       
    NeedNoClearance = False
    If IsStart Then
        If TypeOf pSegments.Item(SegIndex - 1) Is IJCurve Then
            Set oSlopedCurve = pSegments.Item(SegIndex - 1)
        End If
        If Not oSlopedCurve Is Nothing Then
            oSlopedCurve.EndPoints o1x, o1y, o1z, o2x, o2y, o2z
            If Abs(o1z - o2z) > dtol Then 'TR#104699
                NeedNoClearance = True
            End If
        End If
    Else
        If TypeOf pSegments.Item(SegIndex + 1) Is IJCurve Then
            Set oSlopedCurve = pSegments.Item(SegIndex + 1)
        End If
        If Not oSlopedCurve Is Nothing Then
            oSlopedCurve.EndPoints o1x, o1y, o1z, o2x, o2y, o2z
            If Abs(o1z - o2z) > dtol Then 'TR#104699
                NeedNoClearance = True
            End If
        End If
    End If
    
Exit Function
ErrHandler:
    Dim oErrors As New IMSErrorLog.JServerErrors
    oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
End Function

'Calculate the clearance percentage on both ends of the segment, also handle the post number round off
Public Sub CalcPostDistancePercentAtTurn(SegmentCount As Integer, SegmentIndex As Integer, SegLength As Double, MaxPostDistClearance As Double, MinPostDistClearance As Double, dFinalClearance As Double, bPostAtEveryTurn As Boolean, ByRef PostDistBeforeTurnPercent As Double, ByRef PostDistAfterTurnPercent As Double, ByRef NoOfIntermediatePosts As Integer, bslopecol As Collection)
    
    Const METHOD = "CalcPostDistancePercentAtTurn"
    On Error GoTo ErrHandler
    If bPostAtEveryTurn Then
        PostDistBeforeTurnPercent = 0
        PostDistAfterTurnPercent = 0
        If NoOfIntermediatePosts = 0 Then
            NoOfIntermediatePosts = 1
        End If
        Exit Sub
    End If
    
    Dim bNoBeginPostClearance As Boolean
    Dim bNoEndPostClearance As Boolean
    
    If Not bslopecol Is Nothing Then
        bNoBeginPostClearance = bslopecol.Item(1)
        bNoEndPostClearance = bslopecol.Item(2)
    End If
    
    PostDistAfterTurnPercent = dFinalClearance / (2 * SegLength)
    PostDistBeforeTurnPercent = dFinalClearance / (2 * SegLength)
    
    If bNoEndPostClearance Then
        PostDistBeforeTurnPercent = 0
    End If
    
    If bNoBeginPostClearance Then
        PostDistAfterTurnPercent = 0
    End If
    If bNoEndPostClearance Or bNoBeginPostClearance Then
        PostDistAfterTurnPercent = 2 * PostDistAfterTurnPercent
        PostDistBeforeTurnPercent = 2 * PostDistBeforeTurnPercent
    End If
    If NoOfIntermediatePosts = 0 And bPostAtEveryTurn = False Then 'add extra post only if bPostAtEveryTurn= false
        Dim TotalClearance As Integer
        TotalClearance = 0
        If Not bslopecol.Item(1) Then
            TotalClearance = TotalClearance + 1
        End If
        If Not bslopecol.Item(2) Then
            TotalClearance = TotalClearance + 1
        End If
        If (SegLength - TotalClearance * MinPostDistClearance) < dtol Then
            NoOfIntermediatePosts = -1 ' no posts on segment if seglength is smaller than 2*min clearacne
        Else
            If (SegLength - TotalClearance * MaxPostDistClearance) < dtol Then
              If TotalClearance = 2 Then
                PostDistBeforeTurnPercent = 0.5
              ElseIf TotalClearance = 1 Then
                If Not bslopecol.Item(1) Then
                    PostDistBeforeTurnPercent = 0
                Else
                    PostDistBeforeTurnPercent = 1
                End If
              End If
              PostDistAfterTurnPercent = 1 - PostDistBeforeTurnPercent
            Else
              NoOfIntermediatePosts = 1
            End If
        End If
        
    End If

Exit Sub
ErrHandler:
    Dim oErrors As New IMSErrorLog.JServerErrors
    oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
End Sub

'add a post into the collection
Public Sub AddPost(Posts As Collection, PostPosition As DPosition, PostDirection As DVector, PathVec As DVector, PostType As Integer, CSectionAngle As Double)
    Const METHOD = "AddPost"
    On Error GoTo ErrHandler
    Dim post As HandrailPost
    Set post = New HandrailPost
    post.SectionAngle = CSectionAngle
    post.PostType = PostType
    Set post.BasePos = PostPosition
    Set post.DirectionVec = PostDirection
    Set post.PathVec = PathVec
    post.Index = Posts.Count + 1
    Posts.Add post

Exit Sub
ErrHandler:
    Dim oErrors As New IMSErrorLog.JServerErrors
    oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
End Sub

Public Function SegmentIsSloped(pSegments As IJElements, SegIndex As Integer) As Boolean
    Const METHOD = "SegmentIsSloped"
    On Error GoTo ErrHandler
    Dim pLine As IJLine
    Dim pCurve As IJCurve
    Dim x1 As Double, y1 As Double, z1 As Double
    Dim curvescope As Geom3dCurveScopeConstants
    SegmentIsSloped = False
    If SegIndex < 1 Or SegIndex > pSegments.Count Then
        Exit Function
    End If
    If TypeOf pSegments(SegIndex) Is IJLine Then
        Set pLine = pSegments(SegIndex)
        pLine.GetDirection x1, y1, z1
        If Abs(z1) > dtol Then
            SegmentIsSloped = True
        End If
    Else
        Set pCurve = pSegments(SegIndex)
        pCurve.Normal curvescope, x1, y1, z1
        If Abs(z1) > dtol And Abs(x1) < dtol And Abs(y1) < dtol Then    'horizontal curve
        Else
            SegmentIsSloped = True
        End If
    End If

Exit Function
ErrHandler:
    Dim oErrors As New IMSErrorLog.JServerErrors
    oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
End Function

'Modified from the sub CreatePads in TypeATopmounted. The pad is on horrizontal plane for "always vertical" or the plane with the segment path vector for "perpendicular to the path"
'
Public Sub CreatePadByPost(pOC As IJDOutputCollection, post As HandrailPost, SectionAngle As Double, dPadOffset As Double, PostSecWidth As Double, PostSecDepth As Double, ByRef outCount As Long)
    Const METHOD = "CreatePadByPost"
    On Error GoTo ErrHandler
    Dim pIJLine As IJLine
    Dim pCSLine As IJLine
    Dim pLineCS As ComplexString3d
    Dim pCSLineCS As ComplexString3d
    Dim x1 As Double, y1 As Double, z1 As Double
    Dim x2 As Double, y2 As Double, z2 As Double
    Dim x3 As Double, y3 As Double, z3 As Double
    Dim x4 As Double, y4 As Double, z4 As Double
    Dim px As Double, py As Double, pz As Double
'    Dim pX1 As Double, pY1 As Double, pZ1 As Double
    Dim Index As Integer
    Dim startingNorm() As Double
    Dim endingNorm() As Double
    Dim vZ As DVector
    Dim vPost As DVector
    Dim vNorm As DVector
    Dim tmpvSeg As DVector
    Dim vNormTemp As DVector
    Dim Pt1 As DPosition
    Dim Pt2 As DPosition
    Dim Pt3 As DPosition
    Dim IdMatrix As IJDT4x4
    Dim SegVec As DVector
    Dim pProjectionEles As IJElements
    
    Dim dAngle As Double
    Dim dPostSecWidth As Double
    Dim dPostSecDepth As Double
    Dim OutStr As String
    Dim pObj As Object
    OutStr = "PostPad"
    dAngle = SectionAngle
    dPostSecWidth = PostSecWidth / 2
    dPostSecDepth = PostSecDepth / 2
        
    Set vZ = New DVector
    Set vPost = New DVector
    Set vNorm = New DVector
    Set tmpvSeg = New DVector
    Set vNormTemp = New DVector
    Set SegVec = New DVector
    Set IdMatrix = New DT4x4

    vPost.Set post.DirectionVec.x, post.DirectionVec.y, post.DirectionVec.z
    SegVec.Set post.PathVec.x, post.PathVec.y, post.PathVec.z
    vZ.Set vPost.x, vPost.y, vPost.z
    Set vNorm = SegVec.Cross(vZ)
    vNorm.Length = 1#
    Set SegVec = vZ.Cross(vNorm)
    SegVec.Length = 1#
    
    'calculations for handling post orientation angle
    IdMatrix.LoadIdentity
    IdMatrix.Rotate dAngle, vPost
    Set tmpvSeg = IdMatrix.TransformVector(SegVec)
    Set vNormTemp = IdMatrix.TransformVector(vNorm)
    SegVec.Set tmpvSeg.x, tmpvSeg.y, tmpvSeg.z
    vNorm.Set vNormTemp.x, vNormTemp.y, vNormTemp.z
    Set vPost = SegVec.Cross(vNorm)
    
    'Set oGeomFactory = New GeometryFactory
    Set pLineCS = New ComplexString3d
    Set pCSLineCS = New ComplexString3d

'    pcurve.- Line along thickness of plate at center of plate .i.e. center line of plate
    Set Pt1 = New DPosition
    Set Pt2 = New DPosition
    
    post.BasePos.Get px, py, pz
    Pt1.Set px, py, pz
    vPost.Set post.DirectionVec.x, post.DirectionVec.y, post.DirectionVec.z
        vPost.Length = dPadOffset

    Set pIJLine = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, px, py, pz, px + vPost.x, py + vPost.y, pz + vPost.z)
    pLineCS.AddCurve pIJLine, False
    

    'Create the CrossSection Geometry
        SegVec.Length = dPostSecDepth
        vNorm.Length = dPostSecWidth
    tmpvSeg.Set -SegVec.x, -SegVec.y, -SegVec.z
    vNormTemp.Set -vNorm.x, -vNorm.y, -vNorm.z
    
    '  Pt3(2)--Pt2(1)--Pt3(1)
    '            |
    '            |
    '           Pt1
    '            |
    '            |
    '  Pt3(3)--Pt2(2)--Pt3(4)
    
    Set Pt2 = New DPosition
    Set Pt3 = New DPosition
    Set Pt2 = Pt1.Offset(vNorm)
    Set Pt3 = Pt2.Offset(SegVec)
    Pt3.Get x1, y1, z1
    
    Set Pt3 = Pt2.Offset(tmpvSeg)
    Pt3.Get x2, y2, z2
    
    
    Set Pt2 = Pt1.Offset(vNormTemp)
    Set Pt3 = Pt2.Offset(tmpvSeg)
    Pt3.Get x3, y3, z3
    
    Set Pt3 = Pt2.Offset(SegVec)
    Pt3.Get x4, y4, z4
    
    'Line 1
    Set pCSLine = Nothing
    Set pCSLine = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, x1, y1, z1, x2, y2, z2)
    pCSLineCS.AddCurve pCSLine, False
    
    'Line 2
    Set pCSLine = Nothing
    Set pCSLine = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, x2, y2, z2, x3, y3, z3)
    pCSLineCS.AddCurve pCSLine, False
    
    'Line 3
    Set pCSLine = Nothing
    Set pCSLine = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, x3, y3, z3, x4, y4, z4)
    pCSLineCS.AddCurve pCSLine, False
    
    'Line 4
    Set pCSLine = Nothing
    Set pCSLine = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, x4, y4, z4, x1, y1, z1)
    pCSLineCS.AddCurve pCSLine, False
    
    'Send appropriate brkcrv & bcapped agruments based on skinning option selected for handaril
    Dim bcapped As Long
    Dim brkcrv As Long
    Call getSkinningoptions(SkinOption, bcapped, brkcrv)
    Set pProjectionEles = m_GeomFactory.GeometryServices.CreateBySingleSweep(Nothing, pLineCS, pCSLineCS, 0, brkcrv, startingNorm, endingNorm, bcapped)
        
    For Index = 1 To pProjectionEles.Count
        InitNewOutput pOC, OutStr & outCount
        Set pObj = pProjectionEles(Index)
        pOC.AddOutput (OutStr & Trim$(Str$(outCount))), pObj
        Set pObj = Nothing
        outCount = outCount + 1
    Next Index
    
    Set vZ = Nothing
    Set vPost = Nothing
    Set vNorm = Nothing
    Set vNormTemp = Nothing
    Set tmpvSeg = Nothing
    Set pLineCS = Nothing
    Set pIJLine = Nothing
    Set pCSLineCS = Nothing
    Set Pt1 = Nothing
    Set Pt2 = Nothing
    Set Pt3 = Nothing
    Set IdMatrix = Nothing
    Set SegVec = Nothing
        
Exit Sub
ErrHandler:
    Dim oErrors As New IMSErrorLog.JServerErrors
    oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
End Sub

Public Function GetAdjustedCP(OldSectionCP As Integer, Mirror As Boolean) As Integer
'Used for TypeASideMount Handrail only
'The section CP should be changed when the section is rotated 90 degrees (clock wise or counter clock wise)
'   7-8-9                           9-6-3
'   4-5-6   -> (counter clock wise) 8-5-2
'   1-2-3                           7-4-1
'
'   7-8-9                   1-4-7
'   4-5-6   -> (clock wise) 2-5-8
'   1-2-3                   3-6-9
'
    
    If Mirror Then
        Select Case OldSectionCP
            Case 1:
                GetAdjustedCP = 7
            Case 2:
                GetAdjustedCP = 4
            Case 3:
                GetAdjustedCP = 1
            Case 4:
                GetAdjustedCP = 8
            Case 6:
                GetAdjustedCP = 2
            Case 7:
                GetAdjustedCP = 9
            Case 8:
                GetAdjustedCP = 6
            Case 9:
                GetAdjustedCP = 3
            Case Else:
                GetAdjustedCP = OldSectionCP
        End Select
    Else
        Select Case OldSectionCP
            Case 1:
                GetAdjustedCP = 3
            Case 2:
                GetAdjustedCP = 6
            Case 3:
                GetAdjustedCP = 9
            Case 4:
                GetAdjustedCP = 2
            Case 6:
                GetAdjustedCP = 8
            Case 7:
                GetAdjustedCP = 1
            Case 8:
                GetAdjustedCP = 4
            Case 9:
                GetAdjustedCP = 7
            Case Else:
                GetAdjustedCP = OldSectionCP
        End Select
    End If

End Function

Public Function ConvertCP(OldSectionCP As Integer) As Integer
'Used for TypeASideMount Handrail only
'The section CP should be changed when the section is rotated 90 degrees (clock wise or counter clock wise)
'   7-8-9       1-2-3
'   4-5-6   ->  4-5-6
'   1-2-3       7-8-9
'
   
    Select Case OldSectionCP
        Case 1:
            ConvertCP = 7
        Case 2:
            ConvertCP = 8
        Case 3:
            ConvertCP = 9
        Case 7:
            ConvertCP = 1
        Case 8:
            ConvertCP = 2
        Case 9:
            ConvertCP = 3
        Case Else:
            ConvertCP = OldSectionCP
    End Select

End Function

Public Function CreateSlopedCirEndTreatment(SegCurve As IJCurve, _
                                       ResourceManager As Object, _
                                       PostStartX As Double, _
                                       PostStartY As Double, _
                                       PostStartZ As Double, _
                                       MidRailDist As Double, _
                                       TotalHt As Double, _
                                       PostVec As DVector, _
                                       bIsbegin As Boolean, dToprailDepth As Double, Optional dRadius As Double) As ComplexString3d
                                       
Const METHOD = "CreateSlopedCirEndTreatment"
On Error GoTo ErrorHandler
  
    Dim Pt1 As DPosition
    Dim Pt2 As DPosition
    Dim postPt As DPosition
    Dim arcPt As DPosition
    Dim vecOV As DVector
    Set vecOV = New DVector
    vecOV.Set PostVec.x, PostVec.y, PostVec.z
    Set Pt1 = New DPosition
    Set Pt2 = New DPosition
    Set postPt = New DPosition
    Set arcPt = New DPosition
    postPt.Set PostStartX, PostStartY, PostStartZ 'end of the path
    
    Dim ht1 As Double
    Dim htTemp1 As Double
    Dim htTemp2 As Double
    
    If dRadius > dtol Then
        ht1 = dRadius
    Else
        ht1 = (TotalHt - MidRailDist) / 6   ' assumption 1/6 of the total ht = arc radius
    
        'If the above assuption is not a good one, use the following formular to get arc radius
        'This value should be consistent with the value of dCirTOffset in PhysicalRepresentation
        If dToprailDepth > (TotalHt - MidRailDist) Then
            ht1 = (TotalHt - MidRailDist) / 2
        ElseIf ht1 < dToprailDepth / 2 Then
            ht1 = dToprailDepth / 4 + (TotalHt - MidRailDist) / 4
        End If
    End If
    Set CreateSlopedCirEndTreatment = Nothing
    
    Dim ParamOnSegment1 As Double
    Dim ParamOnSegment2 As Double
    Dim StartPar As Double, EndPar As Double
    
    SegCurve.ParamRange StartPar, EndPar

    Dim Arc1X1 As Double, Arc1Y1 As Double, Arc1Z1 As Double
    Dim Arc1X2 As Double, Arc1Y2 As Double, Arc1Z2 As Double
    Dim Arc1X3 As Double, Arc1Y3 As Double, Arc1Z3 As Double
    
    Dim Arc2X1 As Double, Arc2Y1 As Double, Arc2Z1 As Double
    Dim Arc2X2 As Double, Arc2Y2 As Double, Arc2Z2 As Double
    Dim Arc2X3 As Double, Arc2Y3 As Double, Arc2Z3 As Double
    
    Dim xPt As Double, yPt As Double, zPt As Double     'end point
    Dim xPt1 As Double, yPt1 As Double, zPt1 As Double  'end point at bottom mid rail
    Dim xPt2 As Double, yPt2 As Double, zPt2 As Double  'end point at top rail
    Dim vtanx As Double, vtany As Double, vtanz As Double
    Dim vTan2X As Double, vTan2Y As Double, vTan2Z As Double
    Dim dAngle As Double
    Dim vecPath As DVector
    Dim vecProjectPath As DVector
    Dim vecChord As DVector
    Dim vecBisection As DVector
    Dim vecNorm As DVector
    Dim line1 As IJLine
    Dim line2 As IJLine
    Dim point1 As DPosition
    Dim point2 As DPosition
    Dim pointMiddle As DPosition
    Set point1 = New DPosition
    Set point2 = New DPosition
    Set pointMiddle = New DPosition
    Set vecChord = New DVector
    If m_GeomFactory Is Nothing Then
        Set m_GeomFactory = New GeometryFactory
    End If
    
    If bIsbegin Then
        SegCurve.Evaluate StartPar, xPt, yPt, zPt, vtanx, vtany, vtanz, vTan2X, vTan2Y, vTan2Z
    Else
        SegCurve.Evaluate EndPar, xPt, yPt, zPt, vtanx, vtany, vtanz, vTan2X, vTan2Y, vTan2Z
    End If
    
    'The calculation is accurate for linear path
    Set vecPath = New DVector
    vecPath.Set vtanx, vtany, vtanz
    Set vecNorm = vecPath.Cross(vecOV)
    Set vecProjectPath = vecNorm.Cross(vecOV)
    dAngle = PI / 2 - Abs(vecPath.Angle(vecOV, vecNorm))
    If bIsbegin Then
        htTemp1 = ht1 * Tan(PI / 4 + dAngle / 2)
        htTemp2 = ht1 * Tan(PI / 4 - dAngle / 2)
    Else
        htTemp1 = ht1 * Tan(PI / 4 - dAngle / 2)
        htTemp2 = ht1 * Tan(PI / 4 + dAngle / 2)
    End If
        ParamOnSegment1 = (EndPar - StartPar) * htTemp1 / SegCurve.Length
        ParamOnSegment2 = (EndPar - StartPar) * htTemp2 / SegCurve.Length
    
    If bIsbegin Then
        SegCurve.Evaluate StartPar + ParamOnSegment1, xPt1, yPt1, zPt1, vtanx, vtany, vtanz, vTan2X, vTan2Y, vTan2Z
        SegCurve.Evaluate StartPar + ParamOnSegment2, xPt2, yPt2, zPt2, vtanx, vtany, vtanz, vTan2X, vTan2Y, vTan2Z
    Else
        SegCurve.Evaluate EndPar - ParamOnSegment1, xPt1, yPt1, zPt1, vtanx, vtany, vtanz, vTan2X, vTan2Y, vTan2Z
        SegCurve.Evaluate EndPar - ParamOnSegment2, xPt2, yPt2, zPt2, vtanx, vtany, vtanz, vTan2X, vTan2Y, vTan2Z
    End If
    
    'Base position for arc1 start or end
    Pt1.Set xPt1, yPt1, zPt1 'TR#5138
    
    'Arc1 start
        vecOV.Length = MidRailDist
    Set arcPt = Pt1.Offset(vecOV)
    Arc1X1 = arcPt.x      'xPt
    Arc1Y1 = arcPt.y      'yPt
    Arc1Z1 = arcPt.z      'zPt + MidRailDist
    Set point1 = arcPt.Clone
    
    'Arc1 end
        vecOV.Length = MidRailDist + htTemp1
    Set arcPt = postPt.Offset(vecOV)
    
    Arc1X3 = arcPt.x      'PostStartX
    Arc1Y3 = arcPt.y      'PostStartY
    Arc1Z3 = arcPt.z      'PostStartZ + MidRailDist + ht1
    Set point2 = arcPt.Clone
    
    'Arc1 center
    pointMiddle.Set (point1.x + point2.x) / 2, (point1.y + point2.y) / 2, (point1.z + point2.z) / 2
    vecChord.Set point1.x - point2.x, point1.y - point2.y, point1.z - point2.z
        vecChord.Length = 1
    Set vecBisection = vecChord.Cross(vecNorm)
    
    Set line1 = m_GeomFactory.Lines3d.CreateByPtVectLength(Nothing, pointMiddle.x, pointMiddle.y, pointMiddle.z, vecBisection.x, vecBisection.y, vecBisection.z, 1)
    Set line2 = m_GeomFactory.Lines3d.CreateByPtVectLength(Nothing, point2.x, point2.y, point2.z, vecProjectPath.x, vecProjectPath.y, vecProjectPath.z, 1)
    line1.Infinite = True
    line2.Infinite = True
    Set arcPt = GetIntersectPoint(line1, line2)
    Set line1 = Nothing
    Set line2 = Nothing
    
    Arc1X2 = arcPt.x      'Arc1X1
    Arc1Y2 = arcPt.y      'Arc1Y1
    Arc1Z2 = arcPt.z      ' Arc1Z1 + ht1
    
    'Base position for arc2 start or end
    Pt1.Set xPt2, yPt2, zPt2 'TR#5138
    
    'Arc2 start
        vecOV.Length = TotalHt
    Set arcPt = Pt1.Offset(vecOV)
    Arc2X1 = arcPt.x        'xPt
    Arc2Y1 = arcPt.y        'yPt
    Arc2Z1 = arcPt.z        'zPt + TotalHt
    Set point1 = arcPt.Clone
    
    'Arc2 end
        vecOV.Length = TotalHt - htTemp2
    Set arcPt = postPt.Offset(vecOV)
    Arc2X3 = arcPt.x        'PostStartX
    Arc2Y3 = arcPt.y        'PostStartY
    Arc2Z3 = arcPt.z       'PostStartZ + TotalHt - ht1
    Set point2 = arcPt.Clone
        
    'Arc2 center
    pointMiddle.Set (point1.x + point2.x) / 2, (point1.y + point2.y) / 2, (point1.z + point2.z) / 2
    vecChord.Set point1.x - point2.x, point1.y - point2.y, point1.z - point2.z
        vecChord.Length = 1
    Set vecBisection = vecChord.Cross(vecNorm)
    
    Set line1 = m_GeomFactory.Lines3d.CreateByPtVectLength(Nothing, pointMiddle.x, pointMiddle.y, pointMiddle.z, vecBisection.x, vecBisection.y, vecBisection.z, 1)
    Set line2 = m_GeomFactory.Lines3d.CreateByPtVectLength(Nothing, point2.x, point2.y, point2.z, vecProjectPath.x, vecProjectPath.y, vecProjectPath.z, 1)
    line1.Infinite = True
    line2.Infinite = True
    Set arcPt = GetIntersectPoint(line1, line2)
    Set line1 = Nothing
    Set line2 = Nothing
    
    Arc2X2 = arcPt.x        'Arc2X1
    Arc2Y2 = arcPt.y        'Arc2Y1
    Arc2Z2 = arcPt.z        'Arc2Z1 - ht1
    
    Set point1 = Nothing
    Set point2 = Nothing
    
    Dim iElements As IJElements
    Set iElements = New JObjectCollection ' IMSElements.DynElements
    
    Dim oLine As IngrGeom3D.Line3d
    Dim oarc1 As IngrGeom3D.Arc3d, oArc2 As IngrGeom3D.Arc3d
    
    ' create an arc with arc2pt1 as start so that the cross section direction will be consistent with the toprail
    Set oarc1 = m_GeomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, _
                                                        Arc2X2, Arc2Y2, Arc2Z2, _
                                                        Arc2X1, Arc2Y1, Arc2Z1, _
                                                        Arc2X3, Arc2Y3, Arc2Z3)
    If Not oarc1 Is Nothing Then
        iElements.Add oarc1
        oarc1.GetEndPoint Arc2X3, Arc2Y3, Arc2Z3
    End If

    Set oLine = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                                    Arc2X3, Arc2Y3, Arc2Z3, _
                                                    Arc1X3, Arc1Y3, Arc1Z3)
    If Not oLine Is Nothing Then iElements.Add oLine
    Set oArc2 = m_GeomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, _
                                                        Arc1X2, Arc1Y2, Arc1Z2, _
                                                        Arc1X3, Arc1Y3, Arc1Z3, _
                                                        Arc1X1, Arc1Y1, Arc1Z1)
    If Not oArc2 Is Nothing Then iElements.Add oArc2
    
    Dim oCmpx As ComplexString3d
    
    Set oCmpx = m_GeomFactory.ComplexStrings3d.CreateByCurves(Nothing, iElements)

    Set CreateSlopedCirEndTreatment = oCmpx
    
    Set oarc1 = Nothing
    Set oArc2 = Nothing
    Set oLine = Nothing
    Set oCmpx = Nothing
    Set iElements = Nothing
    Set Pt1 = Nothing
    Set postPt = Nothing
    Set arcPt = Nothing
    
    Exit Function
    
ErrorHandler:
    Dim oErrors As New IMSErrorLog.JServerErrors
    oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
End Function

' When the intersection angle between the rail and the vertical treatment part is not 90 degrees,
' it is necessary to adjust the offset distances for the same radius for the turning arc.
' The calculation is based on the assumption that the path for rail is linear.
' For curved path, the approximation is acceptable if the curvature of the path is not too big
' dBottomRailBeginOffset is used for bottom midrail at begin end
' dTopRailBeginOffset is used for toprail at begin end
' dBottomRailEndOffset is used for bottom midrail at end end
' dTopRailEndOffset is used for toprail at end end
' CirTreatOffsetDelta is used when the toprail CP is not centered
Public Sub CalcCirTreatOffsets(ByVal oSegments As IJElements, _
                           ByVal CirTreatOffset As Double, _
                           ByVal CirTreatOffsetDelta As Double, _
                           ByVal Orientation As Integer, _
                           ByRef dBottomRailBeginOffset As Double, _
                           ByRef dTopRailBeginOffset As Double, _
                           ByRef dBottomRailEndOffset As Double, _
                           ByRef dTopRailEndOffset As Double)
    Dim SegCurve As IJCurve
    Dim xPt As Double, yPt As Double, zPt As Double
    Dim vtanx As Double, vtany As Double, vtanz As Double
    Dim vTan2X As Double, vTan2Y As Double, vTan2Z As Double
    Dim dAngle As Double
    Dim StartPar As Double, EndPar As Double
    Dim vecPath As DVector
    Dim vecNorm As DVector
    Dim vecOV As DVector
    Set vecOV = New DVector
    vecOV.Set 0, 0, 1
    
    'initialized with original value
    dBottomRailBeginOffset = CirTreatOffset
    dTopRailBeginOffset = CirTreatOffset
    dBottomRailEndOffset = CirTreatOffset
    dTopRailEndOffset = CirTreatOffset
    
    If Orientation = 1 Then Exit Sub
    
    Set SegCurve = oSegments.Item(1)
    SegCurve.ParamRange StartPar, EndPar

    SegCurve.Evaluate StartPar, xPt, yPt, zPt, vtanx, vtany, vtanz, vTan2X, vTan2Y, vTan2Z
    Set vecPath = New DVector
    If Abs(vtanz) > dtol Then
        vecPath.Set vtanx, vtany, vtanz
        Set vecNorm = vecPath.Cross(vecOV)
        dAngle = PI / 2 - Abs(vecPath.Angle(vecOV, vecNorm))
        dBottomRailBeginOffset = CirTreatOffset * Tan(PI / 4 + dAngle / 2) - CirTreatOffsetDelta * Sin(dAngle)
        dTopRailBeginOffset = CirTreatOffset * Tan(PI / 4 - dAngle / 2)
    End If
    
    Set SegCurve = oSegments.Item(oSegments.Count)
    SegCurve.ParamRange StartPar, EndPar

    SegCurve.Evaluate EndPar, xPt, yPt, zPt, vtanx, vtany, vtanz, vTan2X, vTan2Y, vTan2Z
    If Abs(vtanz) > dtol Then
        vecPath.Set vtanx, vtany, vtanz
        Set vecNorm = vecPath.Cross(vecOV)
        dAngle = PI / 2 - Abs(vecPath.Angle(vecOV, vecNorm))
        dTopRailEndOffset = CirTreatOffset * Tan(PI / 4 + dAngle / 2)
        dBottomRailEndOffset = CirTreatOffset * Tan(PI / 4 - dAngle / 2) + CirTreatOffsetDelta * Sin(dAngle)
    End If

End Sub


Public Sub CreateTopRail_Ex(pOutputParentObject As Object, _
                          ByVal pComplex As ComplexString3d, _
                          ByVal DirPt As DPosition, _
                          Height As Double, _
                          Orientation As Integer, _
                          TopRailSectionName As String, _
                          TopRailSectionStandard As String, _
                          ByVal CSectionCP As Integer, _
                          ByVal CSectionAngle As Double, _
                          CirTreatOffset As Double, _
                          BeginTreatType As Integer, _
                          EndTreatType As Integer, _
                          Mirror As Boolean)
                          

Const METHOD = "CreateTopRail"
On Error GoTo ErrorHandler
    Dim nCreated As Long
    Dim strOutputName As String
    Dim oTrans4x4       As IJDT4x4
    Dim oVector         As IJDVector
    Dim pCurve          As ComplexString3d
    Dim oSegments       As IJElements
    
    Dim oHandrail       As ISPSHandrail
    Dim pOC             As IJDOutputCollection
    Dim dBottomRailBeginOffset As Double
    Dim dTopRailBeginOffset As Double
    Dim dBottomRailEndOffset As Double
    Dim dTopRailEndOffset As Double
    
    strOutputName = "TopRail"

    If TypeOf pOutputParentObject Is IJDOutputCollection Then
        Set pOC = pOutputParentObject
    Else
        Set oHandrail = pOutputParentObject
    End If

    If m_GeomFactory Is Nothing Then
        Set m_GeomFactory = New GeometryFactory
    End If
    
    If Orientation Then
        Dim pIJCurve As IJCurve
        Set pIJCurve = pComplex
        
        If pIJCurve.Scope = CURVE_SCOPE_PLANAR Then
            On Error Resume Next
            Set pComplex = m_GeomFactory.GeometryServices.CreateByOffset(Nothing, pComplex, _
                                                                        DirPt.x, DirPt.y, DirPt.z, Height, 1)
            On Error GoTo ErrorHandler
        ElseIf ComplexstringcontainsArc(pComplex) = False Then
            'TR#51383- use this function because CreateByOffset doesn't work correctly for non-planar path
            Set pComplex = GetPathByOffset(pComplex, Height, 0) 'TR#53722- last arg added to indicate direction of offset
        End If
    End If
    
    Set oVector = New DVector
    oVector.Set 0, 0, Height
    
    Set oTrans4x4 = New DT4x4
    oTrans4x4.LoadIdentity
    oTrans4x4.Translate oVector

    pComplex.GetCurves oSegments
    
    
    If BeginTreatType = 5 Or EndTreatType = 5 Then
        CalcCirTreatOffsets oSegments, CirTreatOffset, 0, Orientation, dBottomRailBeginOffset, dTopRailBeginOffset, dBottomRailEndOffset, dTopRailEndOffset
        
        ' need to offset the first and last segments of the input curve
        
        Dim oCurve1 As IJCurve, oCurveN As IJCurve
        Dim oline1 As IJLine, olineN As IJLine
        Dim oarc1 As IngrGeom3D.Arc3d, oarcN As IngrGeom3D.Arc3d
        Dim bIsseg1Line As Boolean, bIssegNLine As Boolean
        Dim SPar As Double, EPar As Double, newPar As Double
        Dim n As Integer
        Dim newEx1 As Double, newEy1 As Double, newEz1 As Double
        Dim newExN As Double, newEyN As Double, newEzN As Double
        Dim newSegs     As IJElements
        
        If TypeOf oSegments.Item(1) Is IJLine Then
            Set oline1 = oSegments.Item(1)
            bIsseg1Line = True
        ElseIf TypeOf oSegments.Item(1) Is Arc3d Then
            Set oarc1 = oSegments.Item(1)
        End If
        
        n = oSegments.Count
        If TypeOf oSegments.Item(n) Is IJLine Then
            Set olineN = oSegments.Item(n)
            bIssegNLine = True
        ElseIf TypeOf oSegments.Item(n) Is Arc3d Then
            Set oarcN = oSegments.Item(n)
        End If

        If oline1 Is Nothing And oarc1 Is Nothing Then
            Err.Raise E_FAIL, METHOD, oLocalizer.GetString(IDS_EDIT_HANDRAIL_PATH, "Handrail path segment is not a line or arc. Edit handrail path and redefine all path segments to be either a line or an arc.")
        ElseIf olineN Is Nothing And oarcN Is Nothing Then
            Err.Raise E_FAIL, METHOD, oLocalizer.GetString(IDS_EDIT_HANDRAIL_PATH, "Handrail path segment is not a line or arc. Edit handrail path and redefine all path segments to be either a line or an arc.")
        ElseIf oSegments.Count = 1 Then     ' only one segment
            Set newSegs = New JObjectCollection  'IMSElements.DynElements
            
            If bIsseg1Line Then         ' line
                Set oCurve1 = oline1
                oCurve1.ParamRange SPar, EPar

                If BeginTreatType = 5 Then      ' change the begin point
                        newPar = (EPar - SPar) * dTopRailBeginOffset / oCurve1.Length
                    oCurve1.Position SPar + newPar, newEx1, newEy1, newEz1
                    If oCurve1.IsPointOn(newEx1, newEy1, newEz1) Then
                        oline1.SetStartPoint newEx1, newEy1, newEz1
                        Set oCurve1 = oline1
                        oCurve1.ParamRange SPar, EPar
                    End If
                    Set oCurve1 = oline1
                    oCurve1.ParamRange SPar, EPar
                End If

                If EndTreatType = 5 Then        ' change the end point
                        newPar = (EPar - SPar) * dTopRailEndOffset / oCurve1.Length
                    oCurve1.Position EPar - newPar, newExN, newEyN, newEzN
                    If oCurve1.IsPointOn(newExN, newEyN, newEzN) Then oline1.SetEndPoint newExN, newEyN, newEzN
                End If

                newSegs.Add oline1
                Set pCurve = m_GeomFactory.ComplexStrings3d.CreateByCurves(Nothing, newSegs)
                Set oline1 = Nothing
                Set oCurve1 = Nothing
                Set newSegs = Nothing

            Else                        ' arc
                Set oCurve1 = oarc1
                oCurve1.ParamRange SPar, EPar
                If BeginTreatType = 5 Then      ' change the begin point
                        newPar = (EPar - SPar) * dTopRailBeginOffset / oCurve1.Length
                    oCurve1.Position SPar + newPar, newEx1, newEy1, newEz1
                    If oCurve1.IsPointOn(newEx1, newEy1, newEz1) Then
                        oarc1.SetStartPoint newEx1, newEy1, newEz1
                        Set oCurve1 = oarc1
                        oCurve1.ParamRange SPar, EPar
                    End If
                    Set oCurve1 = oarc1
                    oCurve1.ParamRange SPar, EPar
                End If

                If EndTreatType = 5 Then        ' change the end point
                        newPar = (EPar - SPar) * dTopRailEndOffset / oCurve1.Length
                    oCurve1.Position EPar - newPar, newExN, newEyN, newEzN
                    If oCurve1.IsPointOn(newExN, newEyN, newEzN) Then
                        oarc1.SetEndPoint newExN, newEyN, newEzN
                    End If
                End If
                newSegs.Add oarc1
                Set pCurve = m_GeomFactory.ComplexStrings3d.CreateByCurves(Nothing, newSegs)
                Set oarc1 = Nothing
                Set oCurve1 = Nothing
                Set newSegs = Nothing
            End If

        Else    ' multiple segments need to apply treatments to first and last segments
'                ' handle first segment
            Set pCurve = m_GeomFactory.ComplexStrings3d.CreateByCurves(Nothing, oSegments)
            If bIsseg1Line Then         ' line

                Set oCurve1 = oline1
                oCurve1.ParamRange SPar, EPar
                If BeginTreatType = 5 Then      ' change the begin point
                        newPar = (EPar - SPar) * dTopRailBeginOffset / oCurve1.Length
                    oCurve1.Position SPar + newPar, newEx1, newEy1, newEz1
                    If oCurve1.IsPointOn(newEx1, newEy1, newEz1) Then
                        oline1.SetStartPoint newEx1, newEy1, newEz1
                        pCurve.RemoveCurve False
                        pCurve.AddCurve oline1, False
                    End If
                End If
                Set oCurve1 = Nothing
                Set oline1 = Nothing
            Else                        ' arc
                Set oCurve1 = oarc1
                oCurve1.ParamRange SPar, EPar

                If BeginTreatType = 5 Then      ' change the begin point
                        newPar = (EPar - SPar) * dTopRailBeginOffset / oCurve1.Length
                    oCurve1.Position SPar + newPar, newEx1, newEy1, newEz1
                    If oCurve1.IsPointOn(newEx1, newEy1, newEz1) Then
                        oarc1.SetStartPoint newEx1, newEy1, newEz1
                        pCurve.RemoveCurve False
                        pCurve.AddCurve oarc1, False
                    End If
                End If
                Set oCurve1 = Nothing
                Set oarc1 = Nothing
            End If
'
'                ' handle last segment
            If bIssegNLine Then         ' line
                Set oCurveN = olineN
                oCurveN.ParamRange SPar, EPar

                If EndTreatType = 5 Then      ' change the end point
                        newPar = (EPar - SPar) * dTopRailEndOffset / oCurveN.Length
                    oCurveN.Position EPar - newPar, newExN, newEyN, newEzN
                    If oCurveN.IsPointOn(newExN, newEyN, newEzN) Then
                        olineN.SetEndPoint newExN, newEyN, newEzN
                        pCurve.RemoveCurve True
                        pCurve.AddCurve olineN, True
                     End If
                End If
                Set oCurveN = Nothing
                Set olineN = Nothing
            Else                        ' arc
                Set oCurveN = oarcN
                oCurveN.ParamRange SPar, EPar

                If EndTreatType = 5 Then      ' change the begin point
                        newPar = (EPar - SPar) * dTopRailEndOffset / oCurveN.Length
                    oCurveN.Position EPar - newPar, newExN, newEyN, newEzN
                    If oCurveN.IsPointOn(newExN, newEyN, newEzN) Then
                        oarcN.SetEndPoint newExN, newEyN, newEzN
                        pCurve.RemoveCurve True
                        pCurve.AddCurve oarcN, True
                    End If
                End If
                Set oCurveN = Nothing
                Set oarcN = Nothing
            End If
        End If
    Else
        Set pCurve = m_GeomFactory.ComplexStrings3d.CreateByCurves(Nothing, oSegments)
    End If
    
    If Orientation = 0 Then   'TR#51383- Transform only if orientation is 0 otherwise we already have at top rail position
        pCurve.Transform oTrans4x4
    End If
    
    Set oSegments = Nothing
    Set oTrans4x4 = Nothing
    Set oVector = Nothing
    
    If Not pOC Is Nothing Then   ' symbol
    
        Dim pCSProfileObj As Object
        If Trim(TopRailSectionStandard) <> "" And Trim(TopRailSectionName) <> "" Then
            Set pCSProfileObj = GetCSProfile(Nothing, TopRailSectionStandard, TopRailSectionName, m_oCatResMgr)
        End If
    
        BuildHandrailOutput pOC, pCurve, pCSProfileObj, CSectionCP, CSectionAngle, Mirror, SkinOption, strOutputName, MemberType_TopRail, nCreated

        If Not pCSProfileObj Is Nothing Then
            Dim otmp As iJDObject
            Set otmp = pCSProfileObj
            otmp.Remove
            Set otmp = Nothing
        End If
    Else   ' Handrail Convert

        Dim oSectionDef As Object       ' catalog cross-section object
        Dim xService As SP3DStructGenericTools.CrossSectionServices
        
        Set xService = New SP3DStructGenericTools.CrossSectionServices

        xService.GetStructureCrossSectionDefinition GetCatalogResourceManager, TopRailSectionStandard, "", TopRailSectionName, oSectionDef

        BuildHandrailOutput oHandrail, pCurve, oSectionDef, CSectionCP, CSectionAngle, Mirror, SkinOption, strOutputName, MemberType_TopRail, nCreated
    
    End If

cleanup:
    Set oSegments = Nothing
    Set oTrans4x4 = Nothing
    Set oVector = Nothing
    Set pCurve = Nothing
    Set pCSProfileObj = Nothing
    
    Exit Sub
    
ErrorHandler:
    Dim oErrors As New IMSErrorLog.JServerErrors
    oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
End Sub

Public Sub CreateMidRails_Ex(pOutputParentObject As Object, _
                           ByVal pComplex As ComplexString3d, _
                           ByVal DirPt As DPosition, _
                           Orientation As Integer, _
                           noofmidrails As Integer, _
                           midrailsplaced As Integer, _
                           MidRailSpacing As Double, _
                           TopOfMidRailDim As Double, _
                           TopOfToePlateDim As Double, _
                           MidRailSection As String, _
                           MidRailSecStandard As String, _
                           ByRef lowestMidRailHt As Double, _
                           ByRef sPt As DPosition, _
                           ByRef ePt As DPosition, _
                           CSectionCP As Integer, _
                           CSectionAngle As Double, _
                           lastMRHt As Double, _
                           CirTreatOffset As Double, _
                           BeginTreatType As Integer, _
                           EndTreatType As Integer, _
                           Mirror As Boolean, _
                           CirTreatOffsetDelta As Double)       'CirTreatOffsetDelta is used for adjustment due the CP of toprail
    Const METHOD = "CreateMidRails"
    On Error GoTo ErrorHandler

    If m_GeomFactory Is Nothing Then
        Set m_GeomFactory = New GeometryFactory
    End If

    Dim strOutput As String
    Dim pHt             As Double
    Dim j               As Integer
    Dim oSegments       As IJElements
    Dim outindex        As Long
    Dim newSegs         As IJElements
    
    Dim oHandrail       As ISPSHandrail
    Dim pOC             As IJDOutputCollection
    Dim nOutputsCreated As Long
    
    Dim oSectionOcc     As Object       ' cross section occurrence used by symbol
    Dim oSectionDef     As Object       ' cross-section definition used by members
    Dim oSectionObj     As Object       ' common cross section obj sent to BuildHandrailOutput
    
    Dim iCurve          As IJCurve
    
    Dim dBottomRailBeginOffset As Double
    Dim dTopRailBeginOffset As Double
    Dim dBottomRailEndOffset As Double
    Dim dTopRailEndOffset As Double

    strOutput = "MidRail"
    nOutputsCreated = 0

    If TypeOf pOutputParentObject Is IJDOutputCollection Then
        Set pOC = pOutputParentObject
    Else
        Set oHandrail = pOutputParentObject
    End If
    
    
    pHt = TopOfMidRailDim
    pComplex.GetCurves oSegments
    
    If BeginTreatType = 5 Or EndTreatType = 5 Then
        CalcCirTreatOffsets oSegments, CirTreatOffset, CirTreatOffsetDelta, Orientation, dBottomRailBeginOffset, dTopRailBeginOffset, dBottomRailEndOffset, dTopRailEndOffset
    End If
    outindex = 1
    
    If Not pOC Is Nothing Then
    
        If Trim(MidRailSecStandard) <> "" And Trim(MidRailSection) <> "" Then
            Set oSectionOcc = GetCSProfile(Nothing, MidRailSecStandard, MidRailSection, m_oCatResMgr)
            Set oSectionObj = oSectionOcc
        End If
    Else

        Dim xService As SP3DStructGenericTools.CrossSectionServices
        Set xService = New SP3DStructGenericTools.CrossSectionServices

        xService.GetStructureCrossSectionDefinition GetCatalogResourceManager, MidRailSecStandard, "", MidRailSection, oSectionDef
        Set oSectionObj = oSectionDef
    End If
    
    For j = 1 To noofmidrails
    
      Dim oCurve          As ComplexString3d
      
      If (pHt > TopOfToePlateDim) And (pHt > 0) Then
    
        Dim x0 As Double, y0 As Double, z0 As Double
        Dim xN As Double, yN As Double, zN As Double
            
        If Orientation Then
            Dim pIJCurve As IJCurve
            Set pIJCurve = pComplex
            If pIJCurve.Scope = CURVE_SCOPE_PLANAR Then
                On Error Resume Next
                Set oCurve = m_GeomFactory.GeometryServices.CreateByOffset(Nothing, pComplex, _
                                                                           DirPt.x, DirPt.y, DirPt.z, pHt, 1)
                On Error GoTo ErrorHandler
            ElseIf ComplexstringcontainsArc(pComplex) = False Then
                'TR#51383- use this function because CreateByOffset doesn't work correctly for non-planar path
                Set oCurve = GetPathByOffset(pComplex, pHt, 0)  'TR#53722- last arg added to indicate direction of offset
            End If
            Set oSegments = Nothing
            If Not oCurve Is Nothing Then oCurve.GetCurves oSegments
        End If 'TR#51383

        If Not oSegments Is Nothing Then
            Dim oTrans4x4       As IJDT4x4
            Dim oVector         As IJDVector
            
            Set oCurve = New ComplexString3d
            
            If Abs(pHt - lastMRHt) > 0.00001 Then   ' MidRailSpacing
                Set oCurve = m_GeomFactory.ComplexStrings3d.CreateByCurves(Nothing, oSegments)
            Else
               If BeginTreatType = 5 Or EndTreatType = 5 Then

                    ' need to offset the first and last segments of the input curve
    
                    Dim oCurve1 As IJCurve, oCurveN As IJCurve
                    Dim oline1 As IJLine, olineN As IJLine
                    Dim oarc1 As Arc3d, oarcN As Arc3d
                    Dim bIsseg1Line As Boolean, bIssegNLine As Boolean
                    Dim SPar As Double, EPar As Double, newPar As Double
                    Dim n As Integer
                    Dim newEx1 As Double, newEy1 As Double, newEz1 As Double
                    Dim newExN As Double, newEyN As Double, newEzN As Double
    
                    If TypeOf oSegments.Item(1) Is IJLine Then
                        Set oline1 = oSegments.Item(1)
                        bIsseg1Line = True
                    ElseIf TypeOf oSegments.Item(1) Is Arc3d Then
                        Set oarc1 = oSegments.Item(1)
                    End If
    
                    n = oSegments.Count
                    If TypeOf oSegments.Item(n) Is IJLine Then
                        Set olineN = oSegments.Item(n)
                        bIssegNLine = True
                    ElseIf TypeOf oSegments.Item(n) Is Arc3d Then
                        Set oarcN = oSegments.Item(n)
                    End If
    
                    If (oline1 Is Nothing And oarc1 Is Nothing) Then
                        Err.Raise E_FAIL, METHOD, oLocalizer.GetString(IDS_EDIT_HANDRAIL_PATH, "Handrail path segment is not a line or arc. Edit handrail path and redefine all path segments to be either a line or an arc.")
                    ElseIf (olineN Is Nothing And oarcN Is Nothing) Then
                        Err.Raise E_FAIL, METHOD, oLocalizer.GetString(IDS_EDIT_HANDRAIL_PATH, "Handrail path segment is not a line or arc. Edit handrail path and redefine all path segments to be either a line or an arc.")
                    ElseIf oSegments.Count = 1 Then     ' only one segment

                        If bIsseg1Line Then         ' line

                            Set oCurve1 = oline1
                            oCurve1.ParamRange SPar, EPar

                            If BeginTreatType = 5 Then      ' change the begin point
                                        newPar = (EPar - SPar) * dBottomRailBeginOffset / oCurve1.Length
                                oCurve1.Position SPar + newPar, newEx1, newEy1, newEz1
                                If oCurve1.IsPointOn(newEx1, newEy1, newEz1) Then
                                    oline1.SetStartPoint newEx1, newEy1, newEz1
                                    Set oCurve1 = oline1
                                    oCurve1.ParamRange SPar, EPar
                                End If
                            End If

                            If EndTreatType = 5 Then        ' change the end point
                                        newPar = (EPar - SPar) * dBottomRailEndOffset / oCurve1.Length
                                oCurve1.Position EPar - newPar, newExN, newEyN, newEzN
                                If oCurve1.IsPointOn(newExN, newEyN, newEzN) Then oline1.SetEndPoint newExN, newEyN, newEzN
                            End If

                            Set newSegs = New JObjectCollection ' IMSElements.DynElements
                            newSegs.Add oline1
                            Set oCurve = m_GeomFactory.ComplexStrings3d.CreateByCurves(Nothing, newSegs)
                            Set oline1 = Nothing
                            Set oCurve1 = Nothing
                            Set newSegs = Nothing

                        Else                        ' arc

                            Set oCurve1 = oarc1
                            oCurve1.ParamRange SPar, EPar

                            If BeginTreatType = 5 Then      ' change the begin point
                                        newPar = (EPar - SPar) * dBottomRailBeginOffset / oCurve1.Length
                                oCurve1.Position SPar + newPar, newEx1, newEy1, newEz1
                                If oCurve1.IsPointOn(newEx1, newEy1, newEz1) Then
                                    oarc1.SetStartPoint newEx1, newEy1, newEz1
                                    Set oCurve1 = oarc1
                                    oCurve1.ParamRange SPar, EPar
                                End If
                            End If

                            If EndTreatType = 5 Then        ' change the end point
                                        newPar = (EPar - SPar) * dBottomRailEndOffset / oCurve1.Length
                                oCurve1.Position EPar - newPar, newExN, newEyN, newEzN
                                If oCurve1.IsPointOn(newExN, newEyN, newEzN) Then oarc1.SetEndPoint newExN, newEyN, newEzN
                            End If

                            Set newSegs = New JObjectCollection  'IMSElements.DynElements
                            newSegs.Add oarc1
                            Set oCurve = m_GeomFactory.ComplexStrings3d.CreateByCurves(Nothing, newSegs)
                            Set oarc1 = Nothing
                            Set oCurve1 = Nothing
                            Set newSegs = Nothing

                        End If

                    Else    ' multiple segments need to apply treatments to first and last segments
        '                ' handle first segment
                        Set oCurve = m_GeomFactory.ComplexStrings3d.CreateByCurves(Nothing, oSegments)

                        If bIsseg1Line Then         ' line

                            Set oCurve1 = oline1
                            oCurve1.ParamRange SPar, EPar

                            If BeginTreatType = 5 Then      ' change the begin point
                                        newPar = (EPar - SPar) * dBottomRailBeginOffset / oCurve1.Length
                                oCurve1.Position SPar + newPar, newEx1, newEy1, newEz1
                                If oCurve1.IsPointOn(newEx1, newEy1, newEz1) Then
                                    oline1.SetStartPoint newEx1, newEy1, newEz1
                                    oCurve.RemoveCurve False
                                    oCurve.AddCurve oline1, False
                                End If
                            End If
                            Set oCurve1 = Nothing
                            Set oline1 = Nothing
                        Else                        ' arc
                            Set oCurve1 = oarc1
                            oCurve1.ParamRange SPar, EPar


                            If BeginTreatType = 5 Then      ' change the begin point
                                        newPar = (EPar - SPar) * dBottomRailBeginOffset / oCurve1.Length
                                oCurve1.Position SPar + newPar, newEx1, newEy1, newEz1
                                If oCurve1.IsPointOn(newEx1, newEy1, newEz1) Then
                                    oarc1.SetStartPoint newEx1, newEy1, newEz1
                                    oCurve.RemoveCurve False
                                    oCurve.AddCurve oarc1, False
                                End If
                            End If
                            Set oCurve1 = Nothing
                            Set oarc1 = Nothing
                        End If
        '
        '                ' handle last segment
                        If bIssegNLine Then         ' line
                            Set oCurveN = olineN
                            oCurveN.ParamRange SPar, EPar

                            If EndTreatType = 5 Then      ' change the end point
                                        newPar = (EPar - SPar) * dBottomRailEndOffset / oCurveN.Length
                                oCurveN.Position EPar - newPar, newExN, newEyN, newEzN
                                If oCurveN.IsPointOn(newExN, newEyN, newEzN) Then
                                    olineN.SetEndPoint newExN, newEyN, newEzN
                                    oCurve.RemoveCurve True
                                    oCurve.AddCurve olineN, True
                                 End If
                            End If
                            Set oCurveN = Nothing
                            Set olineN = Nothing
                        Else                        ' arc
                            Set oCurveN = oarcN
                            oCurveN.ParamRange SPar, EPar

                            If EndTreatType = 5 Then      ' change the begin point
                                        newPar = (EPar - SPar) * dBottomRailEndOffset / oCurveN.Length
                                oCurveN.Position EPar - newPar, newExN, newEyN, newEzN
                                If oCurveN.IsPointOn(newExN, newEyN, newEzN) Then
                                    oarcN.SetEndPoint newExN, newEyN, newEzN
                                    oCurve.RemoveCurve True
                                    oCurve.AddCurve oarcN, True
                                End If
                            End If
                            Set oCurveN = Nothing
                            Set oarcN = Nothing
                        End If
                    End If
                End If
            End If
            
            Set oVector = New DVector
            oVector.Set 0, 0, pHt
            
            Set oTrans4x4 = New DT4x4
            oTrans4x4.LoadIdentity
            oTrans4x4.Translate oVector
    
            If Orientation = 0 Then 'TR#51383- Transform only if orientation is 0 otherwise we already have at mid rail position
                oCurve.Transform oTrans4x4
            End If
    
            Set oTrans4x4 = Nothing
            Set oVector = Nothing
            BuildHandrailOutput pOutputParentObject, oCurve, oSectionObj, CSectionCP, CSectionAngle, Mirror, SkinOption, strOutput, MemberType_MidRail, nOutputsCreated

            midrailsplaced = midrailsplaced + 1
        
            If MidRailSpacing <> 0 Then
                pHt = pHt - MidRailSpacing
            Else
                Exit Sub
            End If
        
        End If
      End If
      
    Next j
    ' return the end points of the last midrail curve
    Set iCurve = oCurve
    iCurve.EndPoints x0, y0, z0, xN, yN, zN

    
    sPt.Set x0, y0, z0
    ePt.Set xN, yN, zN
    
cleanup:
    Set oCurve = Nothing
    Set oSegments = Nothing
    
    If Not oSectionOcc Is Nothing Then
        Dim iIJDObject As iJDObject
        Set iIJDObject = oSectionOcc
        iIJDObject.Remove
        Set iIJDObject = Nothing
    End If

    
    If noofmidrails > 1 And midrailsplaced > 1 Then
        lowestMidRailHt = pHt + MidRailSpacing
    End If
    
    Exit Sub
    
ErrorHandler:
    Dim oErrors As New IMSErrorLog.JServerErrors
    oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
End Sub


Public Sub CreateToePlate_Ex(oOutputParentObject As Object, _
                          ByVal pComplex As ComplexString3d, _
                          ByVal DirPt As DPosition, _
                          TopOfToePlateDim As Double, _
                          Orientation As Integer, _
                          ToePlateSection As String, _
                          ToePlateSecStandard As String, _
                          CSectionCP As Integer, _
                          CSectionAngle As Double, _
                          Mirror As Boolean)

Const METHOD = "CreateToePlate"
On Error GoTo ErrorHandler

    Dim pOC             As IJDOutputCollection
    Dim strOutput       As String
    Dim nOutputsCreated As Long
    
    Dim oSectionOcc     As Object       ' cross section occurrence used by symbol
    Dim oSectionDef     As Object       ' cross-section definition used by members
    Dim oSectionObj     As Object       ' common cross section obj sent to BuildHandrailOutput
    
    strOutput = "ToePlate"
    nOutputsCreated = 0
    If m_GeomFactory Is Nothing Then
        Set m_GeomFactory = New GeometryFactory
    End If

    If TopOfToePlateDim > 0# Then
        If Orientation Then
            Dim pIJCurve As IJCurve
            Set pIJCurve = pComplex
            If pIJCurve.Scope = CURVE_SCOPE_PLANAR Then
            
                On Error Resume Next
                Set pComplex = m_GeomFactory.GeometryServices.CreateByOffset(Nothing, pComplex, _
                                                        DirPt.x, DirPt.y, DirPt.z, TopOfToePlateDim, 0)
                On Error GoTo ErrorHandler
            ElseIf ComplexstringcontainsArc(pComplex) = False Then
                Set pComplex = GetPathByOffset(pComplex, TopOfToePlateDim, 0)
            End If
        Else
            Dim oVector As IJDVector
            Set oVector = New DVector
            oVector.Set 0, 0, TopOfToePlateDim
            Dim oTrans4x4 As IJDT4x4
            Set oTrans4x4 = New DT4x4
            oTrans4x4.LoadIdentity
            oTrans4x4.Translate oVector
            pComplex.Transform oTrans4x4
        End If
    End If

    If TypeOf oOutputParentObject Is IJDOutputCollection Then
        If Trim(ToePlateSecStandard) <> "" And Trim(ToePlateSection) <> "" Then
            Set pOC = oOutputParentObject
            Set oSectionOcc = GetCSProfile(Nothing, ToePlateSecStandard, ToePlateSection, m_oCatResMgr)
            Set oSectionObj = oSectionOcc
        End If
    Else

        Dim xService As SP3DStructGenericTools.CrossSectionServices
        Set xService = New SP3DStructGenericTools.CrossSectionServices

        xService.GetStructureCrossSectionDefinition GetCatalogResourceManager, ToePlateSecStandard, "", ToePlateSection, oSectionDef
        Set oSectionObj = oSectionDef

    End If

    BuildHandrailOutput oOutputParentObject, pComplex, oSectionObj, CSectionCP, CSectionAngle, Mirror, SkinOption, strOutput, MemberType_ToePlate, nOutputsCreated
    
cleanup:
    If Not oSectionOcc Is Nothing Then
        Dim iIJDObject As iJDObject
        Set iIJDObject = oSectionOcc
        iIJDObject.Remove
        Set iIJDObject = Nothing
    End If
    
    Exit Sub
    
ErrorHandler:
    Dim oErrors As New IMSErrorLog.JServerErrors
    oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
End Sub

Public Function GetHintPoint(curveToEvaluate As IJCurve, onRight As Boolean) As IJDPosition
    Const METHOD = "GetHintPoint"
    On Error GoTo ErrorHandler
    ' This function computes a point on the left or right (assuming a right-hand rule) of the input curve
    ' by projecting the mid-point of that curve's first segment.  For a closed, polygonal curve the start point,
    ' if used, could project back onto the curve itself so the mid point is cused instead:
    '  *_____  <-- start point
    '  |     |       /|\
    '  |     |        |
    '  |     |  offset distance
    '  |     |        |
    ' \|/    |       \|/
    '  *     | <-- offset point
    '  |     |
    '  |_____|
    
    ' get the normal from the curve
    Dim curvescope As Geom3dCurveScopeConstants
    Dim nX As Double
    Dim nY As Double
    Dim nZ As Double
   
    curveToEvaluate.Normal curvescope, nX, nY, nZ
    If curvescope <> CURVE_SCOPE_PLANAR Then
        Err.Raise E_FAIL, "SPSHandrailMacros.ComputeStuff", "Method is valid only for planar curves"
    End If
    ' see if we're dealing with a complex string or just a curve
    Dim curveSegment As IJCurve
    If TypeOf curveToEvaluate Is ComplexString3d Then
        Dim cpx As IJComplexString
        Set cpx = curveToEvaluate
        cpx.GetCurve 1, curveSegment
    Else
        Set curveSegment = curveToEvaluate
    End If
     
    'get the  parameter at the mid-point of the first (and potentially only) segment of the input curve
    Dim startParam As Double
    Dim endParam As Double
    curveSegment.ParamRange startParam, endParam
    
    'evaluate the segment at its mid point to get its tangent
    Dim MidX As Double
    Dim MidY As Double
    Dim MidZ As Double
    Dim tanX As Double
    Dim tanY As Double
    Dim tanZ As Double
    Dim tan2x As Double
    Dim tan2y As Double
    Dim tan2z As Double
    curveSegment.Evaluate (endParam - startParam) / 2, _
                        MidX, _
                        MidY, _
                        MidZ, _
                        tanX, _
                        tanY, _
                        tanZ, _
                        tan2x, _
                        tan2y, _
                        tan2z
                        
   'cross the curve normal with the tangent do get the "left" direction
   Dim NormVec As IJDVector
   Dim tanVec As IJDVector
   Set NormVec = New DVector
   Set tanVec = New DVector
   
   Dim dirVec As IJDVector
   NormVec.Set nX, nY, nZ
   tanVec.Set tanX, tanY, tanZ
    ' for closed curves, the normal returned by curveToEvaluate.Normal is
   ' an oriented normal, but we want to interpret "Right" or "Left"
   ' looking at the curve from above.  This way, "right" for a closed
   ' curve will always be "inside"  whether the user drew it clockwise
   ' or counter clockwise.
   
    If curveToEvaluate.Form <> CURVE_FORM_OPEN Then
            NormVec.Length = NormVec.Length * -1
    End If
   Set dirVec = NormVec.Cross(tanVec)
   
   ' get the position at the mid parameter
   Dim midPos As IJDPosition
   Set midPos = New DPosition
   midPos.Set MidX, MidY, MidZ
   
   ' set the length of the dirVec to a small value (+- 1 mm).
   
    ' we  do not want to offset the point to the wrong side of
    ' the overall curve in case of an acute angle, so set the
    ' distance to a small value
    ' __*___  <-- mid point
    ' \ |   |       /|\
    '  \|   |        |
    '   \___|  offset distance
    '   |            |
    '  \|/          \|/
    '   *     <-- offset point
    If True = onRight Then
            dirVec.Length = -0.001
    Else
            dirVec.Length = 0.001
    End If


    ' project the mid point along the direction vector
    Dim projectedPoint As IJDPosition
    Set projectedPoint = midPos.Offset(dirVec)
    Debug.Print "Mid Position: (" & CStr(midPos.x) + ", " + CStr(midPos.y) + ", " + CStr(midPos.z) + ")" + vbCrLf + _
           "Tangent:      (" & CStr(tanVec.x) + ", " + CStr(tanVec.y) + ", " + CStr(tanVec.z) + ")" + vbCrLf + _
           "Curve Norm:   (" & CStr(nX) + ", " + CStr(nY) + ", " + CStr(nZ) + ")" + vbCrLf + _
           "Adjusted Norm:(" & CStr(NormVec.x) + ", " + CStr(NormVec.y) + ", " + CStr(NormVec.z) + ")" + vbCrLf + _
           "Cross Prod:   (" & CStr(dirVec.x) + ", " + CStr(dirVec.y) + ", " + CStr(dirVec.z) + ")" + vbCrLf + _
           "Hint Position:(" & CStr(projectedPoint.x) + ", " + CStr(projectedPoint.y) + ", " + CStr(projectedPoint.z) + ")" + vbCrLf + _
           "On Right:      " & CStr(onRight)
    Set GetHintPoint = projectedPoint
    Exit Function
ErrorHandler:
    Dim oErrors As New IMSErrorLog.JServerErrors
    oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
End Function


Public Function GetHandrailSketchPathLength(oHandrail As ISPSHandrail) As Double
        Const METHOD = "GetHandrailSketchPathLength"
        On Error GoTo ErrorHandler
        Dim oHandRailSketchPath As Sketch3d
        Dim oHandRailComplexString As ComplexString3d
        Dim oHandRailPathCurve As IJCurve

        Set oHandRailSketchPath = oHandrail.SketchPath
        Set oHandRailComplexString = oHandRailSketchPath.GetComplexString
        Set oHandRailPathCurve = oHandRailComplexString
        GetHandrailSketchPathLength = oHandRailPathCurve.length

        Set oHandRailSketchPath = Nothing
        Set oHandRailComplexString = Nothing
        Set oHandRailPathCurve = Nothing
        Exit Function
ErrorHandler:
        Dim oErrors As New IMSErrorLog.JServerErrors
        oErrors.Add Err.Number, METHOD, Err.Description
        Err.Raise E_FAIL
    End Function

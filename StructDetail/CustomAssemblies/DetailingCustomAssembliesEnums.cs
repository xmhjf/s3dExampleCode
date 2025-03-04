//-------------------------------------------------------------------------------------------------------
//Copyright 1998 - 2016 Intergraph Corporation, Intergraph Process, Power & Marine. All rights reserved.
//
//File
//  DetailingCustomAssembliesEnums.cs
//
//Abstract
//	DetailingCustomAssembliesEnums is StructDetailing custom assemblies enums.
//-------------------------------------------------------------------------------------------------------
namespace Ingr.SP3D.Content.Structure.Services
{
    /// <summary>
    /// Enumerated values for CollarSideCol codelist table. This is a lookup definition of the "official" CodeList definition, which can be found in
    /// <![CDATA[<InstallDir>]]>\ShipCatalogData\BulkLoad\DataFiles\StructDetail\SM_AllShipCodeListsStructDetails.xls
    /// </summary>
    public enum CollarSide
    {
        Flip = 65536,
        NoFlip = 65537,
        Centered = 65538
    }

    /// <summary>
    /// Enumerated values for BooleanCol codelist table. This is a lookup definition of the "official" CodeList definition, which can be found in
    /// <![CDATA[<InstallDir>]]>\ShipCatalogData\BulkLoad\DataFiles\StructDetail\SM_AllShipCodeListsStructDetails.xls
    /// </summary>
    public enum Answer
    {
        Yes = 65536,
        No = 65537
    }

    /// <summary>
    /// Enumerated values for EndCutTypeCodeList codelist table. This is a lookup definition of the "official" CodeList definition, which can be found in
    /// <![CDATA[<InstallDir>]]>\ShipCatalogData\BulkLoad\DataFiles\StructDetail\SM_AllShipCodeListsStructDetails.xls
    /// </summary>
    public enum EndCutTypes
    {
        Welded = 65536,
        Snip = 65537,
        Cutback = 65538,
        Bracketed = 65539
    }

    /// <summary>
    /// Enumerated values for SplitEndCutTypes codelist table. This is a lookup definition of the "official" CodeList definition, which can be found in
    /// <![CDATA[<InstallDir>]]>\ShipCatalogData\BulkLoad\DataFiles\StructDetail\SM_AllShipCodeListsStructDetails.xls
    /// </summary>
    public enum SplitEndCutTypes
    {
        NoAngle = 65536,
        AngleWebSquareFlange = 65537,
        AngleWebBevelFlange = 65538,
        AngleWebAngleFlange = 65539,
        DistanceWebDistanceFlange = 65540,
        OffsetWebOffsetFlange = 65541,
        NoAngleOffset = 65542,
        AlongGlobalAxis = 65543
    }

    /// <summary>
    /// Enumerated values for SlotTypeCol codelist table. This is a lookup definition of the "official" CodeList definition, which can be found in
    /// <![CDATA[<InstallDir>]]>\ShipCatalogData\BulkLoad\DataFiles\StructDetail\SM_AllShipCodeListsStructDetails.xls
    /// </summary>
    public enum SlotType
    {
        Default = 65536,
        NonTight = 65537,
        Tight = 65538
    }

    /// <summary>
    /// Enumerated values for AssyMethodCol codelist table. This is a lookup definition of the "official" CodeList definition, which can be found in
    /// <![CDATA[<InstallDir>]]>\ShipCatalogData\BulkLoad\DataFiles\StructDetail\SM_AllShipCodeListsStructDetails.xls
    /// </summary>
    public enum AssemblyMethod
    {
        Drop = 65536,
        Slide = 65537,
        DropAtAngle = 65538,
        VerticalDrop = 65539,
        Default = 65540
    }

    /// <summary>
    /// Enumerated values for ChamferMeasurementCodeList codelist table. This is a lookup definition of the "official" CodeList definition, which can be found in
    /// <![CDATA[<InstallDir>]]>\ShipCatalogData\BulkLoad\DataFiles\StructDetail\SM_AllShipCodeListsStructDetails.xls
    /// </summary>
    public enum ChamferMeasurement
    {
        Slope = 65536,
        Angle = 65537
    }

    /// <summary>
    /// Enumerated values for BevelMethod codelist table. This is a lookup definition of the "official" CodeList definition, which can be found in
    /// <![CDATA[<InstallDir>]]>\ShipCatalogData\BulkLoad\DataFiles\StructDetail\SM_AllShipCodeListsStructDetails.xls
    /// </summary>
    public enum BevelMethod
    {
        Constant = 65536,
        Varying = 65537

    }

    /// <summary>
    /// Enumerated values for TeeWeldCategory codelist table. This is a lookup definition of the "official" CodeList definition, which can be found in
    /// <![CDATA[<InstallDir>]]>\ShipCatalogData\BulkLoad\DataFiles\StructDetail\SM_AllShipCodeListsStructDetails.xls
    /// </summary>
    public enum TeeWeldCategory
    {
        Unspecified = 65536,
        Normal = 65537,
        Deep = 65538,
        Full = 65539,
        Chain = 65540,
        Staggered = 65541,
        OneSidedBevel = 65542,
        TwoSidedBevel = 65543,
        Chill = 65544
    }

    /// <summary>
    /// Enumerated values for ClassSocietyCol codelist table. This is a lookup definition of the "official" CodeList definition, which can be found in
    /// <![CDATA[<InstallDir>]]>\ShipCatalogData\BulkLoad\DataFiles\StructDetail\SM_AllShipCodeListsStructDetails.xls
    /// </summary>
    public enum ClassSociety
    {
        Lloyds = 65536,
        ABS = 65537,
        DNV = 65538
    }

    /// <summary>
    /// Enum of values to define the WeldSymbol.
    /// </summary>
    public enum WeldSymbol
    {
        None = 0,
        Fillet = 1,
        StaggeredFillet = 2
    }

    /// <summary>
    /// Enumerated values for WeldFilletMeasure codelist table. This is a lookup definition of the "official" CodeList definition, which can be found in
    /// <![CDATA[<InstallDir>]]>\ShipCatalogData\BulkLoad\DataFiles\StructDetail\SM_AllShipCodeListsStructDetails.xls
    /// </summary>
    public enum WeldFilletMeasure
    {
        Leg = 65536,
        Throat = 65537
    }

    /// <summary>
    /// Enum of values to define the weld groove types. This is a lookup definition of the "official" CodeList definition, which can be found in
    /// <![CDATA[<InstallDir>]]>\CommonSchema\Middle\Schema\Codelists\CommonSchemaCodelists.xls
    /// </summary>
    public enum WeldGroove
    {
        None = 0,
        Square = 1,
        V = 2,
        Bevel = 3,
        U = 4,
        J = 5,
        FlareV = 6,
        FlareBevel = 7
    }

    /// <summary>
    /// Enum of values to define the bottom flange requires or not.
    /// </summary>
    public enum BottomFlange
    {
        No = 0,
        Yes = 1
    }

    /// <summary>
    /// Enumerated values for EndCutShapeCol codelist table. This is a lookup definition of the "official" CodeList definition, which can be found in
    /// <![CDATA[<InstallDir>]]>\ShipCatalogData\BulkLoad\DataFiles\StructDetail\SM_AllShipCodeListsStructDetails.xls
    /// </summary>
    public enum EndCutShape
    {
        Straight = 65536,
        Sniped = 65537,
        NotApplied = 65538
    }

    /// <summary>
    /// Enumerated values for BraceTypeCol codelist table. This is a lookup definition of the "official" CodeList definition, which can be found in
    /// <![CDATA[<InstallDir>]]>\ShipCatalogData\BulkLoad\DataFiles\StructDetail\SM_AllShipCodeListsStructDetails.xls
    /// </summary>
    public enum BraceType
    {
        NotApplicable = 65536,
        None = 65537,
        InsetMember = 65538,
        FlushMember = 65539
    }

    /// <summary>
    /// Enum of values to define the bottom flange.
    /// </summary>
    public enum SectionAlias
    {
        Web = 0,
        WebTopFlangeRight = 1,
        WebBuiltUpTopFlangeRight = 2,
        WebBottomFlangeRight = 3,
        WebBuiltUpBottomFlangeRight = 4,
        WebTopAndBottomRightFlanges = 5,
        WebTopFlange = 6,
        WebBottomFlange = 7,
        WebTopFlangeLeft = 8,
        WebBuiltUpTopFlangeLeft = 9,
        WebBottomFlangeLeft = 10,
        WebBuiltUpBottomFlangeLeft = 11,
        WebTopAndBottomLeftFlanges = 12,
        WebTopAndBottomFlanges = 13,
        FlangeLeftAndRightBottomWebs = 14,
        FlangeLeftAndRightTopWebs = 15,
        FlangeLeftAndRightWebs = 16,
        TwoWebsTwoFlanges = 17,
        TwoFlangesBetweenWebs = 18,
        TwoWebsBetweenFlanges = 19
    }

    /// <summary>
    /// Enumerated values for ShapeAtEdgeOverlapCol codelist table. This is a lookup definition of the "official" CodeList definition, which can be found in
    /// <![CDATA[<InstallDir>]]>\ShipCatalogData\BulkLoad\DataFiles\StructDetail\SM_AllShipCodeListsStructDetails.xls
    /// </summary>
    public enum ShapeAtEdgeOverlap
    {
        None = 65536,
        FaceToInsideCorner = 65537,
        FaceToEdge = 65538,
        FaceToOutsideCorner = 65539,
        FaceToOutside = 65540,
        InsideToEdge = 65541,
        InsideToOutsideCorner = 65542,
        InsideToOutside = 65543,
        InsideCornerToOutside = 65544,
        EdgeToOutside = 65546
    }

    /// <summary>
    /// Enumerated values for ShapeAtEdgeOutsideCol codelist table. This is a lookup definition of the "official" CodeList definition, which can be found in
    /// <![CDATA[<InstallDir>]]>\ShipCatalogData\BulkLoad\DataFiles\StructDetail\SM_AllShipCodeListsStructDetails.xls
    /// </summary>
    public enum ShapeAtEdgeOutside
    {
        EdgeToFlange = 65536,
        EdgeToOutside = 65537,
        CornerToFlange = 65538,
        CornerToOutside = 65539,
        OutsideToEdge = 65540,
        OutsideToFlange = 65541,
        OutsideToOutside = 65542,
        None = 65543
    }

    /// <summary>
    /// Enumerated values for ShapeAtEdgeCol codelist table. This is a lookup definition of the "official" CodeList definition, which can be found in
    /// <![CDATA[<InstallDir>]]>\ShipCatalogData\BulkLoad\DataFiles\StructDetail\SM_AllShipCodeListsStructDetails.xls
    /// </summary>
    public enum ShapeAtEdge
    {
        None = 65536,
        FaceToCorner = 65537,
        FaceToEdge = 65538,
        FaceToFlange = 65539,
        FaceToOutside = 65540,
        InsideToEdge = 65541,
        InsideToFlange = 65542,
        InsideToOutside = 65543,
        CornerToFlange = 65544,
        CornerToOutside = 65545,
        EdgeToFlange = 65546,
        EdgeToOutside = 65547
    }

    /// <summary>
    /// Enumerated values for ShapeAtFaceCol codelist table. This is a lookup definition of the "official" CodeList definition, which can be found in
    /// <![CDATA[<InstallDir>]]>\ShipCatalogData\BulkLoad\DataFiles\StructDetail\SM_AllShipCodeListsStructDetails.xls
    /// </summary>
    public enum ShapeAtFace
    {
        None = 65536,
        Inside = 65537,
        Cope = 65538,
        Snipe = 65539
    }

    /// <summary>
    /// Enum of values to define cutting behaviour value.
    /// </summary>
    public enum CuttingBehaviourValue
    {
        AvoidWeb = 1,
        ToFlangeInnerSurface = 2,
        ToBoundingEdge = 3
    }

    /// <summary>
    /// The location of a load point in the measurement symbol with respect 
    /// to a bounded edge. Or
    /// the location of a bounded edge relative to specified bounding edge.
    /// </summary>
    internal enum RelativePointPosition
    {
        Above = 1,
        Below = 2,
        Coincident = 3
    }

    /// <summary>
    /// Enum of values to define the ExtendOrOffsetCol. This is a lookup definition of the "official" CodeList definition, which can be found in
    /// <![CDATA[<InstallDir>]]>\ShipCatalogData\BulkLoad\DataFiles\StructDetail\SM_AllShipCodeListsStructDetails.xls
    /// </summary>
    public enum ExtentOrOffset
    {
        ExtendFarCorner = 65536,
        OffsetFarCorner = 65537,
        ExtendFarSide = 65538,
        OffsetFarSide = 65539,
        ExtendNearSide = 65540,
        OffsetNearSide = 65541,
        ExtendNearCorner = 65542,
        OffsetNearCorner = 65543
    }

    /// <summary>
    /// Enum of values to define the InsideCornerCol. This is a lookup definition of the "official" CodeList definition, which can be found in
    /// <![CDATA[<InstallDir>]]>\ShipCatalogData\BulkLoad\DataFiles\StructDetail\SM_AllShipCodeListsStructDetails.xls
    /// </summary>
    public enum InsideCorner
    {
        None = 65536,
        Snipe = 65537,
        Scallop = 65538,
        Fillet = 65539
    }

    /// <summary>
    /// Enumerated values for FlipCol codelist table. This is a lookup enum of the "official" CodeList definition, which can be found in
    /// <![CDATA[<InstallDir>]]>\ShipCatalogData\BulkLoad\DataFiles\StructDetail\SM_AllShipCodeListsStructDetails.xls
    /// </summary>
    public enum Flip
    {
        Yes = 65536,
        No = 65537
    }

    /// <summary>
    /// Enum of values to define the FirstWeldingSide. This is a duplicate definition of the "official" CodeList definition, which can be found in
    /// <![CDATA[<InstallDir>]]>\ShipCatalogData\BulkLoad\DataFiles\StructDetail\SM_AllShipCodeListsStructDetails.xls
    /// </summary>
    public enum FirstWeldingSide
    {
        MoldedSide = 65536,
        AntiMoldedSide = 65537
    }
}
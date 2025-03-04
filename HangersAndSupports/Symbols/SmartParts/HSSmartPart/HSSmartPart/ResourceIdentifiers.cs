//-----------------------------------------------------------------------------
// Copyright 1992 - 2013 Intergraph Corporation, Intergraph Process, Power & Marine. All rights reserved.
//
// File
//      ResourceIdentifiers.cs
// Author:     
//      Sridhar Bathina     
//
// Abstract:
//     Smart Parts resource identifiers
//-----------------------------------------------------------------------------

namespace Ingr.SP3D.Content.Support.Symbols
{
    /// <summary>
    /// Smart Parts Symbol resource identifiers.
    /// </summary>
    public static class SmartPartSymbolResourceIDs
    {
        /// <summary>
        /// Default Resource Name.
        /// </summary>   
        public const string DEFAULT_RESOURCE = "Ingr.SP3D.Content.Support.Symbols.Resources.HSSmartPart";
        /// <summary>
        /// Default Assembly Name.
        /// </summary>
        public const string DEFAULT_ASSEMBLY = "HSSmartPart";
        /// <summary>
        /// Invalid input arguments.
        /// </summary>
        public const int ErrInvalidArguments = 1;
        /// <summary>
        /// Section not found in Catalog
        /// </summary>
        public const int ErrSectionNotFound = 2;
        /// <summary>
        /// Invalid rodEndType code list value
        /// </summary>
        public const int ErrInvRodEndTypeeCodeListValue = 3;
        /// <summary>
        /// Error while constructing outputs
        /// </summary>
        public const int ErrConstructOutputs = 4;
        /// <summary>
        /// Error while creating the Ports
        /// </summary>
        public const int ErrCreatePorts = 5;
        /// <summary>
        /// Error while setting the Port Orientation
        /// </summary>
        public const int ErrPortOrientation = 6;
        /// <summary>
        /// Error while getting Material Type or Material Grade
        /// </summary>
        public const int ErrGettingMaterial = 7;
        /// <summary>
        /// Invalid code list value
        /// </summary>
        public const int ErrInvalidCodeListValue = 8;
        /// <summary>
        /// Invalid property value
        /// </summary>
        public const int ErrInvalidPropertyValue = 9;
        /// <summary>
        /// Error while loading data
        /// </summary>
        public const int ErrLoadNutData = 10;
        /// <summary>
        /// Error occure while adding nut additional Inputs
        /// </summary>
        public const int ErrNutAdditionalInputs = 11;
        /// <summary>
        /// Error occure while adding nut additional Outputs
        /// </summary>
        public const int ErrNutAdditionalOutputs = 12;
        /// <summary>
        /// Error occure while creating nut geometry
        /// </summary>
        public const int ErrAddNutMethod = 13;
        /// <summary>
        /// Invalid Nut Shape Type code list value
        /// </summary>
        public const int ErrShapeTypeCodeListValue = 14;
        /// <summary>
        /// Invalid Nut ShapeWidth1 value
        /// </summary>
        public const int ErrNutShapeWidth1Value = 15;
        /// <summary>
        /// Invalid Shape Length value
        /// </summary>
        public const int ErrNutShapeLengthValue = 16;
        /// <summary>
        /// Error occure while creating pin geometry
        /// </summary>
        public const int ErrAddPinMethod = 17;
        /// <summary>
        /// Invalid Pin Diameter value
        /// </summary>
        public const int ErrInvalidPinDiameterValue = 18;
        /// <summary>
        /// Invalid Cotter Diameter value
        /// </summary>
        public const int ErrPinCotterDiameterValue = 19;
        /// <summary>
        /// Invalid Pin Length value
        /// </summary>
        public const int ErrInvalidLengthValue = 20;
        /// <summary>
        /// Invalid Cotter Length value
        /// </summary>
        public const int ErrPinCotterLengthValue = 21;
        /// <summary>
        /// Error while loading Pin data
        /// </summary>
        public const int ErrLoadPinData = 22;
        /// <summary>
        /// Error while implementing weight CG.
        /// </summary>
        public const int ErrWeightCG = 23;
        /// <summary>
        /// Error while creating turnbuckle geometry
        /// </summary>
        public const int ErrAddAddTurnbuckle = 24;
        /// <summary>
        /// Turnbuckle has no Center Tube or Side Blocks
        /// </summary>
        public const int ErrNoCenterTubeOrSideBlocks = 25;
        /// <summary>
        /// Invalid Diameter1  value
        /// </summary>
        public const int ErrInvalidDiameter1Value = 26;
        /// <summary>
        /// Invalid Turnbuckle Opening1 value
        /// </summary>
        public const int ErrTurnbuckleOpening1Value = 27;
        /// <summary>
        /// Invalid Length2 value
        /// </summary>
        public const int ErrInvalidLength2Value = 28;
        /// <summary>
        /// Invalid Width2 value
        /// </summary>
        public const int ErrInvalidWidth2Value = 29;
        /// <summary>
        /// Invalid Thickness2 value
        /// </summary>
        public const int ErrInvalidThickness2Value = 30;
        /// <summary>
        /// Error while loading Turnbuckle data
        /// </summary>
        public const int ErrLoadTurnbuckleData = 31;
        /// <summary>
        /// Error while loading WBABolt data
        /// </summary>
        public const int ErrLoadWBABoltData = 32;
        /// <summary>
        /// Error while creating WBABolt geometry
        /// </summary>
        public const int ErrAddAddWBABolt = 33;
        /// <summary>
        /// Invalid Gap1 value
        /// </summary>
        public const int ErrInvalidGap1ValueNLTZero = 34;
        /// <summary>
        /// Invalid Thickness1 value
        /// </summary>
        public const int ErrInvalidThickness1NLTZero = 35;
        /// <summary>
        /// Invalid Height1 value
        /// </summary>
        public const int ErrInvalidHeight1NLTZero = 36;
        /// <summary>
        /// Invalid Width1 value
        /// </summary>
        public const int ErrInvalidWidth1NLTZero = 37;
        /// <summary>
        /// Invalid TLCornerType code list value
        /// </summary>
        public const int ErrInvTLCornerTypeCodeListValue = 38;
        /// <summary>
        /// Invalid TRCornerType code list value
        /// </summary>
        public const int ErrInvTRCornerTypeCodeListValue = 39;
        /// <summary>
        /// Error while loading WBAHole data
        /// </summary>
        public const int ErrLoadWBAHoleData = 40;
        /// <summary>
        /// Error while creating WBAHole geometry
        /// </summary>
        public const int ErrAddAddWBAHole = 41;
        /// <summary>
        /// Invalid WbaHole WBAHoleConfig code list value
        /// </summary>
        public const int ErrWBAHoleConfigCodeListValue = 42;
        /// <summary>
        /// Invalid WbaHole SimpShapeType code list value
        /// </summary>
        public const int ErrWBAHoleSimpShapeTypeCodeListValue = 43;
        /// <summary>
        /// Invalid Thickness4 value
        /// </summary>
        public const int ErrInvalidThickness4NLTZero = 44;
        /// <summary>
        /// Invalid Length2 value
        /// </summary>
        public const int ErrInvalidLength2NLTZero = 45;
        /// <summary>
        /// Invalid Width2 value
        /// </summary>
        public const int ErrInvalidWidth2NLTZero = 46;
        /// <summary>
        /// Invalid Thickness2 value
        /// </summary>
        public const int ErrInvalidThickness2NLTZero = 47;
        /// <summary>
        /// Invalid Offset1 value
        /// </summary>
        public const int ErrInvalidOffset1NLTZero = 48;
        /// <summary>
        /// Invalid Offset2 value
        /// </summary>
        public const int ErrInvalidOffset2NLTZero = 49;
        /// <summary>
        /// Invalid Thickness3 value
        /// </summary>
        public const int ErrInvalidThickness3NLTZero = 50;
        /// <summary>
        /// Invalid Width of Flat Spot value
        /// </summary>
        public const int ErrInvalidWidthFlatSpot = 51;
        /// <summary>
        /// Invalid pipe outside diameter value
        /// </summary>
        public const int ErrInvalidPipeOD = 52;
        /// <summary>
        /// Error while loading UBolt data
        /// </summary>
        public const int ErrLoadUBoltData = 53;
        /// <summary>
        /// Error while creating UBolt geometry
        /// </summary>
        public const int ErrAddUBoltMethod = 54;
        /// <summary>
        /// Invalid PipeToSteel value
        /// </summary>
        public const int ErrInvalidPipeToSteel = 55;
        /// <summary>
        /// Invalid UBolt SimpShapeType code list value
        /// </summary>
        public const int ErrUBoltUBoltOneSidedCodeListValue = 56;
        /// <summary>
        /// Invalid BLCornerType code list value
        /// </summary>
        public const int ErrInvBLCornerTypeCodeListValue = 57;
        /// <summary>
        /// Invalid BRCornerType code list value
        /// </summary>
        public const int ErrInvBRCornerTypeCodeListValue = 58;
        /// <summary>
        /// Error while loading Plate data
        /// </summary>
        public const int ErrLoadPlateData = 59;
        /// <summary>
        ///Rectangular Notch sizes invalid - ignoring " + corner[ID].name + " corner type.
        /// </summary>
        public const int ErrInvalidRectangularNotchsizes = 60;
        /// <summary>
        ///Angled Notch sizes invalid - ignoring " + corner[ID].name + " corner type. 
        /// </summary>
        public const int ErrInvalidAngledNotchsizes = 61;
        /// <summary>            
        ///Angled Notch X size with specified radius is larger than plate width - ignoring " + corner[ID].name + " corner type. 
        /// </summary>
        public const int ErrInvalidAngledNotchX = 62;
        /// <summary>
        ///Angled Notch Y size with specified radius is larger than plate length - ignoring " + corner[ID].name + " corner type.   
        /// </summary>
        public const int ErrInvalidAngledNotchY = 63;
        /// <summary>
        ///Round Notch is offset too far off plate, not visible - ignoring " + corner[ID].name + " corner type.         
        /// </summary>   
        public const int ErrInvalidRoundNotch = 64;
        /// <summary>
        ///Round Notch X or Y is offset too far off plate, not visible - ignoring " + corner[ID].name + " corner type.  
        /// </summary>
        public const int ErrInvalidRoundNotchXorY = 65;
        /// <summary>
        ///Round Notch has invalid radius - ignoring " + corner[ID].name + " corner type.
        /// </summary>
        public const int ErrInvalidRoundNotchRadius = 66;
        /// <summary>
        ///Round Notch radius with specified offsets is larger than plate width - ignoring " + corner[ID].name + " corner type. 
        /// </summary>
        public const int ErrInvalidRoundCornerRadius = 67;
        /// <summary>
        ///Rounded Corner has invalid radius - ignoring " + corner[ID].name + " corner type.
        /// </summary>
        public const int ErrInvalidRoundTopCorner = 68;
        /// <summary>
        ///Rounded Top Corner has positive X offset, not allowed - using negative of value. 
        /// </summary>
        public const int ErrInvalidRoundTopCornerX = 69;
        /// <summary>
        ///Rounded Top Corner has positive Y offset, not allowed - using negative of value.  
        /// </summary>
        public const int ErrInvalidRoundTopCornerY = 70;
        /// <summary>
        ///Rounded Corner with offsets would be off the plate - ignoring " + corner[ID].name + " corner type.
        /// </summary>
        public const int ErrInvalidRoundCornerOffset = 71;
        /// <summary>
        ///Round Corner with offsets larger than plate width - ignoring " + corner[ID].name + " corner type. 
        /// </summary>
        public const int ErrInvalidRCOffsettGTPlateWidth = 72;
        /// <summary>
        ///Curved End Y offset is greater than specified Curved End radius - ignoring Curved End.       
        /// </summary>
        public const int ErrInvalidCurveEndY = 73;
        /// <summary>
        ///Curved End doesn't fit on remaining top edge - ignoring Curved End.    
        /// </summary>        
        public const int ErrInvalidCurveEnd = 74;
        /// <summary>
        /// Error while creating Plate geometry
        /// </summary>
        public const int ErrAddPlateMethod = 75;
        /// <summary>
        /// Error while loading Clevis data
        /// </summary>
        public const int ErrLoadClevisData = 76;
        /// <summary>
        /// Error while creating Clevis geometry
        /// </summary>
        public const int ErrAddClevisMethod = 77;
        /// <summary>
        /// Error occure while adding Clevis additional Inputs
        /// </summary>
        public const int ErrClevisAdditionalInputs = 78;
        /// <summary>
        /// Error occure while adding Clevis additional Outputs
        /// </summary>
        public const int ErrClevisAdditionalOutputs = 79;
        /// <summary>
        /// Error occure while adding Plate additional Inputs
        /// </summary>
        public const int ErrPlateAdditionalInputs = 80;
        /// <summary>
        /// Error occure while adding Plate additional Outputs
        /// </summary>
        public const int ErrPlateAdditionalOutputs = 81;
        /// <summary>
        /// Error occure while adding WBABolt additional Inputs
        /// </summary>
        public const int ErrWBABoltAdditionalInputs = 82;
        /// <summary>
        /// Error occure while adding WBABolt additional Outputs
        /// </summary>
        public const int ErrWBABoltAdditionalOutputs = 83;
        /// <summary>
        /// Error occure while adding WBAHole additional Inputs
        /// </summary>
        public const int ErrWBAHoleAdditionalInputs = 84;
        /// <summary>
        /// Error occure while adding WBAHole additional Outputs
        /// </summary>
        public const int ErrWBAHoleAdditionalOutputs = 85;
        /// <summary>
        /// Error occure while adding UBolt additional Inputs
        /// </summary>
        public const int ErrUBoltAdditionalInputs = 86;
        /// <summary>
        /// Error occure while adding UBolt additional Outputs
        /// </summary>
        public const int ErrUBoltAdditionalOutputs = 87;
        /// <summary>
        /// Error occure while adding Turnbuckle additional Inputs
        /// </summary>
        public const int ErrTurnbuckleAdditionalInputs = 88;
        /// <summary>
        /// Error occure while adding Turnbuckle additional Outputs
        /// </summary>
        public const int ErrTurnbuckleAdditionalOutputs = 89;
        /// <summary>
        /// Error occure while adding Pin additional Inputs
        /// </summary>
        public const int ErrPinAdditionalInputs = 90;
        /// <summary>
        /// Error occure while adding Pin additional Outputs
        /// </summary>
        public const int ErrPinAdditionalOutputs = 91;
        /// <summary>
        /// Invalid Width3 value
        /// </summary>
        public const int ErrInvalidWidth3GTZero = 92;
        /// <summary>
        /// Invalid Thickness3 value
        /// </summary>
        public const int ErrInvalidThickness3GTZero = 93;
        /// <summary>
        /// Invalid Pin1Diameter value
        /// </summary>
        public const int ErrInvalidPin1DiameterNZero = 94;
        /// <summary>
        /// Invalid Pin1Length value
        /// </summary>
        public const int ErrInvalidPin1LengthNZero = 95;
        /// <summary>
        /// Invalid Gap2 value
        /// </summary>
        public const int ErrInvalidGap2NLTZero = 96;
        /// <summary>
        /// Clevis has neither Round Ends or Side Blocks AddCreateBoxClevis
        /// </summary>
        public const int ErrInvalidRoundEndsSideBlocks = 97;
        /// <summary>
        /// Error while loading StrutB data
        /// </summary>
        public const int ErrLoadStrutBData = 98;
        /// <summary>
        /// Error while creating StrutB geometry
        /// </summary>
        public const int ErrAddStrutBMethod = 99;
        /// <summary>
        /// Error occure while adding StrutB additional Inputs
        /// </summary>
        public const int ErrStrutBAdditionalInputs = 100;
        /// <summary>
        /// Error occure while adding StrutB additional Outputs
        /// </summary>
        public const int ErrStrutBAdditionalOutputs = 101;
        /// <summary>
        /// Invalid Diameter1 value
        /// </summary>
        public const int ErrInvalidDiameter1NZero = 102;
        /// <summary>
        /// Invalid Thickness1 value
        /// </summary>
        public const int ErrInvalidThickness1NZero = 103;
        /// <summary>
        /// Invalid Width1 value
        /// </summary>
        public const int ErrInvalidWidth1NZero = 104;
        /// <summary>
        /// Invalid Width2 value
        /// </summary>
        public const int ErrInvalidWidth2NZero = 105;
        /// <summary>
        /// Invalid Thickness2 value
        /// </summary>
        public const int ErrInvalidThickness2NZero = 106;
        /// <summary>
        /// Invalid Pin1Diameter value
        /// </summary>
        public const int ErrInvalidPinDiameterNLTZero = 107;
        /// <summary>
        /// Invalid Thickness1 value
        /// </summary>
        public const int ErrInvalidThickness1GTZero = 108;
        /// <summary>
        /// Invalid Gap2 value
        /// </summary>
        public const int ErrInvalidGap2ValueNLTZero = 109;
        /// <summary>
        /// Invalid Angle1 value
        /// </summary>
        public const int ErrInvalidAngle1ValueNLTZero = 110;
        /// <summary>
        /// Invalid Width5 value
        /// </summary>
        public const int ErrInvalidWidth5ValueGTZero = 111;
        /// <summary>
        /// Invalid Width1 and Width2 values
        /// </summary>
        public const int ErrInvalidWidth1andWidth2ValuesGTZero = 112;
        /// <summary>
        /// Invalid Offset1 value
        /// </summary>
        public const int ErrInvalidOffset1GTZero = 113;
        /// <summary>
        /// Invalid Width3 value
        /// </summary>
        public const int ErrInvalidWidth3NGTDiameter = 114;
        /// <summary>
        /// Invalid Width3 value
        /// </summary>
        public const int ErrInvalidWidth3GTHeight = 115;
        /// <summary>
        /// Invalid Offset2 value
        /// </summary>
        public const int ErrInvalidOffset2GTZero = 116;
        /// <summary>
        /// Invalid Height2 value
        /// </summary>
        public const int ErrInvalidHeight2GTZero = 117;
        /// <summary>
        /// Error while creating PipeClamp geometry
        /// </summary>
        public const int ErrAddPipeClampMethod = 118;
        /// <summary>
        /// Error in AddClamp
        /// </summary>
        public const int ErrAddClampMethod = 119;
        // <summary>
        /// Error in AddClampWithwings
        /// </summary>
        public const int ErrAddClampWithwingsMethod = 120;
        // <summary>
        /// Error in AddClampShape
        /// </summary>
        public const int ErrAddClampShapeMethod = 121;
        // <summary>
        /// Error in AddGussets
        /// </summary>
        public const int ErrAddGussetsMethod = 122;
        // <summary>
        /// Error in AddBoltRows
        /// </summary>
        public const int ErrAddBoltRowsMethod = 123;
        // <summary>
        /// Error in AddBoltByRow
        /// </summary>
        public const int ErrAddBoltByRowMethod = 124;
        // <summary>
        /// Error in MutiPosition
        /// </summary>
        public const int ErrMutiPositionMethod = 125;
        // <summary>
        /// Invalid Qty2 Value
        /// </summary>
        public const int ErrInvalidQty2 = 126;
        // <summary>
        /// Invalid Qty3 Value
        /// </summary>
        public const int ErrInvalidQty3 = 127;
        // <summary>
        /// Invalid Width for Qty Value
        /// </summary>
        public const int ErrInvalidWidthQty = 128;
        // <summary>
        /// Invalid Distance Value
        /// </summary>
        public const int ErrInvalidDistance = 129;
        // <summary>
        /// Error in LoadBoltRowData
        /// </summary>
        public const int ErrLoadBoltRowData = 130;
        // <summary>
        /// Error in AddBoltRowInputs
        /// </summary>
        public const int ErrAddBoltRowInputs = 131;
        // <summary>
        /// Error in CreateSideConnection
        /// </summary>
        public const int ErrCreateSideConnection = 132;
        // <summary>
        /// Error in AddPipeClampInputs
        /// </summary>
        public const int ErrAddPipeClampInputs = 133;
        // <summary>
        /// Error in AddPipeClampOutputs
        /// </summary>
        public const int ErrAddPipeClampOutputs = 134;
        // <summary>
        /// Error in LoadPipeClampData
        /// </summary>
        public const int ErrLoadPipeClampData = 135;
        // <summary>
        /// Invalid Width1 Value
        /// </summary>
        public const int ErrInvalidWidth1GTZero = 136;
        // <summary>
        /// Invalid Width2 Value
        /// </summary>
        public const int ErrInvalidWidth2GTZero = 137;
        // <summary>
        /// Invalid Width3 Value
        /// </summary>
        public const int ErrInvalidClampWidth3GTZero = 138;
        // <summary>
        /// Invalid Width4 Value
        /// </summary>
        public const int ErrInvalidWidth4GTZero = 139;
        /// <summary>
        /// Invalid StiffenerHeight value
        /// </summary>
        public const int ErrInvalidStiffenerHeightValue = 140;
        // <summary>
        /// Invalid Thickness2 Value
        /// </summary>
        public const int ErrInvalidThickness2GTZero = 141;
        /// <summary>
        /// Invalid Pin1Diameter value
        /// </summary>
        public const int ErrInvalidPin1DiameterNLTZero = 142;
        /// <summary>
        /// Invalid StiffenerLength value
        /// </summary>
        public const int ErrInvalidStiffenerLengthValue = 143;
        // <summary>
        /// Invalid Length1 Value
        /// </summary>
        public const int ErrInvalidLength1GTZero = 144;
        // <summary>
        /// Invalid Length2 Value
        /// </summary>
        public const int ErrInvalidLength2GTZero = 145;
        // <summary>
        /// Invalid Length3 Value
        /// </summary>
        public const int ErrInvalidLength3GTZero = 146;
        // <summary>
        /// Invalid Length4 Value
        /// </summary>
        public const int ErrInvalidLength4GTZero = 147;
        /// <summary>
        /// Error in AddGussetsByRow Method
        /// </summary>
        public const int ErrAddGussetsByRow = 148;

        // <summary>
        /// Error in LoadBlockClampData
        /// </summary>
        public const int ErrLoadBlockClampData = 149;
        // <summary>
        /// Error in AddBlockClampInputs
        /// </summary>
        public const int ErrAddBlockClampInputs = 150;
        // <summary>
        /// Error in AddBlockClampOutputs
        /// </summary>
        public const int ErrAddBlockClampOutputs = 151;
        // <summary>
        /// Error in LoadPlateDataByQuery
        /// </summary>
        public const int ErrLoadPlateDataByQuery = 152;
        // <summary>
        /// Error in AddBlockClamp
        /// </summary>
        public const int ErrAddBlockClamp = 153;
        // <summary>
        /// Error in AddBlockClampShape
        /// </summary>
        public const int ErrAddBlockClampShape = 154;
        /// <summary>
        /// Error in AddGuideInputs
        /// </summary>
        public const int ErrGuideAdditionalInputs = 155;
        /// <summary>
        ///  Error in AddGuideOutputs
        /// </summary>
        public const int ErrGuideAdditionalOutputs = 156;
        /// <summary>
        /// Invalid Width1 Value
        /// </summary>
        public const int ErrInvalidWidth1 = 157;
        /// <summary>
        /// Invalid Width3 Value
        /// </summary>
        public const int ErrInvalidWidth3 = 158;
        ///<summary>
        /// Invalid Offset Value
        /// </summary>
        public const int ErrInvalidOffset = 159;
        /// <summary>
        /// Invalid SolidBaseVerPl Value
        /// </summary>
        public const int ErrInvalidSolidBaseVerPl = 160;
        /// <summary>
        /// Invalid SolidVerHorPl Value
        /// </summary>
        public const int ErrInvalidSolidVerHorPl = 161;
        /// <summary>
        /// Invalid Thickness2 Value
        /// </summary>
        public const int ErrInvalidThickness2 = 162;
        /// <summary>
        /// Invalid Length2 Value
        /// </summary>
        public const int ErrInvalidLength2 = 163;
        /// <summary>
        /// Error in AddGuideMethod
        /// </summary>
        public const int ErrAddGuideMethod = 164;
        /// <summary>
        /// Error in AddRodInputs Method
        /// </summary>
        public const int ErrAddRodInputs = 165;
        /// <summary>
        /// Error in LoadGuideData
        /// </summary>
        public const int ErrLoadGuideData = 166;
        /// <summary>
        ///  Invalid SecConfig code list value
        /// </summary>
        public const int ErrInvalidSecConfigCodeListValue = 167;
        /// <summary>
        ///  Invalid VerPlSecStand value
        /// </summary>
        public const int ErrInvalidVerPlSecStand = 168;
        /// <summary>
        ///  Invalid VerPlSecSize value
        /// </summary>
        public const int ErrInvalidVerPlSecSize = 169;
        /// <summary>
        ///  Invalid VerPlSecType value
        /// </summary>
        public const int ErrInvalidVerPlSecType = 170;
        /// <summary>
        ///  Invalid HorPlSecStand value
        /// </summary>
        public const int ErrInvalidHorPlSecStand = 171;
        /// <summary>
        ///  Invalid HorPlSecSize value
        /// </summary>
        public const int ErrInvalidHorPlSecSize = 172;
        /// <summary>
        ///  Invalid HorPlSecType value
        /// </summary>
        public const int ErrInvalidHorPlSecType = 173;
        /// <summary>
        /// Invalid Connection code list value
        /// </summary>
        public const int ErrInvalidConnection = 174;
        /// <summary>
        /// Invalid Mirrored code list value
        /// </summary>
        public const int ErrInvalidMirrored = 175;
        /// <summary>
        ///  Error in AddGussetsWithChamferByRow 
        /// </summary>
        public const int ErrAddGussetsWithChamferByRowMethod = 176;
        /// <summary>
        ///  Error in AddSlidePlateInputs 
        /// </summary>
        public const int ErrSlidePlateAdditionalInputs = 177;
        /// <summary>
        ///  Error in AddSlidePlateOutputs
        /// </summary>
        public const int ErrSlidePlateAdditionalOutputs = 178;
        /// <summary>
        ///  Error in LoadSlidePlateData 
        /// </summary>
        public const int ErrLoadSlidePlateData = 179;
        /// <summary>
        ///  Error in AddSlidePlate 
        /// </summary>
        public const int ErrAddSlidePlateMethod = 180;
        /// <summary>
        /// Error in AddSwivelRingInputs
        /// </summary>
        public const int ErrSwivelRingAdditionalInputs = 181;
        /// <summary>
        ///  Error in AddSwivelRingOutputs
        /// </summary>
        public const int ErrSwivelRingAdditionalOutputs = 182;
        /// <summary>
        ///  Error in LoadSwivelRingData 
        /// </summary>
        public const int ErrLoadSwivelRingData = 183;
        /// <summary>
        ///  Error in AddSwivelRing
        /// </summary>
        public const int ErrAddSwivelRingMethod = 184;
        /// <summary>
        ///  Error in LoadGuideDataByQuery
        /// </summary>
        public const int ErrLoadGuideDataByQuery = 185;
        /// <summary>
        ///  Invalid RodDiameter value
        /// </summary>
        public const int ErrInvalidRodDiameter = 186;
        /// <summary>
        ///  Invalid Width1AndWidht2 value
        /// </summary>
        public const int ErrInvalidWidth1AndWidht2 = 187;
        /// <summary>
        ///  Invalid WrapOnJHangers value
        /// </summary>
        public const int ErrInvalidWrapOnJHangers = 188;
        /// <summary>
        ///  Error in AddSwivelLiner
        /// </summary>
        public const int ErrAddSwivelLinerMethod = 189;
        /// <summary>
        /// Error with pin diameter.
        /// </summary>
        public const int ErrEyeNutPinDiameter = 190;
        /// <summary>
        /// Error AddEyeNut Method.
        /// </summary>
        public const int ErrAddEyeNutMethod = 191;
        /// <summary>
        /// Invalid InnerWidth2 Value.
        /// </summary>
        public const int ErrInnerWidth2Argument = 192;
        /// <summary>
        /// Invalid InnerLength2 Value.
        /// </summary>
        public const int ErrInnerLength2Argument = 193;
        /// <summary>
        /// Error in LoadEyeNutData
        /// </summary>
        public const int ErrLoadEyeNutData = 194;
        /// <summary>
        /// Error in EyeNutAdditionalInputs
        /// </summary>
        public const int ErrEyeNutAdditionalInputs = 195;
        /// <summary>
        /// Error in AddEyeNutOutputs
        /// </summary>
        public const int ErrAddEyeNutOutputs = 196;
        /// <summary>
        /// Error in YokeClampAdditionalInputs
        /// </summary>
        public const int ErrYokeClampAdditionalInputs = 197;
        /// <summary>
        /// Error in AddYokeClampOutputs
        /// </summary>
        public const int ErrAddYokeClampOutputs = 198;
        /// <summary>
        /// Enter Positive value for Diameter1
        /// </summary>
        public const int ErrDiameterGTZero = 199;
        /// <summary>
        /// Rod Takeout Should greater than half of the Diamter1 plus half of the Pin1 Diameter
        /// </summary>
        public const int ErrRodTakeOut = 200;
        ///<summary>
        /// Invalid Length1 Value
        /// </summary>
        public const int ErrInvalidLength1NZero = 201;
        /// <summary>
        /// Invalid Length2 Value
        /// </summary>
        public const int ErrInvalidLength2NZero = 202;
        /// <summary>
        /// Invalid Length2 Value
        /// </summary>
        public const int ErrMissingInputs = 203;
        /// <summary>
        /// Invalid PipeOD Value
        /// </summary>
        public const int ErrInvalidPipeODGTZero = 204;
        /// <summary>
        /// Invalid Thickness1 and Radius1 Value
        /// </summary>
        public const int ErrInvalidThickness1AndRadius1NGTZero = 205;
        /// <summary>
        /// Invalid Angle1 Value
        /// </summary>
        public const int ErrInvalidAngle1GTZero = 206;
        /// <summary>
        /// Error in LoadPipeClampDataByQuery
        /// </summary>
        public const int ErrLoadPipeClampDataByQuery = 207;
        /// <summary>
        /// Error in AddProtectSaddle
        /// </summary>
        public const int ErrAddProtectSaddle = 208;
        /// <summary>
        /// Error in AddProtectSaddleInputs
        /// </summary>
        public const int ErrAddProtectSaddleInputs = 209;
        /// <summary>
        /// Error in AddProtectSaddleOutputs
        /// </summary>
        public const int ErrAddProtectSaddleOutputs = 210;
        /// <summary>
        /// Error in LoadProtectSaddle
        /// </summary>
        public const int ErrLoadProtectSaddle = 211;
        /// <summary>
        /// Invalid Height1 value
        /// </summary>
        public const int ErrInvalidHeight1GTZero = 212;
        /// <summary>
        /// Invalid Width1 and Width2 GTZero And LTPipeOD values
        /// </summary>
        public const int ErrInvalidWidth1andWidth2GTZeroAndLTPipeOD = 213;
        ///<summary>
        /// Invalid Angle1 & Angle2 values
        /// </summary>
        public const int ErrInvalidAngle1andAngle2GTZero = 214;
        /// <summary>
        /// Invalid Angle3 & Angle4 values
        /// </summary>
        public const int ErrInvalidAngle3andAngle4GTZero = 215;
        /// <summary>
        /// Invalid Angle1 & Angle2 values
        /// </summary>
        public const int ErrInvalidAngle1andAngle2NGT360 = 216;
        /// <summary>
        /// Invalid Angle1 & Angle2 & Angle3 & Angle4 values
        /// </summary>
        public const int ErrInvalidAngle1andAngle2andAngle3andAngle4NGT360 = 217;
        /// <summary>
        /// Invalid Angle1 & Angle2 & Angle3 & Angle4 values
        /// </summary>
        public const int ErrInvalidAngle1andAngle2andAngle3andAngle4E360 = 218;
        /// <summary>
        /// Invalid Angle5 and Height1 values
        /// </summary>
        public const int ErrInvalidAngle5andHeight1GTZero = 219;
        /// <summary>
        /// Error in AddSheild
        /// </summary>
        public const int ErrAddSheild = 220;
        /// <summary>
        /// Error in AddSheildShape
        /// </summary>
        public const int ErrAddSheildShape = 221;
        /// <summary>
        /// Error in AddShieldInputs
        /// </summary>
        public const int ErrAddShieldInputs = 222;
        /// <summary>
        /// Error in AddShieldOutputs
        /// </summary>
        public const int ErrAddShieldOutputs = 223;
        /// <summary>
        /// Error in LoadShieldData
        /// </summary>
        public const int ErrLoadShieldData = 224;
        /// <summary>
        /// Error in LoadLineData
        /// </summary>
        public const int ErrLoadLineData = 225;
        // <summary>
        /// Error in ShieldMultiPosition
        /// </summary>
        public const int ErrShieldMultiPosition = 226;

        /// <summary>
        /// Invalid RodDiameter value
        /// </summary>
        public const int ErrInvalidRodDiameterGTZero = 227;
        /// <summary>
        /// Invalid Length Of Rod value
        /// </summary>
        public const int ErrInvalidLengthOfRod = 228;
        /// <summary>
        /// Error in AddRod
        /// </summary>
        public const int ErrAddRod = 229;
        /// <summary>
        /// Error in AddRod1Inputs
        /// </summary>
        public const int ErrAddRod1Inputs = 230;
        /// <summary>
        /// Error in AddRod1Outputs
        /// </summary>
        public const int ErrAddRod1Outputs = 231;
        /// <summary>
        /// Error in LoadRod1dData
        /// </summary>
        public const int ErrLoadRod1Data = 232;

        /// <summary>
        /// Invalid Diameter2 value
        /// </summary>
        public const int ErrInvalidDiameter2GTZero = 233;
        /// <summary>
        /// Invalid Pin2Diameter value
        /// </summary>
        public const int ErrInvalidPin2DiameterGTZero = 234;
        /// <summary>
        /// Invalid Pin2Length value
        /// </summary>
        public const int ErrInvalidPin2LengthNZero = 235;
        /// <summary>
        /// Invalid Height1 lessthan RodTakeout value
        /// </summary>
        public const int ErrInvalidHeight1LTRodTakeout = 236;
        /// <summary>
        /// Invalid RodTakeout value
        /// </summary>
        public const int ErrInvalidRodTakeoutGTDiameter = 237;
        /// <summary>
        /// Invalid Length2 lessthan Pin1Length value
        /// </summary>
        public const int ErrInvalidLength2LTPin1Length = 238;
        /// <summary>
        /// Invalid Length2 equalsto Diameter1 value
        /// </summary>
        public const int ErrInvalidLength2EDiameter1 = 239;
        /// <summary>
        /// Invalid Height8 lessthan Height5 value
        /// </summary>
        public const int ErrInvalidHeight8LTHeight5 = 240;
        /// <summary>
        /// Invalid Pin2Length lessthan or equalsto Length2 value
        /// </summary>
        public const int ErrInvalidPin2LengthLTELength2 = 241;
        /// <summary>
        /// Invalid Pin2Length lessthan or equalsto Length2 & Diameter1 value
        /// </summary>
        public const int ErrInvalidPin2LengthLTELength2AndDiameter1 = 242;
        /// <summary>
        /// Error in AddClevisHanger
        /// </summary>
        public const int ErrAddClevisHanger = 243;
        /// <summary>
        /// Error in AddClevisHangerInputs
        /// </summary>
        public const int ErrAddClevisHangerInputs = 244;
        /// <summary>
        /// Error in AddClevisHangerOutputs
        /// </summary>
        public const int ErrAddClevisHangerOutputs = 245;
        /// <summary>
        /// Error in LoadClevisHangerData
        /// </summary>
        public const int ErrLoadClevisHangerData = 246;
        /// <summary>
        /// Enter Positive value for Pin1Diameter
        /// </summary>
        public const int ErrPin1DiameterGTZero = 247;
        /// <summary>
        /// Error AddYokeClamp Method.
        /// </summary>
        public const int ErrAddYokeClampMethod = 248;
        /// <summary>
        /// Width3 Should not be equal to Zero.
        /// </summary>
        public const int ErrInvalidWidth3NETZero = 249;
        /// <summary>
        /// RodTakeOut Should not be less then or equal to zero.
        /// </summary>
        public const int ErrInvalidRodTakeOut = 250;
        /// <summary>
        /// Heigth6 Should not be less then zero.
        /// </summary>
        public const int ErrInvalidHeigth6GTZero = 251;
        /// <summary>
        /// Heigth5 Should not be greater then zero.
        /// </summary>
        public const int ErrInvalidHeigth5LTZero = 252;
        /// <summary>
        /// Pin1Diameter Should not be less then or equal to zero.
        /// </summary>
        public const int ErrInvalidPin1DiameterGTZero = 253;
        /// <summary>
        /// Heigth2 Should not be less then or equal to zero.
        /// </summary>
        public const int ErrInvalidHeigth2GTZero = 254;
        /// <summary>
        /// Error in LoadYokeClampData
        /// </summary>
        public const int ErrLoadYokeClampData = 255;
        /// <summary>
        /// Length5 Should not be greater then zero.
        /// </summary>
        public const int ErrInvalidLength5LTZero = 256;
        /// <summary>
        /// Error while loading ElbowLug data
        /// </summary>
        public const int ErrLoadElbowLugData = 257;
        /// <summary>
        /// Error in AddElbowLugInputs
        /// </summary>
        public const int ErrAddElbowLugInputs = 258;
        /// <summary>
        /// Error in AddElbowLugInputs
        /// </summary>
        public const int ErrAddElbowLugOutputs = 259;
        /// <summary>
        /// Error in AddElbowLug Method
        /// </summary>
        public const int ErrAddElbowLug = 260;
        /// <summary>
        /// Invalid Angle1 value
        /// </summary>
        public const int ErrAngle1Range = 261;
        /// <summary>
        /// Invalid Angle2 value
        /// </summary>
        public const int ErrAngle2Range = 262;
        /// <summary>
        /// Invalid Angle3 value
        /// </summary>
        public const int ErrAngle3Range = 263;
        /// <summary>
        /// Invalid Angle4 value
        /// </summary>
        public const int ErrAngle4Range = 264;
        /// <summary>
        /// Invalid Angle3 value
        /// </summary>
        public const int ErrInvalidAngle3 = 265;
        /// <summary>
        /// Invalid Gap1 value
        /// </summary>
        public const int ErrInvalidGap1Value = 266;
        /// <summary>
        /// Invalid Height value
        /// </summary>
        public const int ErrInvalidHeightValue = 267;
        /// <summary>
        /// Error while loading Strap data
        /// </summary>
        public const int ErrLoadStrapData = 268;
        /// <summary>
        /// Error in AddStrapInputs
        /// </summary>
        public const int ErrAddStrapInputs = 269;
        /// <summary>
        /// Error in AddStrapInputs
        /// </summary>
        public const int ErrAddStrapOutputs = 270;
        /// <summary>
        /// Error in AddStrap Method
        /// </summary>
        public const int ErrAddStrap = 271;
        /// <summary>
        /// Invalid StrapFlatSpot value
        /// </summary>
        public const int ErrStrapFlatSpot = 272;
        /// <summary>
        /// Invalid StrapWidthWings value
        /// </summary>
        public const int ErrStrapWidthWings = 273;
        /// <summary>
        /// Invalid pipeOD value
        /// </summary>
        public const int ErrPipeODNGTStrapWidthInside = 274;
        /// <summary>
        /// Invalid Length3 Value
        /// </summary>
        public const int ErrInvalidLengthNZero = 275;

        /// <summary>
        /// Error in AddBeamClampInputs Method
        /// </summary>
        public const int ErrAddBeamClampInputs = 276;
        /// <summary>
        /// Error in AddAddBoltInputs Method
        /// </summary>
        public const int ErrAddBoltInputs = 277;
        /// <summary>
        /// Error in AddBeamClampOutputs Method
        /// </summary>
        public const int ErrAddBeamClampOutputs = 278;
        /// <summary>
        /// Error in LoadBoltData Method
        /// </summary>
        public const int ErrLoadBoltData = 279;
        /// <summary>
        /// Error in LoadBeamClipDataByQuery Method
        /// </summary>
        public const int ErrLoadBeamClipDataByQuery = 280;
        /// <summary>
        /// Error in LoadNutDataByQuery Method
        /// </summary>
        public const int ErrLoadNutDataByQuery = 281;
        /// <summary>
        /// Error in LoadBeamClampData Method
        /// </summary>
        public const int ErrLoadBeamClampData = 282;
        /// <summary>
        /// Error in AddBoltWithHead Method
        /// </summary>
        public const int ErrAddBoltWithHead = 283;
        /// <summary>
        /// Error in AddBeamClamp Method
        /// </summary>
        public const int ErrAddBeamClamp = 284;
        /// <summary>
        /// Error in AddBeamClip Method
        /// </summary>
        public const int ErrAddBeamClip = 285;
        /// <summary>
        /// Error in AddMalleableBeamClampInputs Method
        /// </summary>
        public const int ErrAddMalleableBeamClampInputs = 286;
        /// <summary>
        /// Error in AddMalleableBeamClampOutputs Method
        /// </summary>
        public const int ErrAddMalleableBeamClampOutputs = 287;
        /// <summary>
        /// Error in LoadMaleableBCData Method
        /// </summary>
        public const int ErrLoadMaleableBCData = 288;
        /// <summary>
        /// Error in LoadEyeNutDataByQuery Method
        /// </summary>
        public const int ErrLoadEyeNutDataByQuery = 289;
        /// <summary>
        /// Error in AddMalleablBeamClamp Method
        /// </summary>
        public const int ErrAddMalleablBeamClamp = 290;
        /// <summary>
        /// Width1 and Width 2 value are not enough to create the Overlap with Other Beam clip
        /// </summary>
        public const int ErrOverlapBeamClip = 291;
        /// <summary>
        /// Length1 is less than Structure Width with Gaps
        /// </summary>
        public const int ErrStructureWidthGaps = 292;
        /// <summary>
        /// Length1 is less than Structure Width with Gaps and Clip Widths. Resetting the value
        /// </summary>
        public const int ErrStructureWidthGapsClips = 293;
        /// <summary>
        /// UBolt Length is less than Structure Width with Gaps. Resetting the value
        /// </summary>
        public const int ErrUBoltStructureWidthGaps = 294;
        /// <summary>
        /// UBolt Length is less than Left Vertical Offset
        /// </summary>
        public const int ErrUBoltLengthLeftOffSet = 295;
        /// <summary>
        /// UBolt Length is less than Right Vertical Offset
        /// </summary>
        public const int ErrUBoltLengthRightOffSet = 296;
        /// <summary>
        /// Depth should not be less than or equal to zero
        /// </summary>
        public const int ErrMalleableDepth = 297;
        /// <summary>
        /// TopWidth should not be less than or equal to zero
        /// </summary>
        public const int ErrMalleableTopWidth = 298;
        /// <summary>
        /// Thickness should not be less than or equal to zero
        /// </summary>
        public const int ErrMalleableThickness = 299;
        /// <summary>
        /// Clamp Thickness should not be less than or equal to zero
        /// </summary>
        public const int ErrMalleableClamp = 300;
        /// <summary>
        /// Structure Width should not be less than or equal to zero
        /// </summary>
        public const int ErrMalleableStructureWidth = 301;
        /// <summary>
        /// Pin1Length should be more than FlangeWidth + 2* Thickness. Resetting value to FlangeWidth + 4* Thickness
        /// </summary>
        public const int ErrMalleablePin1Length = 302;
        /// <summary>
        /// Pin2 Diameter should not be more than the Thickness of the Malleable Beam Clamp
        /// </summary>
        public const int ErrMalleablePin2Diameter = 303;
        /// <summary>
        /// UBolt Length is less than Right Vertical Offset
        /// </summary>
        public const int ErrMalleablePin2Length = 304;
        /// <summary>
        /// Shoe Width is less than zero of Shoe of Type DrawBeam_T Method
        /// </summary>
        public const int ErrShoeOtypeTWidth = 305;
        /// <summary>
        /// Shoe Length is less than zero of Shoe of Type DrawBeam_T Method
        /// </summary>
        public const int ErrShoeOtypeTLength = 306;
        /// <summary>
        /// Error in DrawBeam_T Method
        /// </summary>
        public const int ErrDrawBeam_T = 307;
        /// <summary>
        /// Error in DrawBeam_B2BL Method
        /// </summary>
        public const int ErrDrawBeam_B2BL = 308;
        /// <summary>
        /// Error in DrawBeam_Cradle Method
        /// </summary>
        public const int ErrDrawBeam_Cradle = 309;
        /// <summary>
        /// Error in DrawBeam_C Method
        /// </summary>
        public const int ErrDrawBeam_C = 310;
        /// <summary>
        /// Error in DrawBeam_2L Method
        /// </summary>
        public const int ErrDrawBeam_2L = 311;
        /// <summary>
        /// Error in DrawBeam_W Method
        /// </summary>
        public const int ErrDrawBeam_W = 312;
        /// <summary>
        /// Error in DrawBeam_Irregular Method
        /// </summary>
        public const int ErrDrawBeam_Irregular = 313;
        /// <summary>
        /// Shoe Length is less than zero of Shoe in Type 2L Method
        /// </summary>
        public const int ErrShoeOtypeDrawBeam_2LLength = 314;
        /// <summary>
        /// Shoe Length is less than zero of Shoe in DrawBeam_Irregular Method
        /// </summary>
        public const int ErrDrawBeam_IrregularLength = 315;
        /// <summary>
        ///Shoe Width is less than zero of Shoe in DrawBeam_Irregular Method
        /// </summary>
        public const int ErrDrawBeam_IrregularWidth = 316;
        /// <summary>
        /// DrawBeam_Irregular as  Lower spacing must be less than upper leg spacing 
        /// </summary>
        public const int ErrDrawBeam_IrregularSpacing = 317;
        /// <summary>
        /// Shoe Length is less than zero of Shoe in Type 2L Method
        /// </summary>
        public const int ErrShoeOtypeDrawBeam_2LLength1 = 318;
        /// <summary>
        /// Shoe Length is less than zero of Shoe in Type DrawBeam_B2BL Method
        /// </summary>
        public const int ErrShoeOtypeDrawBeam_B2BLLength = 319;
        /// <summary>
        /// Shoe Length is less than zero of Shoe in Type DrawBeam_Cradle Method
        /// </summary>
        public const int ErrShoeOtypeDrawBeam_CradleLength = 320;
        /// <summary>
        /// Shoe Length is less than zero of Shoe in Type DrawBeam_C Method
        /// </summary>
        public const int ErrShoeOtypeDrawBeam_CLength = 321;
        /// <summary>
        /// Shoe Width is less than zero of Shoe in Type DrawBeam_B2BL Method
        /// </summary>
        public const int ErrShoeOtypeDrawBeam_B2BLWidth = 322;
        /// <summary>
        /// Error in StarPlateShape Method
        /// </summary>
        public const int ErrStarPlateShape = 323;
        /// <summary>
        /// Error in DrawStandardSection Method
        /// </summary>
        public const int ErrDrawStandardSection = 324;
        /// <summary>
        /// Error in Shoe Method
        /// </summary>
        public const int ErrShoe = 325;
        /// <summary>
        /// Error in LoadShoeDataByQuery method
        /// </summary>
        public const int ErrLoadShoeDataByQuery = 326;
        /// <summary>
        /// Error in LoadUBoltDataByQuery method
        /// </summary>
        public const int ErrLoadUBoltDataByQuery = 327;
        /// <summary>
        /// Error in LoadStrapDataByQuery method
        /// </summary>
        public const int ErrLoadStrapDataByQuery = 328;
        /// <summary>
        /// Error in LoadSheildDataByQuery method
        /// </summary>
        public const int ErrLoadSheildDataByQuery = 329;
        /// <summary>
        /// Error in LoadSlidePlateDataByQuery method
        /// </summary>
        public const int ErrLoadSlidePlateDataByQuery = 330;
        /// <summary>
        /// Invalid Length Of DummyLeg value
        /// </summary>
        public const int ErrInvalidLengthOfDummyLeg = 331;
        /// <summary>
        /// Invalid Shape Of DummyLeg Stanchion
        /// </summary>
        public const int ErrInvalidTypeStanchion = 332;
        /// <summary>
        /// Invalid PipeDia  Of DummyLeg Stanchion
        /// </summary>
        public const int ErrInvalidPipeDiaStanchion = 333;
        /// <summary>
        /// Invalid BotShape Of Stanchion
        /// </summary>
        public const int ErrInvalidBotShapeStanchion = 334;
        /// <summary>
        /// Invalid Length For Botshape Of Stanchion
        /// </summary>
        public const int ErrInvalidBotShapeLengthStanchion = 335;
        /// <summary>
        /// InValid BotShape Of Stanchion
        /// </summary>
        public const int ErrInvalidReqBotShapeStanchion = 336;
        /// <summary>
        /// InValid Height For Bottom And Top Shapes
        /// </summary>
        public const int ErrInvalidBotTopHeightStanchion = 337;
        /// <summary>
        /// InValid BotHeight For VariableLength
        /// </summary>
        public const int ErrInvalidBotHeightStanchion = 338;
        /// <summary>
        /// InValid BotHeight For VariableLength
        /// </summary>
        public const int ErrNoHeightStanchion = 339;
        /// <summary>
        /// steel standard is not valid for stanch
        /// </summary>
        public const int ErrNoStanchforStandardsteel = 340;
        /// <summary>
        /// PipeDia Required
        /// </summary>
        public const int ErrPipeDiaReq = 341;
        /// <summary>
        /// Error in LoadDummyLegData
        /// </summary>
        public const int ErrLoadDummyLegData = 342;
        /// <summary>
        /// Error in LoadStanchShapeData
        /// </summary>
        public const int ErrLoadStanchShapeDataByQuery = 343;
        /// <summary>
        /// Error in LoadDummyLegData By Query 
        /// </summary>
        public const int ErrLoadDummyLegShapeDataByQuery = 344;
        /// <summary>
        /// Error in AddStanchionShape
        /// </summary>
        public const int ErrAddStanchShape = 345;
        /// <summary>
        /// Error in AddDummyLegShape
        /// </summary>
        public const int ErrAddDummyLegShape = 346;
        /// <summary>
        /// Top Plate Required For Stanch Shape
        /// </summary>
        public const int ErrNoTopPlate = 347;
        /// <summary>
        ///Moe Offset
        /// </summary>
        public const int ErrMoreOffset = 348;
        /// <summary>
        /// Moe Offset
        /// </summary>
        public const int ErrMoreOffset1 = 349;
        /// <summary>
        /// Moe pipeOD then DummyWidth
        /// </summary>
        public const int ErrMorePipeOD = 350;
        /// <summary>
        /// steel standard is not valid for dummy
        /// </summary>
        public const int ErrSteelDummy = 351;
        /// <summary>
        /// Error in AddPlateShapeWithBolts
        /// </summary>
        public const int ErrPlateshapeBolts = 352;
        /// <summary>
        /// Error in CreChamBoxLine
        /// </summary>
        public const int ErrCreChamBoxLine = 353;
        /// <summary>
        /// Error in AddDummyLegInputs
        /// </summary>
        public const int ErrAddDummyLegInputs = 354;
        /// <summary>
        /// Error in AddDummyLegOutputs
        /// </summary>
        public const int ErrAddDummyLegoutputs = 355;
        /// <summary>
        /// Error in AddShoeInputs
        /// </summary>
        public const int ErrAddShoeInputs = 355;
        /// <summary>
        /// Shoe Gap should Not be greater than zero
        /// </summary>
        public const int ErrShoeGap = 356;
        /// <summary>
        /// \nLength must be greater than 0.
        /// </summary>
        public const int ErrInvalidnLengthValue = 357;
        /// <summary>
        /// \nOffset PipeClamp Configuration is not applicable for a SplitPipe Ring.
        /// </summary>
        public const int ErrInvalidnOffset = 358;
        /// <summary>
        /// \nWidth3 is needed, Setting it to Height1 + Height2
        /// </summary>
        public const int ErrInvalidnWidth3 = 359;
        /// <summary>
        /// Error In LoadHolePortData
        /// </summary>
        public const int ErrLoadHolePortData = 360;
        /// <summary>
        /// Error In LoadRoutePortOutput
        /// </summary>
        public const int ErrLoadRoutePortData = 361;
        /// <summary>
        /// Error In AddRoutePortINput
        /// </summary>
        public const int ErrAddRoutePortInputs = 362;
        /// <summary>
        /// Error In LoadRoutePortData
        /// </summary>
        public const int ErrLoadRoutePortInputs = 363;
        /// <summary>
        /// Error In hs_Parts_RotatePort
        /// </summary>
        public const int ErrhsPartsRotatePort = 364;
        /// <summary>
        /// Error In hsPartsHoleRotatePort
        /// </summary>
        public const int ErrhsPartsHoleRotatePort = 365;
        /// <summary>
        /// Error In LoadStandardPortData
        /// </summary>
        public const int ErrStandardPortInputs = 366;
        /// <summary>
        /// Error In AddStandardPortInputs
        /// </summary>
        public const int ErrAddStandardPortInputs = 367;
        /// <summary>
        /// Error In AddStandardPortoutputs
        /// </summary>
        public const int ErrAddStandardPortoutputs = 368;
        /// <summary>
        ///Error In AddHolePortinputs
        /// </summary>
        public const int ErrAddHolePortinputs = 369;
        /// <summary>
        /// Error In AddHolePortoutputs
        /// </summary>
        public const int ErrAddHolePortoutputs = 370;
        /// <summary>
        /// Error In RoutePortOutput
        /// </summary>
        public const int ErrAddRoutePortoutputs = 371;
        /// <summary>
        /// Error Invalid elbow radius
        /// </summary>
        public const int ErrInvalidElbowRadius = 372;
        /// <summary>
        /// Error Invalid FacetoCenter
        /// </summary>
        public const int ErrInvalidFacetoCenter = 373;
        /// <summary>
        /// Error Invalid Invalid width length thickness
        /// </summary>
        public const int ErrInvalidLenWidthThick = 374;
        /// <summary>
        /// Error Invalid Invalid Diameter
        /// </summary>
        public const int ErrInvalidDiameter1NGZero = 375;
          /// <summary>
        /// Error Invalid Invalid ShieldPipeOD
        /// </summary>
        public const int ErrInvalidShieldPipeOD = 376;
        /// <summary>
        /// Error Invalid Stanchion Gap
        /// </summary>
        public const int ErrInvalidStanchionGapAndHeight = 377;

        
        /// <summary>
        /// Springrate must be greater than 0.
        /// </summary>
        public const int ErrInvalidnSpringrateValueGTZero = 378;
        /// <summary>
        /// OperatingLoad is below the minimum working range.
        /// </summary>
        public const int ErrInvalidnMinOperatingLoad = 379;
        /// <summary>
        /// OperatingLoad is above the maximum working range.
        /// </summary>
        public const int ErrInvalidnMaxOperatingLoad = 380;
        /// <summary>
        /// InstalledLoad is below the minimum working range.
        /// </summary>
        public const int ErrInvalidnMinInstalledLoad = 381;
        /// <summary>
        /// InstalledLoad is above the maximum working range.
        /// </summary>
        public const int ErrInvalidnMaxInstalledLoad = 382;
        /// <summary>
        /// Invalid Variability 
        /// </summary>
        public const int ErrInvalidnVariability = 383;
        /// <summary>
        /// Error occured while implementing BOM.
        /// </summary>
        public const int ErrVariableSpringBOMDescription = 384;
        /// <summary>
        /// CrossSection not found in Catalog
        /// </summary>
        public const int ErrCrossSectionNotFound = 385;
        /// <summary>
        /// Invald CC and lenAdj Values.
        /// </summary>
        public const int ErrInvalidCC = 386;

        /// <summary>
        /// Error In LoadStrutAData
        /// </summary>
        public const int ErrLoadStrutAData = 387;
        /// <summary>
        /// Error In TravelAllowance
        /// </summary>
        public const int ErrInvalidTravelAllowance = 388;
        /// <summary>
        /// Error In MinMaxLength
        /// </summary>
        public const int ErrInvalidMinMaxLength = 389;
        /// <summary>
        /// Error In StructA
        /// </summary>
        public const int ErrAddStrutAMethod = 390;
        /// <summary>
        /// Error In MakeShapeMethod
        /// </summary>
        public const int ErrMakeShapeMethod = 391;
        /// <summary>
        /// Error In LoadTurnbuckleByQuery
        /// </summary>
        public const int ErrLoadTurnbuckleByQuery = 392;
        /// <summary>
        /// Error In TotalTravel
        /// </summary>
        public const int ErrInvTotalTravel = 393;
        /// <summary>
        /// Error In TotalTravelUnits
        /// </summary>
        public const int ErrInvTotalTravelUnits = 394;
        /// <summary>
        /// Error In Configuration
        /// </summary>
        public const int ErrUConfig = 395;
        /// <summary>
        /// Error In Configuration
        /// </summary>
        public const int ErrGConfig = 396;
        /// <summary>
        /// Error In StellDetails
        /// </summary>
        public const int ErrCrossSectionStellDetails = 397;
        /// <summary>
        /// Error In LoadWBABoltDataByQuery
        /// </summary>
        public const int ErrLoadWBABoltDataByQuery = 398;
        /// <summary>
        /// Error In GetSmartShapeType
        /// </summary>
        public const int ErrGetSmartShapeType = 399;
        /// <summary>
        /// Error In LoadRefData
        /// </summary>
        public const int ErrLoadRefData = 400;
        /// <summary>
        /// Error In LoadClevisDataByQuery
        /// </summary>
        public const int ErrLoadClevisDataByQuery = 401;
        /// <summary>
        /// Error In LoadPinDataByQuery
        /// </summary>
        public const int ErrLoadPinDataByQuery = 402;
        /// <summary>
        /// Error In GetConstantSizeRowProperties
        /// </summary>
        public const int ErrGetConstantSizeRowProperties = 403;
        /// <summary>
        /// Error In GetConstantTravelRowAttributes
        /// </summary>
        public const int ErrGetConstantTravelRowAttributes = 404;
        /// <summary>
        /// Error In GetConstantLoadAttributes
        /// </summary>
        public const int ErrGetConstantLoadAttributes = 405;
        /// <summary>
        /// Error In LoadRodDataByQuery
        /// </summary>
        public const int ErrLoadRodDataByQuery = 406;
        /// <summary>
        /// Error occured while implementing Constant BOM.
        /// </summary>
        public const int ErrConstantBOMDescription = 407;
        /// <summary>
        /// Error occured while implementing StrutA_BOM.
        /// </summary>
        public const int ErrStrutABOMDescription = 408;
        /// <summary>
        /// Error Invalid GroutAdditionalInputs
        /// </summary>
        public const int ErrGroutAdditionalInputs = 409;
        /// <summary>
        /// Error Invalid LoadGroutData
        /// </summary>
        public const int ErrLoadGroutData = 410;
        /// <summary>
        /// Error Invalid LoadGroutData
        /// </summary>
        public const int ErrGroutAdditionalOutputs = 411;
        /// <summary>
        /// Error Invalid LoadGroutData
        /// </summary>
        public const int ErrAddGroutCircular = 412;
        /// <summary>
        /// Error Invalid LoadGroutData
        /// </summary>
        public const int ErrAddGrout = 413;
        /// <summary>
        /// Error Invalid LoadGroutData
        /// </summary>
        public const int ErrGroutBottomWidth1 = 414;
        /// <summary>
        /// Error Invalid LoadGroutData
        /// </summary>
        public const int ErrGroutTopWidth1 = 415;
        /// <summary>
        /// Error Invalid LoadGroutData
        /// </summary>
        public const int ErrGroutHeight = 416;
        /// <summary>
        /// Error Invalid SetPosition
        /// </summary>
        public const int ErrInvalidSetposition = 417;
        /// <summary>
        /// Error In AddPipeRPadInputs
        /// </summary>
        public const int ErrAddPipeRPadInputs = 418;
        /// <summary>
        /// Error In AddPipeRPadOutputs
        /// </summary>
        public const int ErrAddPipeRPadOutputs = 419;
        /// <summary>
        /// Error In LoadPipeRPadData
        /// </summary>
        public const int ErrLoadPipeRPadData = 420;
        /// <summary>
        /// Error In AddElbowRPadInputs
        /// </summary>
        public const int ErrAddElbowRPadInputs = 421;
        /// <summary>
        /// Error In AddElbowRPadOutputs
        /// </summary>
        public const int ErrAddElbowRPadOutputs = 422;
        /// <summary>
        /// Error In LoadElbowRPadData
        /// </summary>
        public const int ErrLoadElbowRPadData = 423;
        /// <summary>
        /// Invalid R-Pad Dimensions
        /// </summary>
        public const int ErrInvalidRPadDimension = 424;
        /// <summary>
        /// Error In AddPipeRPad
        /// </summary>
        public const int ErrAddPipeRPad = 425;
        /// <summary>
        /// Error In AddElbowRPad
        /// </summary>
        public const int ErrAddElbowRPad = 426;
        /// <summary>
        /// Error Occured in LoadSpreaderBeamData
        /// </summary>
        public const int ErrLoadSpreaderBeamData = 427;
        /// <summary>
        /// Error Occured in LoadClevisHangerDataByQuery
        /// </summary>
        public const int ErrLoadClevisHangerDataByQuery = 428;
        /// <summary>
        /// Error Occured in LoadSwivelDataByQuery
        /// </summary>
        public const int ErrLoadSwivelDataByQuery = 429;
        /// <summary>
        /// Error Occured in LoadWBAHoleDataByQuery
        /// </summary>
        public const int ErrLoadWBAHoleDataByQuery = 430;
        /// <summary>
        /// Error in Length
        /// </summary>
        public const int ErrInvalidSteelLength = 431;
        /// <summary>
        /// Error  in SteelName
        /// </summary>
        public const int ErrInvalidSteelName = 432;
        /// <summary>
        /// Error  in SteelStandard
        /// </summary>
        public const int ErrInvalidSteelStandard = 433;
        /// <summary>
        /// Error  in SteelType
        /// </summary>
        public const int ErrInvalidSteelType = 434;
        /// <summary>
        /// Error occured in AddSpreaderBeamInputs
        /// </summary>
        public const int ErrAddSpreaderBeamInputs = 435;
        /// <summary>
        /// Error occured in AddDummyLegInsulationShape
        /// </summary>
        public const int ErrInsulationDummy = 436;
        /// <summary>
        /// Invalid Width1 and Width2 value
        /// </summary>
        public const int ErrInvalidWidth1Width2NZero = 437;
        /// <summary>
        /// No data Entered fro VerticalPlate Shape
        /// </summary>
        public const int ErrVerticalPlateShapeRequired=438;
        /// <summary>
        /// Total height of vertical plate is less than the minimum allowable height.
        /// </summary>
        public const int ErrInvalidMinVerticalPLateHeight = 439;
        /// <summary>
        /// Total height of vertical plate is greater than the maximum allowable height.
        /// </summary>
        public const int ErrInvalidMaxVerticalPLateHeight = 440;
        /// <summary>
        /// The C-C Value is less than the minimum allowable C-C Value.
        /// </summary>
        public const int ErrIvalidminCCValue = 441;
        /// <summary>
        /// The C-C Value is greater than the maximum allowable C-C Value.
        /// </summary>
        public const int ErrIvalidmaxCCValue = 442;
        /// <summary>
        ///Error in AddPlateShapeWithHole method.
        /// </summary>
        public const int ErrAddPlateShapeWithHoleMethod = 443;
        /// <summary>
        ///Error in Hole1HInset input value.
        /// </summary>
        public const int ErrAddPlateShapeWithHoleHoleH1Inset = 444;
        /// <summary>
        ///Error in Hole1VInset input value.
        /// </summary>
        public const int ErrAddPlateShapeWithHoleHole1VInset = 445;
        /// <summary>
        ///Error in Hole1Diameter input value.
        /// </summary>
        public const int ErrAddPlateShapeWithHoleHole1Diameter = 446;
        /// <summary>
        /// Invalid DummylegInsulation Thickness
        /// </summary>
        public const int ErrDummyInsulationThick = 447;
        /// <summary>
        /// Invalid DummylegInsulation Length
        /// </summary>
        public const int ErrInvalidInsulationLength = 448;
		/// <summary>
        /// Given Data is not in the range 
        /// </summary>
        public const int ErrDataOutsideRange = 449;
        /// <summary>
        /// Invalid Grip Clamp Data 
        /// </summary>
        public const int ErrLoadGripclampData = 450;
        /// <summary>
        /// Unable to create Grip Clamp Graphics 
        /// </summary>
        public const int ErrAddGripclampData = 451;
        /// <summary>
        /// Invalid Horizontal Traveler Data 
        /// </summary>
        public const int ErrLoadHorizontalTravelerData = 452;
        /// <summary>
        /// Unable to create Horizontal Traveler Graphics 
        /// </summary>
        public const int ErrAddHorizontalTraveler = 453;
        /// <summary>
        /// Invalid Pipe Saddle Inputs
        /// </summary>
        public const int ErrAddPipeSaddleInputs = 454;
        /// <summary>
        /// Invalid Pipe Saddle Data
        /// </summary>
        public const int ErrLoadPipeSaddleData = 455;
        /// <summary>
        /// Unable to create Pipe Saddle Graphics 
        /// </summary>
        public const int ErrAddPipeSaddle = 456;
        /// <summary>
        /// Unable to create Bolt Graphics  
        /// </summary>
        public const int ErrAddBoltSmartpart = 457;
        /// <summary>
        /// Invalid Bolt Data
        /// </summary>
        public const int ErrLoadBoltSmartpartData = 458;
        /// <summary>
        /// Invalid Bolt Inputs 
        /// </summary>
        public const int ErrAddBoltSmartpartInputs = 459;
        /// <summary>
        /// Invalid Pipe Lug Data
        /// </summary>
        public const int ErrLoadPipeLugInputsData = 460;
        /// <summary>
        /// Unable to create Welding Lug Graphics 
        /// </summary>
        public const int ErrAddWeldingLug = 461;
        /// <summary>
        /// Invalid Gap for Bolt
        /// </summary>
        public const int ErrInvalidGap1OfBoltSmartpart = 462;
        /// <summary>
        /// Invalid Lug Plate Data
        /// </summary>s
        public const int ErrLoadLugPlateData = 463;
        /// <summary>
        /// Unable to create Lug Plate Graphics   
        /// </summary>
        public const int ErrAddLugPlate = 464;
    }
}



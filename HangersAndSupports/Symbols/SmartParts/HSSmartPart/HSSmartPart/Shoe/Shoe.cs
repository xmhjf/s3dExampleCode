//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   Shoe.cs
//   HSSmartPart,Ingr.SP3D.Content.Support.Symbols.Shoe
//   Author       :  Pradeep
//   Creation Date:  14-03-2013
//   Description:CR-CP- 222489        Initial Creation

//   Change History:
//   dd.mmm.yyyy     who     change description
//   14-03-2013     Pradeep  CR-CP-222489 Initial Creation
//   25/Mar/2013    Pradeep  DI-CP-228142  Modify the error handling for delivered H&S symbols 
//   20/Aug/2013    Vijaya   DM-CP-236506  Plate Shapes need to adjust automatically for Var Height Shoes
//   15-may-2014    Vinod    DI-CP-249570  Add missing smart parts in Anvil2010 catalog 
//   27-Aug-2014    Siva     TR-CP-251283  Added necessary optional inputs to control the shpaes even after changing the lenght width of the shoe
//   29-Sept-2014   NDR      DM-CP-241276  Shoe SmartPart - Need to be able to override the clamp width
//   17-Oct-2014    Chethan  CR-CP-253372  Add Maintenance Aspect to shoe smartpart    
//   02-Dec-2014    PVK      DI-CP-253817	Fix priority 2 items to .net SmartParts as a result of new testing 
//   12-12-2014     PVK      TR-CP-264951	Resolve P3 coverity issues found in November 2014 report
//   11-02-2015     Chethan  DI-CP-263820   Fix priority 3 items to .net SmartParts as a result of new testing  
//   05-03-2015     PVK     TR-CP-267603	Stiffener, EndPlate Heights are not adjusting when Slideplate shape is added
//   17-04-2015     JRM     CR-CP-206536   Implemented change where when weight is 0, and the input weight is 0, we calculate the weight based upon the volume and the material
//   10-06-2015     PVK	    TR-CP-274155	SmartPart TDL Errors should be corrected.
//   13.08.2015     PR      TR 276067	Few smart parts are having cache option as non-cached when they should be cached
//   30-11=2015     VDP     Integrate the newly developed SmartParts into Product(DI-CP-282644)
//   16-12-2015     Siva    TR-CP-283495  implemented automatic height calculations for end plate shape.
//   05-04-2016     Vinay   TR-CP-283495  Removed the redundant PlateDiaRatio Condition
//   20-07-2016     Ramya   TR 298244	Insul. Thick. Prop for Generic Insulated Clamped Shoe doesn’t accept + ve values
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using System.Collections.Generic;
using System.Text;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.ReferenceData.Middle.Services;

namespace Ingr.SP3D.Content.Support.Symbols
{
    //-----------------------------------------------------------------------------------
    //Namespace of this class is Ingr.SP3D.Content.Support.Symbols
    //It is recommended that customers specify namespace of their symbols to be
    //CompanyName.SP3D.Content.Specialization.
    //It is also recommended that if customers want to change this symbol to suit their
    //requirements, they should change namespace/symbol name so the identity of the modified
    //symbol will be different from the one delivered by Intergraph.
    //-----------------------------------------------------------------------------------
    [CacheOption(CacheOptionType.Cached)]
    [SymbolVersion("1.0.0.1")]
    [VariableOutputs]
    public class Shoe : SmartPartComponentDefinition, ICustomWeightCG
    //, ICustomWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "HSSmartPart,Ingr.SP3D.Content.Support.Symbols.Shoe"
        //----------------------------------------------------------------------------------
        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        //#region "Definition of Additional Inputs"
        public override IEnumerable<Input> AdditionalInputs
        {
            get
            {
                int endIndex;
                List<Input> additionalInputs = new List<Input>();
                AddShoeInputs(2, out endIndex, additionalInputs);
                // Add Optional Shoe Inputs
                additionalInputs.Add(new InputDouble(endIndex, "RepadLength", "RepadLength", 0, true));
                additionalInputs.Add(new InputDouble(++endIndex, "ShieldLength", "ShieldLength", 0, true));
                additionalInputs.Add(new InputDouble(++endIndex, "Shield2Length", "Shield2Length", 0, true));
                //  Add Optional Stiffner Inputs
                additionalInputs.Add(new InputDouble(++endIndex, "StiffLength", "StiffLength", 0, true));
                additionalInputs.Add(new InputDouble(++endIndex, "StiffCRad", "StiffCRad", 0, true));
                additionalInputs.Add(new InputDouble(++endIndex, "StiffCEndX", "StiffCEndX", 0, true));
                additionalInputs.Add(new InputDouble(++endIndex, "StiffCEndY", "StiffCEndY", 0, true));
                // Add Optional Shoe Input
                additionalInputs.Add(new InputDouble(++endIndex, "ClampAngle", "ClampAngle", 0, true));
                // New Insulation Thickness Property
                additionalInputs.Add(new InputDouble(++endIndex, "InsulationTh", "InsulationTh", 0, true));
                additionalInputs.Add(new InputDouble(++endIndex, "InsulationLength", "InsulationLength", 0, true));
                additionalInputs.Add(new InputDouble(++endIndex, "RepadThickness", "RepadThickness", 0, true));
                additionalInputs.Add(new InputDouble(++endIndex, "RibGap", "RibGap", 0, true));
                //GuideGap
                additionalInputs.Add(new InputDouble(++endIndex, "GuideGap", "GuideGap", 0, true));
                //LegSpace
                additionalInputs.Add(new InputDouble(++endIndex, "LegSpace", "LegSpace", 0, true));
                additionalInputs.Add(new InputDouble(++endIndex, "LegLowSpace", "LegLowSpace", 0, true));
                //ClampWidth
                additionalInputs.Add(new InputDouble(++endIndex, "ClampWidthSel", "ClampWidthSel", 1, true));
                additionalInputs.Add(new InputDouble(++endIndex, "ClampLenOvrHng", "ClampLenOvrHng", 0, true));
                additionalInputs.Add(new InputDouble(++endIndex, "ClampWidth1", "ClampWidth1", 0, true));
                additionalInputs.Add(new InputDouble(++endIndex, "ClampWidth2", "ClampWidth2", 0, true));
                //Shield and Repad Length
                additionalInputs.Add(new InputDouble(++endIndex, "ShieldLengthSel", "ShieldLengthSel", 1, true));
                additionalInputs.Add(new InputDouble(++endIndex, "ShldLenOvrHng", "ShldLenOvrHng", 0, true));
                additionalInputs.Add(new InputDouble(++endIndex, "PadLengthSel", "PadLengthSel", 1, true));
                additionalInputs.Add(new InputDouble(++endIndex, "PadLenOvrHng", "PadLenOvrHng", 0, true));
                //More Clamp Inputs
                additionalInputs.Add(new InputDouble(++endIndex, "ClampDiameter1", "ClampDiameter1", 0, true));
                additionalInputs.Add(new InputDouble(++endIndex, "ClampBoltOffset1", "ClampBoltOffset1", 0, true));
                additionalInputs.Add(new InputDouble(++endIndex, "ClampBoltOffset2", "ClampBoltOffset2", 0, true));
                additionalInputs.Add(new InputDouble(++endIndex, "ClampBoltOffset3", "ClampBoltOffset3", 0, true));
                additionalInputs.Add(new InputDouble(++endIndex, "ClampBoltOffset4", "ClampBoltOffset4", 0, true));
                additionalInputs.Add(new InputDouble(++endIndex, "ClampBoltOffset5", "ClampBoltOffset5", 0, true));
                additionalInputs.Add(new InputDouble(++endIndex, "ClampBoltOffset6", "ClampBoltOffset6", 0, true));
                additionalInputs.Add(new InputDouble(++endIndex, "ClampDiameter4", "ClampDiameter4", 0, true));
                additionalInputs.Add(new InputDouble(++endIndex, "ClampWidth4", "ClampWidth4", 0, true));
                additionalInputs.Add(new InputDouble(++endIndex, "MaintenanceThickness", "MaintenanceThickness", 0, true));
                additionalInputs.Add(new InputDouble(++endIndex, "CreateMaintenanceAspect", "CreateMaintenanceAspect", 0, true));

                return additionalInputs;
            }
        }
        #endregion

        #region "Definitions of Aspects and their Outputs"
        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Route", "Route")]
        [SymbolOutput("Steel", "Steel")]
        [SymbolOutput("Weld", "Weld")]
        [SymbolOutput("Shoe", "Shoe")]
        public AspectDefinition m_oSymbolic;

        [Aspect("Maintenance", "Maintenance Aspect", AspectID.Maintenance)]
        public AspectDefinition m_Maintenance;

        #endregion

        #region "Construct Outputs"
        /// <summary>
        /// Construct symbol outputs in aspects.
        /// </summary>
        /// <remarks></remarks>
        protected override void ConstructOutputs()
        {
            try
            {
                Part part = (Part)m_PartInput.Value;

                ShoeInputs shoeInputs;
                PlateInputs endPlateInputs, ribInputs, gussetInputs, stiffenerInputs, additionalPlateInputs;
                ShieldInputs repadInputs, shieldInputs, shield2Inputs, clamp2Inputs;
                PipeClampInputs clampInputs;
                UBoltInputs uBoltInputs;
                StrapInputs strapInputs;
                SlidePlateInputs slidePlateInputs;
                GuideInputs guideInputs;
                double x1Offset, x2Offset, guideLength, clampThickness, uboltThichness, strapThickness, shoeShapeHeight, currentTotalHeight;
                Boolean bClamp, bRepad;
                Double SlopeAngleDeg, ClampOffset1;
                StringBuilder error = new StringBuilder();
                double shoeoff = 0;
                String shoeName, endPlateName, repadName, clampName, uboltName, strapName, gussetName, shieldName, shield2Name, stiffenerName, slidePlateName, guideName, ribName;
                double[] xLocation;
                double shoeHeight, shoeWidth, shoeLength, shoeGap, diameter1, slopeAngle, slideVertOffset, shieldODia, shield2ODia, repadODia, clampOdia, clamp2Odia, repadOffset1, shieldOffset1, shield2Offset1, outerMostRadius, tempHeight, gussetHeight;
                double ribQty, outerMostDia, upwardRibOffset, ribHeight, plateAngle, offsettoShoeBottom = 0, clampSlopeOffset, shieldSlopeOffset, shield2SlopeOffset, repadSlopeOffset, addPlateOffset1, uboltSlopeOffset, strapSlopeOffset;
                int shoeQty;
                //End Plate Attributes
                double endPlateNum;
                //Multi Positioning Attributes for Repad
                double multi1Qty, multi1LocateBy, multi1Location;
                //Shoe Pad Override Properties
                double endPlateHeight;
                //Multi Positioning Attributes for Clamps
                Double multi2Qty, multi2LocateBy, multi2Location;
                //Multi Positioning Attributes for Ubolt
                Double multi3Qty, multi3LocateBy, multi3Location;
                //Multi Positioning Attributes for Strap
                Double multi4Qty, multi4LocateBy, multi4Location;

                //Multi Positioning Attributes for Gusset
                Double multi5Qty, multi5LocateBy, multi5Location, gussetGap;

                //Multi Positioning Attributes for Guide
                Double multi6Qty, multi6LocateBy, multi6Location, guideVerticalOffset;
                //Multi Positioning Attributes for Shield
                Double multi7Qty, multi7LocateBy, multi7Location;

                //Multi Positioning Attributes for Stiffener
                Double multi8Qty, multi8LocateBy, multi8Location, stiffenerVerticalOffset;

                //Multi Positioning Attributes for Slide Plates
                Double multi9Qty, multi9LocateBy, multi9Location;

                //Multi Positioning Attributes for shield2
                Double multi10Qty, multi10LocateBy, multi10Location, endPlateOffset;

                //Additional Plates
                String additionalPlateName;
                Double plateNumber, centerDiameter, plateRotationAngle;
                //Clamp2
                string clamp2Name;
                Double multi11Qty, multi11LocateBy, multi11Location;

                Matrix4X4 matrix = new Matrix4X4();
                //Retrieve XLS data through array of inputs
                int j = 2;

                shoeName = GetStringeorDefaultValue(j);
                shoeHeight = GetDoubleorDefaultValue(++j);
                shoeWidth = GetDoubleorDefaultValue(++j);
                shoeLength = GetDoubleorDefaultValue(++j);
                shoeGap = GetDoubleorDefaultValue(++j);
                if (shoeGap > 0)
                    RaiseSmartPartTDLError(SmartPartSymbolResourceIDs.ErrShoeGap, "The Shoe Gap should not be grater than zero");

                diameter1 = GetDoubleorDefaultValue(++j);
                slopeAngle = GetDoubleorDefaultValue(++j);
                shoeQty = (int)GetDoubleorDefaultValue(++j);
                //EndPlate
                endPlateNum = GetDoubleorDefaultValue(++j);
                endPlateHeight = GetDoubleorDefaultValue(++j);
                endPlateName = GetStringeorDefaultValue(++j);
                endPlateOffset = GetDoubleorDefaultValue(++j);
                //Inputs for Repad
                repadName = GetStringeorDefaultValue(++j);
                multi1Qty = GetDoubleorDefaultValue(++j);
                multi1LocateBy = GetDoubleorDefaultValue(++j);
                multi1Location = GetDoubleorDefaultValue(++j);
                repadOffset1 = GetDoubleorDefaultValue(++j);
                //Inputs for Clamps
                clampName = GetStringeorDefaultValue(++j);
                multi2Qty = GetDoubleorDefaultValue(++j);
                multi2LocateBy = GetDoubleorDefaultValue(++j);
                multi2Location = GetDoubleorDefaultValue(++j);
                ClampOffset1 = GetDoubleorDefaultValue(++j);
                //Inputs for Ubolts
                uboltName = GetStringInputValue(++j);
                multi3Qty = GetDoubleorDefaultValue(++j);
                multi3LocateBy = GetDoubleorDefaultValue(++j);
                multi3Location = GetDoubleorDefaultValue(++j);
                //Inputs for Strap
                strapName = GetStringeorDefaultValue(++j);
                multi4Qty = GetDoubleorDefaultValue(++j);
                multi4LocateBy = GetDoubleorDefaultValue(++j);
                multi4Location = GetDoubleorDefaultValue(++j);
                gussetName = GetStringeorDefaultValue(++j);
                multi5Qty = GetDoubleorDefaultValue(++j);
                multi5LocateBy = GetDoubleorDefaultValue(++j);
                multi5Location = GetDoubleorDefaultValue(++j);
                gussetGap = GetDoubleorDefaultValue(++j);
                gussetHeight = GetDoubleorDefaultValue(++j);
                guideName = GetStringeorDefaultValue(++j);
                multi6Qty = GetDoubleorDefaultValue(++j);
                multi6LocateBy = GetDoubleorDefaultValue(++j);
                multi6Location = GetDoubleorDefaultValue(++j);
                guideVerticalOffset = GetDoubleorDefaultValue(++j);
                //Inputs for Shield
                shieldName = GetStringeorDefaultValue(++j);
                multi7Qty = GetDoubleorDefaultValue(++j);
                multi7LocateBy = GetDoubleorDefaultValue(++j);
                multi7Location = GetDoubleorDefaultValue(++j);
                shieldOffset1 = GetDoubleorDefaultValue(++j);
                //Inputs for Stiffener
                stiffenerName = GetStringeorDefaultValue(++j);
                multi8Qty = GetDoubleorDefaultValue(++j);
                multi8LocateBy = GetDoubleorDefaultValue(++j);
                multi8Location = GetDoubleorDefaultValue(++j);
                stiffenerVerticalOffset = GetDoubleorDefaultValue(++j);
                //Inputs for Slide Plates
                slidePlateName = GetStringeorDefaultValue(++j);
                multi9Qty = GetDoubleorDefaultValue(++j);
                multi9LocateBy = GetDoubleorDefaultValue(++j);
                multi9Location = GetDoubleorDefaultValue(++j);
                //Inputs for Shield2
                shield2Name = GetStringeorDefaultValue(++j);
                multi10Qty = GetDoubleorDefaultValue(++j);
                multi10LocateBy = GetDoubleorDefaultValue(++j);
                multi10Location = GetDoubleorDefaultValue(++j);
                shield2Offset1 = GetDoubleorDefaultValue(++j);
                //Additional plates
                additionalPlateName = GetStringeorDefaultValue(++j);
                plateNumber = GetDoubleorDefaultValue(++j);
                centerDiameter = GetDoubleorDefaultValue(++j);
                plateRotationAngle = GetDoubleorDefaultValue(++j);
                plateAngle = GetDoubleorDefaultValue(++j);
                addPlateOffset1 = GetDoubleorDefaultValue(++j);
                clamp2Name = GetStringeorDefaultValue(++j);
                multi11Qty = GetDoubleorDefaultValue(++j);
                multi11LocateBy = GetDoubleorDefaultValue(++j);
                multi11Location = GetDoubleorDefaultValue(++j);
                //'Rib
                ribName = GetStringeorDefaultValue(++j);
                ribHeight = GetDoubleorDefaultValue(++j);
                ribQty = GetDoubleorDefaultValue(++j);

                //Optional Shoe Inputs
                double dExtensionAdjOffset = 0, repadLength, shieldLength, shield2Length, clampAngle;

                repadLength = GetDoubleorDefaultValueonPart(part, "IJUAhsRepadLength", "RepadLength", ++j);
                shieldLength = GetDoubleorDefaultValueonPart(part, "IJUAhsShieldLength", "ShieldLength", ++j);
                shield2Length = GetDoubleorDefaultValueonPart(part, "IJUAhsShield2Length", "Shield2Length", ++j);
                //Optional Stiffener Inputs
                Double stifflength, stiffCurveRadious, stiffCurveEndX, stiffCurveEndY;
                stifflength = GetDoubleorDefaultValueonPart(part, "IJOAhsStiff", "StiffLength", ++j);
                stiffCurveRadious = GetDoubleorDefaultValueonPart(part, "IJOAhsStiff", "StiffCRad", ++j);
                stiffCurveEndX = GetDoubleorDefaultValueonPart(part, "IJOAhsStiff", "StiffCEndX", ++j);
                stiffCurveEndY = GetDoubleorDefaultValueonPart(part, "IJOAhsStiff", "StiffCEndY", ++j);
                clampAngle = GetDoubleorDefaultValueonPart(part, "IJOAhsClampAngle", "ClampAngle", ++j);
                //Optional inputs for Insulation
                Double insulationThickness, insulationLength, repadThickness, ribGap;
                insulationThickness = GetDoubleorDefaultValueonPart(part, "IJOAhsInsulationTh", "InsulationTh", ++j);
                insulationLength = GetDoubleorDefaultValueonPart(part, "IJOAhsInsulationL", "InsulationLength", ++j);
                repadThickness = GetDoubleorDefaultValueonPart(part, "IJUAhsRepadThickness", "RepadThickness", ++j);
                ribGap = GetDoubleorDefaultValueonPart(part, "IJUAhsRibGap", "RibGap", ++j);
                //Optional inputs for Guide
                Double guideGap, LegSpace, LeglowSpace;
                guideGap = GetDoubleorDefaultValueonPart(part, "IJUAhsGuideGap", "GuideGap", ++j);
                LegSpace = GetDoubleorDefaultValueonPart(part, "IJUAhsShoeLegSpace", "LegSpace", ++j);
                LeglowSpace = GetDoubleorDefaultValueonPart(part, "IJUAhsShoeLegLowSpace", "LegLowSpace", ++j);
                //Clamp inputs
                Double ClampLenOvrHng, ClampWidth1, ClampWidth2;
                int ClampWidthSel;
                ClampWidthSel = (int)GetDoubleorDefaultValueonPart(part, "IJUAhsClampWidthSel", "ClampWidthSel", ++j);
                ClampLenOvrHng = GetDoubleorDefaultValueonPart(part, "IJUAhsClampLenOvrHng", "ClampLenOvrHng", ++j);
                ClampWidth1 = GetDoubleorDefaultValueonPart(part, "IJUAhsClampWidth1", "ClampWidth1", ++j);
                ClampWidth2 = GetDoubleorDefaultValueonPart(part, "IJUAhsClampWidth2", "ClampWidth2", ++j);
                //Shield and Repad optional inputs
                Double ShieldLenOvrHng, PadLenOveHng;
                int PadLengthSel, ShieldLengthSel;
                ShieldLengthSel = (int)GetDoubleorDefaultValueonPart(part, "IJUAhsShieldLengthSel", "ShieldLengthSel", ++j);
                ShieldLenOvrHng = GetDoubleorDefaultValueonPart(part, "IJUAhsShieldLenOvrHng", "ShldLenOvrHng", ++j);
                PadLengthSel = (int)GetDoubleorDefaultValueonPart(part, "IJUAhsPadLengthSel", "PadLengthSel", ++j);
                PadLenOveHng = GetDoubleorDefaultValueonPart(part, "IJUAhsPadLenOvrHng", "PadLenOvrHng", ++j);
                //More Clamp Inputs
                Double ClampDiameter1, ClampBoltOffset1, ClampBoltOffset2, ClampBoltOffset3, ClampBoltOffset4, ClampBoltOffset5, ClampBoltOffset6, ClampDiameter4, ClampWidth4;
                ClampDiameter1 = GetDoubleorDefaultValueonPart(part, "IJUAhsClampDiameter1", "ClampDiameter1", ++j);
                ClampBoltOffset1 = GetDoubleorDefaultValueonPart(part, "IJUAhsClampBoltOff1", "ClampBoltOffset1", ++j);
                ClampBoltOffset2 = GetDoubleorDefaultValueonPart(part, "IJUAhsClampBoltOff2", "ClampBoltOffset2", ++j);
                ClampBoltOffset3 = GetDoubleorDefaultValueonPart(part, "IJUAhsClampBoltOff3", "ClampBoltOffset3", ++j);
                ClampBoltOffset4 = GetDoubleorDefaultValueonPart(part, "IJUAhsClampBoltOff4", "ClampBoltOffset4", ++j);
                ClampBoltOffset5 = GetDoubleorDefaultValueonPart(part, "IJUAhsClampBoltOff5", "ClampBoltOffset5", ++j);
                ClampBoltOffset6 = GetDoubleorDefaultValueonPart(part, "IJUAhsClampBoltOff6", "ClampBoltOffset6", ++j);
                ClampDiameter4 = GetDoubleorDefaultValueonPart(part, "IJUAhsClampDiameter4", "ClampDiameter4", ++j);
                ClampWidth4 = GetDoubleorDefaultValueonPart(part, "IJUAhsClampWidth4", "ClampWidth4", ++j);
                Double maintenanceThickness = GetDoubleorDefaultValueonPart(part, "IJUAhsMaintenanceTh", "MaintenanceThickness", ++j);
                bool createMaintenanceAspect=false;
                createMaintenanceAspect = Convert.ToBoolean(GetDoubleInputValue(++j));
                

                //=============================='
                //'Load The Shape Data Structures'
                //'=============================='
                //'Load Shoe Shape data
                shoeInputs = LoadShoeDataByQuery(shoeName);
                //Load End Plate data by query
                endPlateInputs = LoadPlateDataByQuery(endPlateName);
                //Load Repad data by query
                repadInputs = LoadSheildDataByQuery(repadName);
                //Load PipeClamp Data by query
                clampInputs = LoadPipeClampDataByQuery(clampName);
                // hs_Parts_LoadPipeClampDataByQuery tClamp, sClamp
                if (HgrCompareDoubleService.cmpdbl(clampAngle, 0) == false)
                    clampInputs.Angle3 = clampAngle;
                if (HgrCompareDoubleService.cmpdbl(ClampDiameter1, 0) == false )
                    clampInputs.Diameter1 = ClampDiameter1;
                if (HgrCompareDoubleService.cmpdbl(ClampBoltOffset1, 0) == false)
                    clampInputs.BoltRow1.Offset = ClampBoltOffset1;
                if (HgrCompareDoubleService.cmpdbl(ClampBoltOffset2, 0) == false)
                    clampInputs.BoltRow2.Offset = ClampBoltOffset2;
                if (HgrCompareDoubleService.cmpdbl(ClampBoltOffset3, 0) == false)
                    clampInputs.BoltRow3.Offset = ClampBoltOffset3;
                if (HgrCompareDoubleService.cmpdbl(ClampBoltOffset4, 0) == false)
                    clampInputs.BoltRow4.Offset = ClampBoltOffset4;
                if (HgrCompareDoubleService.cmpdbl(ClampBoltOffset5, 0) == false)
                    clampInputs.BoltRow5.Offset = ClampBoltOffset5;
                if (HgrCompareDoubleService.cmpdbl(ClampBoltOffset6, 0) == false)
                    clampInputs.BoltRow6.Offset = ClampBoltOffset6;
                if (HgrCompareDoubleService.cmpdbl(ClampDiameter4, 0) == false)
                    clampInputs.Diameter4 = ClampDiameter4;
                if (HgrCompareDoubleService.cmpdbl(ClampWidth4, 0) == false)
                    clampInputs.Width4 = ClampWidth4;

                if (insulationThickness > 0)
                {
                    if (clampInputs.Diameter1 > 0)
                        clampInputs.Diameter1 = insulationThickness * 2 + clampInputs.Diameter1;
                    if (clampInputs.Diameter4 > 0)
                        clampInputs.Diameter4 = insulationThickness * 2 + clampInputs.Diameter4;
                    clampInputs.Height1 = clampInputs.Height1 + insulationThickness;
                    clampInputs.Height2 = clampInputs.Height2 + insulationThickness;

                    clampInputs.BoltRow1.Offset = clampInputs.BoltRow1.Offset + insulationThickness;
                    clampInputs.BoltRow2.Offset = clampInputs.BoltRow2.Offset + insulationThickness;
                    clampInputs.BoltRow3.Offset = clampInputs.BoltRow3.Offset + insulationThickness;
                    clampInputs.BoltRow4.Offset = clampInputs.BoltRow4.Offset + insulationThickness;
                    clampInputs.BoltRow5.Offset = clampInputs.BoltRow5.Offset + insulationThickness;
                    clampInputs.BoltRow6.Offset = clampInputs.BoltRow6.Offset + insulationThickness;

                    if (insulationLength > 0)
                        clampInputs.Width4 = insulationLength;
                }

                //Load U-Bolt Data By Query
                uBoltInputs = LoadUBoltDataByQuery(uboltName);
                //Load Strap Data By Query
                strapInputs = LoadStrapDataByQuery(strapName);
                //Load Gusset Data By Query
                gussetInputs = LoadPlateDataByQuery(gussetName);
                //Load Guide data by query
                guideInputs = LoadGuideDataByQuery(guideName);
                //Load shield data by query
                shieldInputs = LoadSheildDataByQuery(shieldName);

                if (insulationThickness > 0 && shieldInputs.PipeOD >0)
                    shieldInputs.PipeOD = shieldInputs.PipeOD + 2 * insulationThickness;

                
                // Load shield2 data by query
                shield2Inputs = LoadSheildDataByQuery(shield2Name);

                if (shield2Length > 0)
                {
                    shield2Inputs.Length1 = shield2Length;
                    shield2Inputs.Length2 = shield2Length;
                    shield2Inputs.Length3 = shield2Length;
                }
                
                // Load Stiffener Data By Query
                stiffenerInputs = LoadPlateDataByQuery(stiffenerName);
                if (stifflength > 0)
                {
                    stiffenerInputs.length1 = stifflength;
                    stiffenerInputs.curvedEndRad = stiffCurveRadious;
                    stiffenerInputs.curvedEndX = stiffCurveEndX;
                    stiffenerInputs.curvedEndY = stiffCurveEndY;
                }

                //'Load slideplate data into 'tSlideplate' struct based on query done using 'sShoe'
                slidePlateInputs = LoadSlidePlateDataByQuery(slidePlateName);

                //'Clamp2 data by query
                clamp2Inputs = LoadSheildDataByQuery(clamp2Name);

                //'Load Additional Plate Data By Query
                additionalPlateInputs = LoadPlateDataByQuery(additionalPlateName);

                ribInputs = LoadPlateDataByQuery(ribName);

                if (base.ToDoListMessage != null)
                    if (base.ToDoListMessage.Type == ToDoMessageTypes.ToDoMessageError)
                        return;

                currentTotalHeight = 0;
                slideVertOffset = slidePlateInputs.Thickness1 + slidePlateInputs.Thickness2 + slidePlateInputs.Thickness3 + slidePlateInputs.Thickness4 + slidePlateInputs.Thickness5 + slidePlateInputs.Thickness6;

                if (clampName != "" && clampName != "No Value")
                    bClamp = true;
                else
                    bClamp = false;

                if (repadName != "" && repadName != "No Value")
                    bRepad = true;
                else
                    bRepad = false;

                // Set all overides
                try
                {
                    if (HgrCompareDoubleService.cmpdbl(shoeWidth, 0) == false)
                        shoeInputs.ShoeWidth = shoeWidth;
                }
                catch { shoeInputs.ShoeWidth = 0; }
                // ShoeAngle
                try
                {
                    if (HgrCompareDoubleService.cmpdbl(slopeAngle, 0) == false)
                        shoeInputs.SlopeAngle = slopeAngle;
                }
                catch { shoeInputs.SlopeAngle = 0; }
                // Shoe Length
                try
                {
                    if (HgrCompareDoubleService.cmpdbl(shoeLength, 0) == false)
                        shoeInputs.ShoeLength = shoeLength;
                }
                catch { shoeInputs.ShoeLength = 0; }

                try
                {
                    if (HgrCompareDoubleService.cmpdbl(shoeInputs.VerticalPlateLength , 0) == true)
                        shoeInputs.VerticalPlateLength = shoeInputs.ShoeLength;
                }
                catch { shoeInputs.VerticalPlateLength = 0; }

                try
                {
                    if (endPlateHeight > 0)
                        endPlateInputs.length1 = endPlateHeight;
                }
                catch { endPlateInputs.length1 = 0; }

                if (HgrCompareDoubleService.cmpdbl(gussetHeight, 0) == false)
                    gussetInputs.length1 = gussetHeight;
                if (ribHeight > 0)
                    ribInputs.length1 = ribHeight;

                //warn for common Errors
                if (bClamp == true && bRepad == true)
                    error.Append(SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrInvalidOffset1NLTZero, "clamp and repad cannot be specified simultaneously"));
                if (!(diameter1 > 0))
                    error.Append(SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrInvalidOffset1NLTZero, "Diameter1 is required"));
                if (shoeQty > 0 && shoeInputs.HasExtension == true)
                    error.Append(SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrInvalidOffset1NLTZero, "Extension plates are not suppoeted for multiple shoes"));
                if (error.Length > 0)
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, error.ToString());
                // move side ways when ext. plt is used on a T
                if (shoeInputs.HasExtension)
                {
                    if (shoeInputs.ShoeType == 1 && shoeInputs.ExtensionType == 0)
                        dExtensionAdjOffset = -((shoeInputs.ShoeThickness1 / 2) + (shoeInputs.TopPlateThickness / 2));
                    else if (shoeInputs.ShoeType == 1 && shoeInputs.ExtensionType == 1)
                        dExtensionAdjOffset = ((shoeInputs.ShoeThickness1 / 2) + (shoeInputs.TopPlateThickness / 2));
                }
                // Set up shield
                
                if (shieldName != "" && shieldName != "No Value")
                {
                    if (shieldInputs.PipeOD <= 0)
                    {
                        shieldInputs.PipeOD = diameter1;
                        if (insulationThickness > 0)
                            shieldInputs.PipeOD = shieldInputs.PipeOD + 2 * insulationThickness;
                    }
                }
                shieldODia = shieldInputs.PipeOD + shieldInputs.Thickness1 + shieldInputs.Thickness1;

                // set up SHIELD 2/Vapor barrier
                shield2ODia = shield2Inputs.PipeOD + shield2Inputs.Thickness1 + shield2Inputs.Thickness1;
                if (shield2Name != "" && shield2Name != "No Value")
                {
                    if (shield2Inputs.PipeOD <= 0)
                        shield2Inputs.PipeOD = GetGreaterValue(Val1: diameter1, Val2: shieldODia);
                }

                // Add Repad
                repadODia = repadInputs.PipeOD + repadInputs.Thickness1 + repadInputs.Thickness1;
                if (repadName != "" && repadName != "No value")
                {
                    if (repadInputs.PipeOD <= 0)
                        repadInputs.PipeOD = GetGreaterValue(Val1: diameter1, Val2: shieldODia, Val3: shield2ODia);
                }

                // Add Clamp
                
                clampThickness = clampInputs.Width1;
                if (clampName != "" && clampName != "No Value")
                {
                    if (clampInputs.Width2 > clampThickness)
                        clampThickness = clampInputs.Width2;
                    if (clampInputs.Diameter1 <= 0)
                    {
                        clampInputs.Diameter1 = GetGreaterValue(diameter1, shieldODia, shield2ODia, repadODia);
                        if (insulationThickness > 0 )
                            clampInputs.Diameter1 = insulationThickness * 2 + clampInputs.Diameter1;
                    }

                    SlopeAngleDeg = shoeInputs.SlopeAngle * 180 / Math.PI;//deg
                }
                clampOdia = clampInputs.Diameter1 + clampInputs.Thickness1 + clampInputs.Thickness1;

                // Clamp2
                clamp2Odia = clamp2Inputs.PipeOD + clamp2Inputs.Thickness1 + clamp2Inputs.Thickness1;
                if (clamp2Name != "" && clamp2Name != "No Value")
                {
                    xLocation = new double[int.Parse(multi11Qty.ToString())];
                    for (int index = 0; index < multi11Qty; index++)
                    {
                        xLocation[index] = ShieldMultiPosition(shoeInputs.ShoeLength, multi11Qty, multi11LocateBy, multi11Location, clamp2Inputs.Length1)[index];
                        matrix = new Matrix4X4();
                        matrix.SetIdentity();
                        matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                        matrix.Translate(matrix.Transform(new Vector(-shoeInputs.ShoeLength / 2 + xLocation[index], 0, 0)));
                        AddShield(clamp2Inputs, matrix, m_oSymbolic.Outputs, "clamp2" + index);
                    }
                }
                outerMostDia = GetGreaterValue(Val1: diameter1, Val2: shieldODia, Val3: shield2ODia, Val4: clampOdia, Val5: GetGreaterValue(repadODia, clamp2Odia, 0, 0, 0));
                outerMostRadius = (outerMostDia / 2) / (Math.Cos(shoeInputs.SlopeAngle));
                tempHeight = outerMostRadius + slideVertOffset + (guideInputs.Thickness1 - diameter1 / 2);
                shoeShapeHeight = shoeHeight - tempHeight;

                // Clamp widths 
                // When clamp quantity is 1 and its location by centre 
                if ((int)multi2Qty == 1 && (int)multi2LocateBy == 1)
                {
                    if (ClampWidthSel == 1)
                    {
                        clampInputs.Width1 = shoeInputs.ShoeLength + (2 * ClampLenOvrHng);
                        clampInputs.Width2 = shoeInputs.ShoeLength + (2 * ClampLenOvrHng);
                    }
                    else if (ClampWidthSel == 2)
                    {
                        if (ClampWidth1 > 0)
                        {
                            clampInputs.Width1 = ClampWidth1;
                        }
                        if (ClampWidth2 > 0)
                        {
                            clampInputs.Width2 = ClampWidth2;
                        }
                    }
                }
                else
                {
                    if (ClampWidth1 > 0)
                    {
                        clampInputs.Width1 = ClampWidth1;
                    }
                    if (ClampWidth2 > 0)
                    {
                        clampInputs.Width2 = ClampWidth2;
                    }
                }

                //Repad Lengths
                // When Repad quantity is 1 and its location by centre 
                if ((int)multi1Qty == 1 && (int)multi1LocateBy == 1)
                {
                    if (PadLengthSel == 1)
                    {
                        repadInputs.Length1 = shoeInputs.ShoeLength + (2 * PadLenOveHng);
                        repadInputs.Length2 = shoeInputs.ShoeLength + (2 * PadLenOveHng);
                        repadInputs.Length3 = shoeInputs.ShoeLength + (2 * PadLenOveHng);
                    }
                    else if (PadLengthSel == 2)
                    {
                        if (repadLength > 0)
                        {
                            repadInputs.Length1 = repadLength;
                            repadInputs.Length2 = repadLength;
                            repadInputs.Length3 = repadLength;
                        }
                    }
                    if (repadThickness > 0)
                    {
                        repadInputs.Thickness1 = repadThickness;
                    }
                }
                else
                {
                    if (repadLength > 0)
                    {
                        repadInputs.Length1 = repadLength;
                        repadInputs.Length2 = repadLength;
                        repadInputs.Length3 = repadLength;
                    }

                    if (repadThickness > 0)
                    {
                        repadInputs.Thickness1 = repadThickness;
                    }
                }

                //Shield Lengths
                // When Shield quantity is 1 and its location by centre 
                if ((int)multi7Qty == 1 && (int)multi7LocateBy == 1)
                {
                    if (ShieldLengthSel == 1)
                    {
                        shieldInputs.Length1 = shoeInputs.ShoeLength + (2 * ShieldLenOvrHng);
                        shieldInputs.Length2 = shoeInputs.ShoeLength + (2 * ShieldLenOvrHng);
                        shieldInputs.Length3 = shoeInputs.ShoeLength + (2 * ShieldLenOvrHng);
                    }
                    else if (ShieldLengthSel == 2)
                    {
                        if (repadLength > 0)
                        {
                            shieldInputs.Length1 = shieldLength;
                            shieldInputs.Length2 = shieldLength;
                            shieldInputs.Length3 = shieldLength;
                        }
                    }
                }
                else
                {
                    if (repadLength > 0)
                    {
                        shieldInputs.Length1 = shieldLength;
                        shieldInputs.Length2 = shieldLength;
                        shieldInputs.Length3 = shieldLength;
                    }
                }

                //Leg spacing
                if (LegSpace > 0)
                {
                    shoeInputs.ShoeSpacing = LegSpace;
                }

                if (LeglowSpace > 0)
                {
                    shoeInputs.LegLowerSpacing = LeglowSpace;
                }

                if (HgrCompareDoubleService.cmpdbl(shoeHeight, 0) == false)
                {
                    shoeInputs.ShoeHeight = shoeShapeHeight;
                    if (shoeShapeHeight < 0)
                        ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrInvalidHeight1NLTZero, "The shoe height is too small"));
                }

                StringBuilder Msg = new StringBuilder();
                SteelMember PipeMember;
                double dCylHeight = 0, diam;

                // Add Shoe shapes
                if (shoeName != "" && shoeName != "No Value")
                {
                    currentTotalHeight = outerMostRadius + shoeInputs.ShoeHeight;
                    offsettoShoeBottom = currentTotalHeight;
                    if (shoeInputs.ShoeType == 10 && (shoeInputs.SectionType.ToUpper() == "PIPE" || shoeInputs.SectionType.ToUpper() == "HSSC" || shoeInputs.SectionType.ToUpper() == "CS"))
                    {
                        PipeMember = GetSectionDataFromSection(shoeInputs.SectionStandard, shoeInputs.SectionType, shoeInputs.SectionName);
                        diam = PipeMember.depth;
                        shoeoff = Math.Sqrt(Math.Pow(((clampInputs.Diameter1 / 2) + clampInputs.Thickness1), 2) - Math.Pow((diam / 2), 2));
                        shoeInputs.ShoeHeight = shoeShapeHeight + tempHeight;
                        dCylHeight = shoeInputs.ShoeHeight - shoeInputs.ShoeThickness2;
                        SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                        Vector normal = new Vector(0, 0, 1);
                        Vector orthogonal = normal.GetOrthogonalVector();
                        symbolGeometryHelper.ActivePosition = new Position(0, 0, 0);
                        symbolGeometryHelper.SetOrientation(normal, orthogonal);
                        Projection3d shoe1 = symbolGeometryHelper.CreateCylinder(null, diam / 2, dCylHeight);
                        matrix = new Matrix4X4();
                        matrix.SetIdentity();
                        matrix.Rotate(Math.PI, new Vector(0, 1, 0), new Position(0, 0, 0));//180-y
                        shoe1.Transform(matrix);
                        matrix = new Matrix4X4();
                        matrix.SetIdentity();
                        matrix.Translate(new Vector(0, 0, -shoeoff));
                        shoe1.Transform(matrix);

                        m_oSymbolic.Outputs.Add("Shoe1", shoe1);

                        if (shoeInputs.ShoeThickness2 > 0)
                        {
                            symbolGeometryHelper = new SymbolGeometryHelper();
                            matrix = new Matrix4X4();
                            symbolGeometryHelper.ActivePosition = new Position(-shoeInputs.HorizontalPlateLength / 2, -shoeInputs.HorizontalPlateLength / 2, -(shoeoff + shoeInputs.ShoeHeight));
                            symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                            Projection3d endcap1 = (Projection3d)symbolGeometryHelper.CreateBox(null, shoeInputs.HorizontalPlateLength, shoeInputs.HorizontalPlateLength, shoeInputs.ShoeThickness2, 9);
                            m_oSymbolic.Outputs.Add("Endcap1", endcap1);
                        }

                        switch (int.Parse(shoeQty.ToString()))
                        {
                            case 2:
                                symbolGeometryHelper = new SymbolGeometryHelper();
                                symbolGeometryHelper.ActivePosition = new Position(0, 0, 0);
                                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), normal.GetOrthogonalVector());
                                Projection3d shoe2 = symbolGeometryHelper.CreateCylinder(null, diam / 2, dCylHeight);
                                matrix = new Matrix4X4();
                                matrix.SetIdentity();
                                matrix.Translate(new Vector(0, 0, shoeoff));
                                shoe2.Transform(matrix);
                                m_oSymbolic.Outputs.Add("Shoe2", shoe2);
                                //output to be defined
                                if (shoeInputs.ShoeThickness2 > 0)
                                {
                                    symbolGeometryHelper = new SymbolGeometryHelper();
                                    symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                                    Projection3d endcap2 = (Projection3d)symbolGeometryHelper.CreateBox(null, shoeInputs.HorizontalPlateLength, shoeInputs.HorizontalPlateLength, shoeInputs.ShoeThickness2, 9);
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Rotate(Math.PI, new Vector(0, 1, 0), new Position(0, 0, 0));//180-y
                                    endcap2.Transform(matrix);
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(shoeInputs.HorizontalPlateLength / 2, -shoeInputs.HorizontalPlateLength / 2, (shoeoff + shoeInputs.ShoeHeight)));
                                    endcap2.Transform(matrix);
                                    m_oSymbolic.Outputs.Add("Endcap2", endcap2);
                                }
                                break;
                        }//switch
                    }//if
                    else
                    {
                        SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                        matrix.SetIdentity();
                        matrix.Origin = new Position(dExtensionAdjOffset, 0, -currentTotalHeight);
                        AddShoe(shoeInputs, shoeGap, matrix, m_oSymbolic.Outputs, outerMostDia, "Shoe1", "shoe1");
                        if ((int)shoeQty == 2)
                        {
                            //AddShoe
                            matrix = new Matrix4X4();
                            matrix.SetIdentity();
                            matrix.Translate(new Vector(dExtensionAdjOffset, 0, -currentTotalHeight));
                            matrix.Rotate(Math.PI, new Vector(0, 1, 0), new Position(0, 0, 0));//180-y
                            AddShoe(shoeInputs, shoeGap, matrix, m_oSymbolic.Outputs, outerMostDia, "Shoe2", "shoe2");
                        }
                        else if (shoeQty == 3)
                        {
                            //Addshoe
                            matrix = new Matrix4X4();
                            matrix.SetIdentity();
                            matrix.Translate(new Vector(dExtensionAdjOffset, 0, -currentTotalHeight));
                            matrix.Rotate(Math.PI / 2, new Vector(0, 1, 0), new Position(0, 0, 0));//90-y
                            AddShoe(shoeInputs, shoeGap, matrix, m_oSymbolic.Outputs, outerMostDia, "Shoe2", "shoe2");
                            //Addshoe
                            matrix = new Matrix4X4();
                            matrix.SetIdentity();
                            matrix.Translate(new Vector(dExtensionAdjOffset, 0, -currentTotalHeight));
                            matrix.Rotate((Math.PI / 2) * 3, new Vector(0, 1, 0), new Position(0, 0, 0));//270-y
                            AddShoe(shoeInputs, shoeGap, matrix, m_oSymbolic.Outputs, outerMostDia, "Shoe3", "shoe3");
                        }
                        else if (shoeQty == 4)
                        {
                            //Addshoe
                            matrix = new Matrix4X4();
                            matrix.SetIdentity();
                            matrix.Translate(new Vector(dExtensionAdjOffset, 0, -currentTotalHeight));
                            matrix.Rotate(Math.PI, new Vector(0, 1, 0), new Position(0, 0, 0));//90-x
                            AddShoe(shoeInputs, shoeGap, matrix, m_oSymbolic.Outputs, outerMostDia, "Shoe2", "shoe2");
                            //Addshoe
                            matrix = new Matrix4X4();
                            matrix.SetIdentity();
                            matrix.Translate(new Vector(dExtensionAdjOffset, 0, -currentTotalHeight));
                            matrix.Rotate(Math.PI / 2, new Vector(0, 1, 0), new Position(0, 0, 0));//90-y
                            AddShoe(shoeInputs, shoeGap, matrix, m_oSymbolic.Outputs, outerMostDia, "Shoe3", "shoe3");
                            //Addshoe
                            matrix = new Matrix4X4();
                            matrix.SetIdentity();
                            matrix.Translate(new Vector(dExtensionAdjOffset, 0, -currentTotalHeight));
                            matrix.Rotate((Math.PI / 2) * 3, new Vector(0, 1, 0), new Position(0, 0, 0));//270-y
                            AddShoe(shoeInputs, shoeGap, matrix, m_oSymbolic.Outputs, outerMostDia, "Shoe4", "shoe4");
                        }
                    }//if-else
                }//if  

                // Add Slide Plate
                if (slidePlateName != "" && slidePlateName != "No Value")
                {
                    double zLoc;
                    currentTotalHeight = currentTotalHeight + slideVertOffset;
                    xLocation = new double[int.Parse(multi9Qty.ToString())];
                    for (int index = 0; index < multi9Qty; index++)
                    {
                        xLocation[index] = ShieldMultiPosition(shoeInputs.ShoeLength, multi9Qty, multi9LocateBy, multi9Location, slidePlateInputs.Height1)[index];
                        if (shoeInputs.ShoeType == 10 && shoeInputs.SectionType.ToUpper() == "PIPE")
                            zLoc = shoeoff + shoeInputs.ShoeHeight + shoeInputs.ShoeThickness2 + slidePlateInputs.Thickness1 + slidePlateInputs.Thickness2 + slidePlateInputs.Thickness3;
                        else
                            zLoc = currentTotalHeight;
                        matrix = new Matrix4X4();
                        matrix.SetIdentity();
                        matrix.Translate(new Vector(dExtensionAdjOffset, -shoeInputs.ShoeLength / 2 + xLocation[index], -zLoc));
                        AddSlidePlate(slidePlateInputs, matrix, m_oSymbolic.Outputs, "Shoe-SlidePlate" + index);

                    }//for end
                }//if

                // Add Guide
                if (guideName != "" && guideName != "No Value")
                {
                    guideLength = guideInputs.Length1;                    
                    guideInputs.Gap1 = guideInputs.Gap1 + (2 * guideGap);
                    guideInputs.Width1 = guideInputs.Width1 + (2 * guideGap) + ((guideInputs.Width2) * 2);

                    if (guideLength < guideInputs.Length2)
                        guideLength = guideInputs.Length2 + (2 * guideGap);
                    currentTotalHeight = currentTotalHeight + guideInputs.Thickness1;
                    xLocation = new double[int.Parse(multi6Qty.ToString())];
                    for (int index = 0; index < multi6Qty; index++)
                    {
                        xLocation[index] = ShieldMultiPosition(shoeInputs.ShoeLength, multi6Qty, multi6LocateBy, multi6Location, guideLength)[index];

                        matrix = new Matrix4X4();
                        matrix.SetIdentity();
                        matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                        matrix.Translate(new Vector(dExtensionAdjOffset, -shoeInputs.ShoeLength / 2 + xLocation[index], -currentTotalHeight));
                        AddGuide(guideInputs, matrix, m_oSymbolic.Outputs, "Shoe-Guide-" + index);
                    }//for end
                }

                // Add End plates
                if (endPlateName != "" && endPlateName != "No Value")
                {
                    //Calculating the values for the Radialcut
                    double PlateDiaRatio = ((endPlateInputs.width1 / outerMostDia));

                    if (PlateDiaRatio > 1)
                    {
                        if (!(endPlateInputs.length1 > 0))
                            endPlateInputs.length1 = (shoeInputs.ShoeHeight - shoeInputs.ShoeThickness2 + clampInputs.Thickness1) + outerMostDia / 2;
                        if (endPlateInputs.length1 - (shoeInputs.ShoeHeight - shoeInputs.ShoeThickness2) > 0)
                        {
                            endPlateInputs.curvedEndY = outerMostDia / 2 - (endPlateInputs.length1 - (shoeInputs.ShoeHeight - shoeInputs.ShoeThickness2 + clampInputs.Thickness1));
                            endPlateInputs.curvedEndRad = (outerMostDia / 2);
                        }
                    }
                    else
                    {
                        if (!(endPlateInputs.length1 > 0))
                            endPlateInputs.length1 = (((shoeInputs.ShoeHeight - shoeInputs.ShoeThickness2) + (outerMostDia / 2) - (Math.Sin(Math.Acos(endPlateInputs.width1 / outerMostDia)) * (outerMostDia / 2))));
                        if (endPlateInputs.length1 - (shoeInputs.ShoeHeight - shoeInputs.ShoeThickness2) > 0)
                        {
                            endPlateInputs.curvedEndY = (outerMostDia / 2 - (endPlateInputs.length1 - (shoeInputs.ShoeHeight - shoeInputs.ShoeThickness2)));
                            endPlateInputs.curvedEndRad = (outerMostDia / 2);
                        }
                    }
                    //Bottom Plates
                    //First End Plate dExtensionAdjOffset+ clampInputs.Thickness1
                    matrix = new Matrix4X4();
                    matrix.SetIdentity();
                    matrix.Translate(new Vector(-(endPlateInputs.width1 / 2) + dExtensionAdjOffset, -currentTotalHeight + shoeInputs.ShoeThickness2 + guideInputs.Thickness1 + slideVertOffset, -(shoeInputs.VerticalPlateLength / 2) + endPlateOffset));
                    matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0), new Position(0, 0, 0));//90-x
                    AddPlate(endPlateInputs, matrix, m_oSymbolic.Outputs, "Shoe-EndPlate-1");
                    ////Second End Plate
                    if (endPlateNum > 1)
                    {
                        matrix = new Matrix4X4();
                        matrix.SetIdentity();//-           
                        matrix.Translate(new Vector(-(endPlateInputs.width1 / 2) + dExtensionAdjOffset, -currentTotalHeight + shoeInputs.ShoeThickness2 + guideInputs.Thickness1 + slideVertOffset, (shoeInputs.VerticalPlateLength / 2) - endPlateInputs.thickness1 - endPlateOffset));
                        matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0), new Position(0, 0, 0));//90-x
                        AddPlate(endPlateInputs, matrix, m_oSymbolic.Outputs, "Shoe-EndPlate-2");
                    }
                    //Top Plate
                    //Third End Plate
                    if (shoeQty == 2)
                    {
                        matrix = new Matrix4X4();
                        matrix.SetIdentity();
                        matrix.Translate(new Vector(-(endPlateInputs.width1 / 2) + dExtensionAdjOffset, -currentTotalHeight + shoeInputs.ShoeThickness2 + guideInputs.Thickness1 + slideVertOffset, -(shoeInputs.VerticalPlateLength / 2) + endPlateOffset));
                        matrix.Rotate(Math.PI / 2 * 3, new Vector(1, 0, 0), new Position(0, 0, 0));//270-x
                        AddPlate(endPlateInputs, matrix, m_oSymbolic.Outputs, "Shoe-EndPlate-3");
                        if (endPlateNum > 1)
                        {
                            //fourth End Plate
                            matrix = new Matrix4X4();
                            matrix.SetIdentity();
                            matrix.Translate(new Vector(-endPlateInputs.width1 / 2 + dExtensionAdjOffset, -currentTotalHeight + shoeInputs.ShoeThickness2 + guideInputs.Thickness1 + slideVertOffset, (shoeInputs.VerticalPlateLength / 2) - endPlateInputs.thickness1 - endPlateOffset));
                            matrix.Rotate(Math.PI / 2 * 3, new Vector(1, 0, 0), new Position(0, 0, 0));//270-x
                            AddPlate(endPlateInputs, matrix, m_oSymbolic.Outputs, "Shoe-EndPlate-4");
                        }
                    }
                    else if (shoeQty == 3)
                    {
                        //Left Plates
                        //Third End Plates
                        matrix = new Matrix4X4();
                        matrix.SetIdentity();
                        matrix.Translate(new Vector(-(endPlateInputs.width1 / 2) + dExtensionAdjOffset, -currentTotalHeight + shoeInputs.ShoeThickness2 + guideInputs.Thickness1 + slideVertOffset, -(shoeInputs.VerticalPlateLength / 2) + endPlateOffset));
                        matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0), new Position(0, 0, 0));//90-x,90-y
                        matrix.Rotate(Math.PI / 2, new Vector(0, 1, 0), new Position(0, 0, 0));
                        AddPlate(endPlateInputs, matrix, m_oSymbolic.Outputs, "Shoe-EndPlate-3");
                        if (endPlateNum > 1)
                        {
                            //fourth End Plate
                            matrix = new Matrix4X4();
                            matrix.SetIdentity();
                            matrix.Translate(new Vector(-endPlateInputs.width1 / 2 + dExtensionAdjOffset, -currentTotalHeight + shoeInputs.ShoeThickness2 + guideInputs.Thickness1 + slideVertOffset, (shoeInputs.VerticalPlateLength / 2) - endPlateInputs.thickness1 - endPlateOffset));
                            matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0), new Position(0, 0, 0));//90-x//90-y
                            matrix.Rotate(Math.PI / 2, new Vector(0, 1, 0), new Position(0, 0, 0));
                            AddPlate(endPlateInputs, matrix, m_oSymbolic.Outputs, "Shoe-EndPlate-4");
                        }
                        //Rigth Plates
                        //Fifth End Plates
                        matrix = new Matrix4X4();
                        matrix.SetIdentity();
                        matrix.Translate(new Vector(-endPlateInputs.width1 / 2 + dExtensionAdjOffset, -currentTotalHeight + shoeInputs.ShoeThickness2 + guideInputs.Thickness1 + slideVertOffset, -(shoeInputs.VerticalPlateLength / 2) + endPlateOffset));
                        matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0), new Position(0, 0, 0));//90-x//270-y
                        matrix.Rotate(Math.PI / 2 * 3, new Vector(0, 1, 0), new Position(0, 0, 0));
                        AddPlate(endPlateInputs, matrix, m_oSymbolic.Outputs, "Shoe-EndPlate-5");

                        if (endPlateNum > 1)
                        {
                            //Sixth End Plate
                            matrix = new Matrix4X4();
                            matrix.SetIdentity();
                            matrix.Translate(new Vector(-endPlateInputs.width1 / 2 + dExtensionAdjOffset, -currentTotalHeight + shoeInputs.ShoeThickness2 + guideInputs.Thickness1 + slideVertOffset, (shoeInputs.VerticalPlateLength / 2) - endPlateInputs.thickness1 - endPlateOffset));
                            matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0), new Position(0, 0, 0));//90-x//270-y
                            matrix.Rotate((Math.PI / 2) * 3, new Vector(0, 1, 0), new Position(0, 0, 0));//90-x//270-y
                            AddPlate(endPlateInputs, matrix, m_oSymbolic.Outputs, "Shoe-EndPlate-6");
                        }

                    }
                    else if (shoeQty == 4)
                    {
                        //Top Plates
                        //Third End Plates
                        matrix = new Matrix4X4();
                        matrix.SetIdentity();
                        matrix.Rotate(Math.PI / 2 * 3, new Vector(1, 0, 0), new Position(0, 0, 0));//270-x      
                        matrix.Translate(new Vector(-(endPlateInputs.width1 / 2) + dExtensionAdjOffset, -currentTotalHeight + shoeInputs.ShoeThickness2 + guideInputs.Thickness1 + slideVertOffset, -(shoeInputs.VerticalPlateLength / 2) + endPlateOffset));
                        AddPlate(endPlateInputs, matrix, m_oSymbolic.Outputs, "Shoe-EndPlate-3");
                        // Fourth End Plate
                        if (endPlateNum > 1)
                        {
                            matrix = new Matrix4X4();
                            matrix.SetIdentity();
                            matrix.Rotate(Math.PI / 2 * 3, new Vector(1, 0, 0), new Position(0, 0, 0));//270-x     
                            matrix.Translate(new Vector(-endPlateInputs.width1 / 2 + dExtensionAdjOffset, -currentTotalHeight + shoeInputs.ShoeThickness2 + guideInputs.Thickness1 + slideVertOffset, (shoeInputs.VerticalPlateLength / 2) - endPlateInputs.thickness1 - endPlateOffset));
                            AddPlate(endPlateInputs, matrix, m_oSymbolic.Outputs, "Shoe-EndPlate-4");
                        }
                        //Left Plates
                        //Fifth End Plates
                        matrix = new Matrix4X4();
                        matrix.SetIdentity();
                        matrix.Translate(new Vector(-endPlateInputs.width1 / 2 + dExtensionAdjOffset, -currentTotalHeight + shoeInputs.ShoeThickness2 + guideInputs.Thickness1 + slideVertOffset, -(shoeInputs.VerticalPlateLength / 2) + endPlateOffset));
                        matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0), new Position(0, 0, 0));//90-x//90-y
                        matrix.Rotate(Math.PI / 2, new Vector(0, 1, 0), new Position(0, 0, 0));
                        AddPlate(endPlateInputs, matrix, m_oSymbolic.Outputs, "Shoe-EndPlate-5");
                        // 'Sixth End Plate
                        if (endPlateNum > 1)
                        {
                            matrix = new Matrix4X4();
                            matrix.SetIdentity();
                            matrix.Translate(new Vector(-endPlateInputs.width1 / 2 + dExtensionAdjOffset, -currentTotalHeight + shoeInputs.ShoeThickness2 + guideInputs.Thickness1 + slideVertOffset, (shoeInputs.VerticalPlateLength / 2) - endPlateInputs.thickness1 - endPlateOffset));
                            matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0), new Position(0, 0, 0));//90-x//90-y
                            matrix.Rotate(Math.PI / 2, new Vector(0, 1, 0), new Position(0, 0, 0));//90-x//90-y
                            AddPlate(endPlateInputs, matrix, m_oSymbolic.Outputs, "Shoe-EndPlate-6");
                        }
                        //Rigth Plate
                        //seventh End plate
                        matrix = new Matrix4X4();
                        matrix.SetIdentity();
                        matrix.Translate(new Vector(-endPlateInputs.width1 / 2 + dExtensionAdjOffset, -currentTotalHeight + shoeInputs.ShoeThickness2 + guideInputs.Thickness1 + slideVertOffset, -(shoeInputs.VerticalPlateLength / 2) + endPlateOffset));
                        matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0), new Position(0, 0, 0));//90-x//270-y
                        matrix.Rotate((Math.PI / 2) * 3, new Vector(0, 1, 0), new Position(0, 0, 0));//90-x//270-y
                        AddPlate(endPlateInputs, matrix, m_oSymbolic.Outputs, "Shoe-EndPlate-7");
                        if (endPlateNum > 1)
                        {
                            // Eighth End Plate
                            matrix = new Matrix4X4();
                            matrix.SetIdentity();
                            matrix.Translate(new Vector(-endPlateInputs.width1 / 2 + dExtensionAdjOffset, -currentTotalHeight + shoeInputs.ShoeThickness2 + guideInputs.Thickness1 + slideVertOffset, (shoeInputs.VerticalPlateLength / 2) - endPlateInputs.thickness1 - endPlateOffset));
                            matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0), new Position(0, 0, 0));//90-x//270-y
                            matrix.Rotate((Math.PI / 2) * 3, new Vector(0, 1, 0), new Position(0, 0, 0));
                            AddPlate(endPlateInputs, matrix, m_oSymbolic.Outputs, "Shoe-EndPlate-8");
                        }
                    }
                }

                //shield
                shieldSlopeOffset = -(Math.Sin(shoeInputs.SlopeAngle) * ((outerMostDia / 2) + shoeInputs.TopPlateGap + shoeInputs.TopPlateOffset));
                if (shieldName != "" && shieldName != "No Value")
                {

                    xLocation = new double[int.Parse(multi7Qty.ToString())];
                    for (int index = 0; index < multi7Qty; index++)
                    {
                        xLocation[index] = ShieldMultiPosition(shoeInputs.ShoeLength, multi7Qty, multi7LocateBy, multi7Location, shieldInputs.Length1)[index];

                        matrix = new Matrix4X4();
                        matrix.SetIdentity();
                        matrix.Translate(new Vector(-shoeInputs.ShoeLength / 2 + xLocation[index] + shieldOffset1 + shieldSlopeOffset, 0, 0));
                        matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                        matrix.Rotate(0, new Vector(1, 0, 0));
                        AddShield(shieldInputs, matrix, m_oSymbolic.Outputs, "Shoe-Shield" + index);
                    }//for end
                }

                //SHIELD 2/Vapor barrier
                if (shield2Name != "" && shield2Name != "No Value")
                {
                    shield2SlopeOffset = -(Math.Sin(shoeInputs.SlopeAngle) * ((outerMostDia / 2) + shoeInputs.TopPlateGap + shoeInputs.TopPlateOffset));
                    xLocation = new double[int.Parse(multi10Qty.ToString())];
                    for (int index = 0; index < multi10Qty; index++)
                    {
                        xLocation[index] = ShieldMultiPosition(shoeInputs.ShoeLength, multi10Qty, multi10LocateBy, multi10Location, shieldInputs.Length1)[index];

                        matrix = new Matrix4X4();
                        matrix.SetIdentity();
                        matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                        matrix.Rotate(shoeInputs.SlopeAngle , new Vector(1, 0, 0), new Position(0, 0, 0));
                        matrix.Translate(new Vector(-shoeInputs.ShoeLength / 2 + xLocation[index] + shieldOffset1 + shieldSlopeOffset, 0, 0));
                        // addshield 
                        AddShield(shield2Inputs, matrix, m_oSymbolic.Outputs, "Shoe-Shield2" + index);
                    }//for end
                }

                //Add Bolt
                if (uboltName != "" && uboltName != "No Value")
                {
                    uboltThichness = uBoltInputs.UBoltRodDia;
                    uboltSlopeOffset = -(Math.Sin(shoeInputs.SlopeAngle) * ((outerMostDia / 2) + shoeInputs.TopPlateOffset + shoeInputs.TopPlateOffset));
                    if (uBoltInputs.UBoltRodDia > uboltThichness)
                        uboltThichness = uBoltInputs.UBoltDia2;
                    xLocation = new double[int.Parse(multi3Qty.ToString())];
                    for (int index = 0; index < multi3Qty; index++)
                    {
                        xLocation[index] = ShieldMultiPosition(shoeInputs.ShoeLength, multi3Qty, multi3LocateBy, multi3Location, uboltThichness)[index];
                        matrix = new Matrix4X4();
                        matrix.SetIdentity();
                        matrix.Translate(new Vector(-shoeInputs.ShoeLength / 2 + xLocation[index] + uboltSlopeOffset, 0, 0));
                        matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                        matrix.Rotate(shoeInputs.SlopeAngle, new Vector(1, 0, 0), new Position(0, 0, 0));
                        // AddUbolt 
                        AddUBolt(uBoltInputs, diameter1, matrix, m_oSymbolic.Outputs, "Shoe-U-Bolt" + index);
                    }//for end
                }

                //Add Strap
                if (strapName != "" && strapName != "No Value")
                {
                    strapThickness = strapInputs.StrapStockWidth;
                    strapSlopeOffset = -(Math.Sin(shoeInputs.SlopeAngle) * ((outerMostDia / 2) + shoeInputs.TopPlateGap + shoeInputs.TopPlateOffset));
                    xLocation = new double[int.Parse(multi4Qty.ToString())];
                    for (int index = 0; index < multi4Qty; index++)
                    {
                        xLocation[index] = ShieldMultiPosition(shoeInputs.ShoeLength, multi4Qty, multi4LocateBy, multi4Location, strapThickness)[index];
                        matrix = new Matrix4X4();
                        matrix.SetIdentity();
                        matrix.Translate(new Vector(-shoeInputs.ShoeLength / 2 + xLocation[index] + shieldSlopeOffset, 0, 0));
                        matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                        matrix.Rotate(shoeInputs.SlopeAngle , new Vector(1, 0, 0), new Position(0, 0, 0));
                        //Addstrap 
                        AddStrap(strapInputs, diameter1, matrix, m_oSymbolic.Outputs, "Shoe-Strap" + index);
                    }//for end
                }

                // Add Repad
                if (repadName != "" && repadName != "No Value")
                {
                    repadSlopeOffset = -(Math.Sin(shoeInputs.SlopeAngle) * ((outerMostDia / 2) + shoeInputs.TopPlateGap + shoeInputs.TopPlateOffset));
                    xLocation = new double[int.Parse(multi1Qty.ToString())];
                    for (int index = 0; index < multi1Qty; index++)
                    {
                        xLocation[index] = ShieldMultiPosition(shoeInputs.ShoeLength, multi1Qty, multi1LocateBy, multi1Location, repadInputs.Length1)[index];
                        matrix = new Matrix4X4();
                        matrix.SetIdentity();
                        matrix.Translate(new Vector(-shoeInputs.ShoeLength / 2 + xLocation[index] + repadOffset1 + repadSlopeOffset, 0, repadInputs.PipeOD / 2 - diameter1 / 2 - shoeGap));
                        matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                        matrix.Rotate(shoeInputs.SlopeAngle , new Vector(1, 0, 0), new Position(0, 0, 0));
                        //AddShield 
                        AddShield(repadInputs, matrix, m_oSymbolic.Outputs, "Shoe-Repad" + index);
                    }//for end
                }

                // Add Clamp
                if (clampName != "" && clampName != "No Value")
                {
                    clampSlopeOffset = -(Math.Sin(shoeInputs.SlopeAngle) * ((outerMostDia / 2) + shoeInputs.TopPlateGap + shoeInputs.TopPlateOffset));
                    xLocation = new double[int.Parse(multi2Qty.ToString())];
                    for (int index = 0; index < multi2Qty; index++)
                    {
                        xLocation[index] = ShieldMultiPosition(shoeInputs.ShoeLength, multi2Qty, multi2LocateBy, multi2Location, clampThickness)[index];
                        matrix = new Matrix4X4();
                        matrix.SetIdentity();
                        matrix.Rotate(0, new Vector(1, 0, 0));
                        matrix.Translate(new Vector(-shoeInputs.ShoeLength / 2 + xLocation[index] + ClampOffset1 + clampSlopeOffset, 0, 0));
                        matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                        matrix.Rotate(shoeInputs.SlopeAngle , new Vector(1, 0,0), new Position(0, 0, 0));
                        //Addpipeclamp 
                        AddPipeClamp(clampInputs, matrix, m_oSymbolic.Outputs, "Shoe-Clamp" + index);

                    }//for end
                }

                //Add Additional Plates
                if (additionalPlateName != "" && additionalPlateName != "No Value")
                {
                    //Add platemethod
                    matrix = new Matrix4X4();
                    matrix.Translate(new Vector(0, addPlateOffset1, 0));
                    matrix.Rotate(plateRotationAngle, new Vector(0, 1, 0), new Position(0, 0, 0));
                    StarPlateShape(additionalPlateInputs, m_oSymbolic.Outputs, matrix, centerDiameter, int.Parse(plateNumber.ToString()), plateAngle, "shoeAddPlates");
                }

                //Add Gusset
                if (gussetName != "" && gussetName != "No Value")
                {
                    xLocation = new double[int.Parse(multi5Qty.ToString())];
                    for (int index = 0; index < multi5Qty; index++)
                    {
                        x1Offset = -gussetInputs.width1 - (gussetGap / 2) + dExtensionAdjOffset;
                        x2Offset = gussetInputs.width1 + (gussetGap / 2) + dExtensionAdjOffset;

                        xLocation[index] = ShieldMultiPosition(shoeInputs.ShoeLength, multi5Qty, multi5LocateBy, multi5Location, gussetInputs.thickness1)[index];
                        //Add Plate
                        matrix = new Matrix4X4();
                        matrix.Translate(new Vector(x1Offset, -currentTotalHeight + shoeInputs.ShoeThickness2 + guideInputs.Thickness1 + slideVertOffset, -(gussetInputs.thickness1 / 2) - (shoeInputs.ShoeLength / 2) + xLocation[index]));
                        matrix.Rotate((Math.PI / 2), new Vector(1, 0, 0), new Position(0, 0, 0));
                        AddPlate(gussetInputs, matrix, m_oSymbolic.Outputs, "Shoe-Gusset1-" + index);
                        //Add Plate
                        matrix = new Matrix4X4();
                        matrix.Rotate((Math.PI), new Vector(0, 1, 0), new Position(0, 0, 0));
                        matrix.Translate(new Vector(x2Offset, -currentTotalHeight + shoeInputs.ShoeThickness2 + guideInputs.Thickness1 + slideVertOffset, (gussetInputs.thickness1 / 2) - (shoeInputs.ShoeLength / 2) + xLocation[index]));
                        matrix.Rotate((Math.PI / 2), new Vector(1, 0, 0), new Position(0, 0, 0));
                        AddPlate(gussetInputs, matrix, m_oSymbolic.Outputs, "Shoe-Gusset2-" + index);
                    }//for end
                }

                //RIB Plates
                if (ribName != "" && ribName != "No Value")
                {
                    if (!(ribInputs.thickness1 > 0))
                        ribInputs.thickness1 = shoeInputs.ShoeThickness1;
                    if (!(ribInputs.width1 > 0))
                        ribInputs.width1 = shoeInputs.VerticalPlateLength;
                    if (!(ribInputs.length1 > 0))
                        ribInputs.length1 = (LegHeightToPipe(ribInputs.thickness1 + ribGap, outerMostDia, shoeInputs.ShoeHeight, 0, shoeGap)) - shoeInputs.ShoeThickness2;
                    upwardRibOffset = -currentTotalHeight + shoeInputs.ShoeThickness2 + guideInputs.Thickness1 + slideVertOffset;
                    matrix = new Matrix4X4();

                    double[] dXLocation = new double[Convert.ToInt16(ribQty)];

                    for (int index = 0; index <= ribQty - 1; index++)
                    {
                        dXLocation[index] = MultiPosition(shoeInputs.ShoeWidth, ribQty, 1, ribGap, ribInputs.thickness1)[index];
                        matrix.SetIdentity();
                        matrix.Rotate(Math.PI / 2, new Vector(0, 1, 0), new Position(0, 0, 0));
                        matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0), new Position(0, 0, 0));
                        matrix.Translate(new Vector((-ribInputs.thickness1 / 2) - (shoeInputs.ShoeWidth / 2) + dXLocation[index], -ribInputs.width1 / 2, upwardRibOffset));
                        AddPlate(ribInputs, matrix, m_oSymbolic.Outputs, "Shoe-Rib1" + index);
                    }

                    if (shoeQty == 2)
                    {
                        for (int index = 0; index <= ribQty - 1; index++)
                        {
                            dXLocation[index] = MultiPosition(shoeInputs.ShoeWidth, ribQty, 1, ribGap, ribInputs.thickness1)[index];
                            matrix.SetIdentity();
                            matrix.Rotate(Math.PI / 2, new Vector(0, 1, 0), new Position(0, 0, 0));
                            matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0), new Position(0, 0, 0));
                            matrix.Translate(new Vector((-ribInputs.thickness1 / 2) - (shoeInputs.ShoeWidth / 2) + dXLocation[index], -ribInputs.width1 / 2, upwardRibOffset));
                            matrix.Rotate(Math.PI, new Vector(0, 1, 0), new Position(0, 0, 0));
                            AddPlate(ribInputs, matrix, m_oSymbolic.Outputs, "Shoe-Rib2" + index);
                        }
                    }
                    else if (shoeQty == 3)
                    {
                        for (int index = 0; index <= ribQty - 1; index++)
                        {
                            dXLocation[index] = MultiPosition(shoeInputs.ShoeWidth, ribQty, 1, ribGap, ribInputs.thickness1)[index];
                            matrix.SetIdentity();
                            matrix.Rotate(Math.PI / 2, new Vector(0, 1, 0), new Position(0, 0, 0));
                            matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0), new Position(0, 0, 0));
                            matrix.Translate(new Vector((-ribInputs.thickness1 / 2) - (shoeInputs.ShoeWidth / 2) + dXLocation[index], -ribInputs.width1 / 2, upwardRibOffset));
                            matrix.Rotate(3 * Math.PI / 2, new Vector(0, 1, 0), new Position(0, 0, 0));
                            AddPlate(ribInputs, matrix, m_oSymbolic.Outputs, "Shoe-Rib2" + index);
                        }
                        for (int index = 0; index <= ribQty - 1; index++)
                        {
                            dXLocation[index] = MultiPosition(shoeInputs.ShoeWidth, ribQty, 1, ribGap, ribInputs.thickness1)[index];
                            matrix.SetIdentity();
                            matrix.Rotate(Math.PI / 2, new Vector(0, 1, 0), new Position(0, 0, 0));
                            matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0), new Position(0, 0, 0));
                            matrix.Translate(new Vector((-ribInputs.thickness1 / 2) - (shoeInputs.ShoeWidth / 2) + dXLocation[index], -ribInputs.width1 / 2, upwardRibOffset));
                            matrix.Rotate(Math.PI / 2, new Vector(0, 1, 0), new Position(0, 0, 0));
                            AddPlate(ribInputs, matrix, m_oSymbolic.Outputs, "Shoe-Rib3" + index);
                        }
                    }
                    else if (shoeQty == 4)
                    {
                        for (int index = 0; index <= ribQty - 1; index++)
                        {
                            dXLocation[index] = MultiPosition(shoeInputs.ShoeWidth, ribQty, 1, ribGap, ribInputs.thickness1)[index];
                            matrix.SetIdentity();
                            matrix.Rotate(Math.PI / 2, new Vector(0, 1, 0), new Position(0, 0, 0));
                            matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0), new Position(0, 0, 0));
                            matrix.Translate(new Vector((-ribInputs.thickness1 / 2) - (shoeInputs.ShoeWidth / 2) + dXLocation[index], -ribInputs.width1 / 2, upwardRibOffset));
                            matrix.Rotate(Math.PI, new Vector(0, 1, 0), new Position(0, 0, 0));
                            AddPlate(ribInputs, matrix, m_oSymbolic.Outputs, "Shoe-Rib2" + index);
                        }
                        for (int index = 0; index <= ribQty - 1; index++)
                        {
                            dXLocation[index] = MultiPosition(shoeInputs.ShoeWidth, ribQty, 1, ribGap, ribInputs.thickness1)[index];
                            matrix.SetIdentity();
                            matrix.Rotate(Math.PI / 2, new Vector(0, 1, 0), new Position(0, 0, 0));
                            matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0), new Position(0, 0, 0));
                            matrix.Translate(new Vector((-ribInputs.thickness1 / 2) - (shoeInputs.ShoeWidth / 2) + dXLocation[index], -ribInputs.width1 / 2, upwardRibOffset));
                            matrix.Rotate(3 * Math.PI / 2, new Vector(0, 1, 0), new Position(0, 0, 0));
                            AddPlate(ribInputs, matrix, m_oSymbolic.Outputs, "Shoe-Rib3" + index);
                        }
                        for (int index = 0; index <= ribQty - 1; index++)
                        {
                            dXLocation[index] = MultiPosition(shoeInputs.ShoeWidth, ribQty, 1, ribGap, ribInputs.thickness1)[index];
                            matrix.SetIdentity();
                            matrix.Rotate(Math.PI / 2, new Vector(0, 1, 0), new Position(0, 0, 0));
                            matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0), new Position(0, 0, 0));
                            matrix.Translate(new Vector((-ribInputs.thickness1 / 2) - (shoeInputs.ShoeWidth / 2) + dXLocation[index], -ribInputs.width1 / 2, upwardRibOffset));
                            matrix.Rotate(Math.PI / 2, new Vector(0, 1, 0), new Position(0, 0, 0));
                            AddPlate(ribInputs, matrix, m_oSymbolic.Outputs, "Shoe-Rib4" + index);
                        }
                    }
                }

                //Stiffener
                if (stiffenerName != "" && stiffenerName != "No Value")
                {
                    //Calculations for the curved end cuts for stiffener plates

                    if (!(stiffenerInputs.length1 > 0))
                        stiffenerInputs.length1 = (shoeInputs.ShoeHeight - stiffenerVerticalOffset) + (outerMostDia / 2) - (Math.Sin(Math.Acos(stiffenerInputs.width1 / outerMostDia)) * (outerMostDia / 2));

                    if ((stiffenerInputs.length1 + stiffenerVerticalOffset) - (shoeInputs.ShoeHeight - shoeInputs.ShoeThickness2) > 0)
                    {
                        stiffenerInputs.curvedEndY = (outerMostDia / 2) - ((stiffenerInputs.length1 + stiffenerVerticalOffset) - (shoeInputs.ShoeHeight));
                        stiffenerInputs.curvedEndRad = (outerMostDia / 2);
                    }

                    xLocation = new double[int.Parse(multi8Qty.ToString())];
                    for (int index = 0; index < multi8Qty; index++)
                    {

                        xLocation[index] = ShieldMultiPosition(shoeInputs.ShoeLength, multi8Qty, multi8LocateBy, multi8Location, stiffenerInputs.thickness1)[index];
                        //Add PlatShape
                        matrix = new Matrix4X4();
                        matrix.SetIdentity();
                        matrix.Translate(new Vector(-stiffenerInputs.width1 / 2 + dExtensionAdjOffset, -currentTotalHeight + stiffenerVerticalOffset  + guideInputs.Thickness1 + slideVertOffset, -(stiffenerInputs.thickness1 / 2) - (shoeInputs.ShoeLength / 2) + xLocation[index]));
                        matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0), new Position(0, 0, 0));
                        AddPlate(stiffenerInputs, matrix, m_oSymbolic.Outputs, "Shoe-Stiffener" + index);
                    }//for end
                }

                //Add Port
                Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(0, Math.Cos(shoeInputs.SlopeAngle), Math.Sin(shoeInputs.SlopeAngle)), new Vector(0, Math.Sin(shoeInputs.SlopeAngle), -Math.Cos(shoeInputs.SlopeAngle)));
                m_oSymbolic.Outputs["Route"] = port1;
                Port port2 = new Port(OccurrenceConnection, part, "Steel", new Position(0, 0, -currentTotalHeight), new Vector(0, 1, 0), new Vector(0, 0, 1));
                m_oSymbolic.Outputs["Steel"] = port2;
                if (shoeName != "" && shoeName != "No Value")
                {
                    double weldPortHeight;
                    double weldYOffset;
                    weldYOffset = GetLegWidthOnTop(shoeInputs) / 2;
                    if ((shoeInputs.ShoeType == 10) && (shoeInputs.SectionType.ToUpper() == "PIPE" || shoeInputs.SectionType.ToUpper() == "HSSC" || shoeInputs.SectionType.ToUpper() == "CS"))
                        weldPortHeight = -((diameter1 / 2) + clampInputs.Thickness1);
                    else
                        weldPortHeight = -offsettoShoeBottom + (LegHeightToPipe((weldYOffset + weldYOffset), outerMostDia, shoeInputs.ShoeHeight, shoeInputs.LegHeight, shoeGap));
                    try
                    {
                        Port port3 = new Port(OccurrenceConnection, part, "Weld", new Position(weldYOffset, 0, weldPortHeight), new Vector(0, 1, 0), new Vector(0, 0, 1));
                        m_oSymbolic.Outputs["Weld"] = port3;
                    }
                    catch
                    {
                        // the port may not be defined in refdata.
                    }
                }
                 // Construction of Maintenance Aspect


                if (createMaintenanceAspect)
                {
                    //=============================='
                    //'Load The Shape Data Structures'
                    //'=============================='
                    //'Load Shoe Shape data
                    shoeInputs = LoadShoeDataByQuery(shoeName);

                    //Load PipeClamp Data by query
                    clampInputs = LoadPipeClampDataByQuery(clampName);
                    // hs_Parts_LoadPipeClampDataByQuery tClamp, sClamp

                    if (HgrCompareDoubleService.cmpdbl(clampAngle, 0) == false)
                        clampInputs.Angle3 = clampAngle;
                    if (insulationThickness > 0)
                    {
                        if (clampInputs.Diameter1 > 0)
                            clampInputs.Diameter1 = insulationThickness * 2 + clampInputs.Diameter1;
                        if (clampInputs.Diameter4 > 0)
                            clampInputs.Diameter4 = insulationThickness * 2 + clampInputs.Diameter4;
                        clampInputs.Height1 = clampInputs.Height1 + insulationThickness;
                        clampInputs.Height2 = clampInputs.Height2 + insulationThickness;

                        clampInputs.BoltRow1.Offset = clampInputs.BoltRow1.Offset + insulationThickness;
                        clampInputs.BoltRow2.Offset = clampInputs.BoltRow2.Offset + insulationThickness;
                        clampInputs.BoltRow3.Offset = clampInputs.BoltRow3.Offset + insulationThickness;
                        clampInputs.BoltRow4.Offset = clampInputs.BoltRow4.Offset + insulationThickness;
                        clampInputs.BoltRow5.Offset = clampInputs.BoltRow5.Offset + insulationThickness;
                        clampInputs.BoltRow6.Offset = clampInputs.BoltRow6.Offset + insulationThickness;

                        if (insulationLength > 0)
                            clampInputs.Width4 = insulationLength;
                    }

                    // For Maintenance Aspect
                    if (clampInputs.Dim2 > 0)
                        clampInputs.Dim2 = clampInputs.Dim2 + maintenanceThickness;
                    else
                        clampInputs.Thickness1 = clampInputs.Thickness1 + maintenanceThickness;
                    if (clampInputs.Height1 > 0)
                        clampInputs.Height1 = clampInputs.Height1 + maintenanceThickness;
                    if (clampInputs.Height2 > 0)
                        clampInputs.Height2 = clampInputs.Height2 + maintenanceThickness;
                    if (clampInputs.Dim1 > 0)
                        clampInputs.Dim1 = clampInputs.Dim1 + maintenanceThickness;
                    if (clampInputs.Thickness2 > 0)
                        clampInputs.Thickness2 = clampInputs.Thickness2 + maintenanceThickness;

                    //'Load slideplate data into 'tSlideplate' struct based on query done using 'sShoe'
                    slidePlateInputs = LoadSlidePlateDataByQuery(slidePlateName);

                    //'Clamp2 data by query
                    clamp2Inputs = LoadSheildDataByQuery(clamp2Name);

                    // Set all overides
                    // ShoeAngle
                    try
                    {
                        if (HgrCompareDoubleService.cmpdbl(slopeAngle, 0) == false)
                            shoeInputs.SlopeAngle = slopeAngle;
                    }
                    catch { shoeInputs.SlopeAngle = 0; }
                    // Shoe Length
                    try
                    {
                        if (HgrCompareDoubleService.cmpdbl(shoeLength, 0) == false)
                            shoeInputs.ShoeLength = shoeLength;
                    }
                    catch { shoeInputs.ShoeLength = 0; }



                    // Add Clamp
                    
                    clampThickness = clampInputs.Width1;
                    if (clampName != "" && clampName != "No Value")
                    {
                        if (clampInputs.Width2 > clampThickness)
                            clampThickness = clampInputs.Width2;
                        if (clampInputs.Diameter1 <= 0)
                        {
                            clampInputs.Diameter1 = GetGreaterValue(diameter1, shieldODia, shield2ODia, repadODia);
                        
                            if (insulationThickness > 0)
                                clampInputs.Diameter1 = insulationThickness * 2 + clampInputs.Diameter1;
                        }
                        SlopeAngleDeg = shoeInputs.SlopeAngle * 180 / Math.PI;//deg
                    }

                    clampOdia = clampInputs.Diameter1 + clampInputs.Thickness1 + clampInputs.Thickness1;

                    // Clamp2
                    clamp2Odia = clamp2Inputs.PipeOD + clamp2Inputs.Thickness1 + clamp2Inputs.Thickness1;
                    if (clamp2Name != "" && clamp2Name != "No Value")
                    {
                        xLocation = new double[int.Parse(multi11Qty.ToString())];
                        for (int index = 0; index < multi11Qty; index++)
                        {
                            xLocation[index] = ShieldMultiPosition(shoeInputs.ShoeLength, multi11Qty, multi11LocateBy, multi11Location, clamp2Inputs.Length1)[index];
                            matrix = new Matrix4X4();
                            matrix.SetIdentity();
                            matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                            matrix.Translate(matrix.Transform(new Vector(-shoeInputs.ShoeLength / 2 + xLocation[index], 0, 0)));
                            AddShield(clamp2Inputs, matrix, m_Maintenance.Outputs, "clamp2" + index);
                        }
                    }
                    outerMostDia = GetGreaterValue(Val1: diameter1, Val2: shieldODia, Val3: shield2ODia, Val4: clampOdia, Val5: GetGreaterValue(repadODia, clamp2Odia, 0, 0, 0));
                    outerMostRadius = (outerMostDia / 2) / (Math.Cos(shoeInputs.SlopeAngle));
                    tempHeight = outerMostRadius + slideVertOffset + (guideInputs.Thickness1 - diameter1 / 2);
                    shoeShapeHeight = shoeHeight - tempHeight;


                    // Add Clamp
                    if (clampName != "" && clampName != "No Value")
                    {
                        clampSlopeOffset = -(Math.Sin(shoeInputs.SlopeAngle) * ((outerMostDia / 2) + shoeInputs.TopPlateGap + shoeInputs.TopPlateOffset));
                        xLocation = new double[int.Parse(multi2Qty.ToString())];
                        for (int index = 0; index < multi2Qty; index++)
                        {
                            xLocation[index] = ShieldMultiPosition(shoeInputs.ShoeLength, multi2Qty, multi2LocateBy, multi2Location, clampThickness)[index];
                            matrix = new Matrix4X4();
                            matrix.SetIdentity();
                            matrix.Rotate(0, new Vector(1, 0, 0));
                            matrix.Translate(new Vector(-shoeInputs.ShoeLength / 2 + xLocation[index] + ClampOffset1 + clampSlopeOffset, 0, 0));
                            matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                            //Addpipeclamp 
                            AddPipeClamp(clampInputs, matrix, m_Maintenance.Outputs, "Shoe-Clamp" + index);

                        }//for end
                    
                    }
                   
                }
            }
            catch (SmartPartSymbolException hgrEx)
            {
                throw hgrEx;
            }
            catch //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of Shoe.cs"));
                    return;
                }
            }

        }
        #endregion

        #region "ICustomWeightCG Members"
        void ICustomWeightCG.EvaluateWeightCG(BusinessObject supportComponentBO)
        {
            try
            {
                ////System WCG Attributes
                Part CatalogPart = (Part)supportComponentBO.GetRelationship("madeFrom", "part").TargetObjects[0];
                CatalogStructHelper catalogStructHelper = new CatalogStructHelper();

                Double weight, cogX, cogY, cogZ;
                try { weight = (double)((PropertyValueDouble)CatalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryWeight")).PropValue; }
                catch { weight = 0; }
                //Center of Gravity
                try { cogX = (double)((PropertyValueDouble)CatalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryCogX")).PropValue; }
                catch { cogX = 0; }
                try { cogY = (double)((PropertyValueDouble)CatalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryCogY")).PropValue; }

                catch { cogY = 0; }
                try { cogZ = (double)((PropertyValueDouble)CatalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryCogZ")).PropValue; }
                catch { cogZ = 0; }
                //Getting Weight as an Input from Excel Sheet
                double inputWeight;
                try { inputWeight = ((double)((PropertyValueDouble)CatalogPart.GetPropertyValue("IJUAhsWeight", "Weight")).PropValue); }
                catch { inputWeight = 0; }

                SupportComponent supportComponent = (SupportComponent)supportComponentBO;

                // 'Get the volume and COG.  Use this if the weight is not defined or set on the override
                VolumeCG shoeVolumeCG;
                shoeVolumeCG = supportComponent.GetVolumeAndCOG();

                string materialType = "", materialGrade = "";
                double materialDensity = 0;
                Material material;

                if (supportComponentBO.SupportsInterface("IJOAhsMaterialEx"))
                {
                    try
                    {
                        materialType = (((PropertyValueString)supportComponentBO.GetPropertyValue("IJOAhsMaterialEx", "MaterialType")).PropValue);
                        materialGrade = (((PropertyValueString)supportComponentBO.GetPropertyValue("IJOAhsMaterialEx", "MaterialGrade")).PropValue);
                    }
                    catch
                    {
                        materialType = String.Empty;
                        materialGrade = String.Empty;
                    }

                }
                else if (CatalogPart.SupportsInterface("IJOAhsMaterialEx"))
                {
                    try
                    {
                        materialType = (((PropertyValueString)CatalogPart.GetPropertyValue("IJOAhsMaterialEx", "MaterialType")).PropValue);
                        materialGrade = (((PropertyValueString)CatalogPart.GetPropertyValue("IJOAhsMaterialEx", "MaterialGrade")).PropValue);
                    }
                    catch
                    {
                        materialType = String.Empty;
                        materialGrade = String.Empty;
                    }

                }
                else
                {
                    materialType = String.Empty;
                    materialGrade = String.Empty;
                }

                try
                {
                    material = catalogStructHelper.GetMaterial(materialType, materialGrade);
                    materialDensity = material.Density;
                }
                catch
                {
                    // the specified MaterialType is not available.refdata needs to be checked.
                    // so assigning 0 to materialDensity.
                    materialDensity = 0;
                }

                //Weigth
                if (!(weight != 0))
                    if (inputWeight != 0)
                    {
                        weight = inputWeight;
                    }
                    else
                    {
                        weight = shoeVolumeCG.Volume * materialDensity;
                    }

                if (!(cogX != 0))
                    cogX = shoeVolumeCG.COGX;

                if (!(cogY != 0))
                    cogY = shoeVolumeCG.COGY;

                if (!(cogZ != 0))
                    cogZ = shoeVolumeCG.COGZ;

                supportComponent.SetWeightAndCOG(weight, cogX, cogY, cogZ);
            }
            catch
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrWeightCG, "Error in Weight CG of Shoe"));
                    return;
                }
            }
        }
        #endregion

    }

}

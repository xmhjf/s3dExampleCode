//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   DummyLeg.cs
//    HSSmartPart,Ingr.SP3D.Content.Support.Symbols.DummyLeg
//   Author       :Manikanth  
//   Creation Date:05/03/2013  
//   Description:Creation Of Dot Net version for Dummyleg

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   05/03/2013  Manikanth    CR-CP-222481  Creation Of Dot Net version for Dummyleg
//   11/08/2014  Ramya        TR-CP-256377  Additional input values are retrieved from catalog part in smart parts  
//   17-Oct-2014 Vinay        CR-CP-253367  Add Insulation Aspect to DummyLeg SmartPart 
//   02-Dec-2014 PVK          DI-CP-253817	Fix priority 2 items to .net SmartParts as a result of new testing  
//   12-12-2014  PVK          TR-CP-264951	Resolve P3 coverity issues found in November 2014 report  
//   28-04-2015  PVK	      Resolve Coverity issues found in April
//   07-05-2015  PVK	      TR-CP-266590	Insulation Aspect of Dummy Leg SmartPart Incorrect
//   10-06-2015  PVK	      TR-CP-274155	SmartPart TDL Errors should be corrected.
//   10-06-2015  PVK	      TR-CP-274155	TR-CP-273182	GetPropertyValue in HSSmartPart should handle null value exception thrown by CLR
//   30-11-2015  VDP          Integrate the newly developed SmartParts into Product(DI-CP-282644)
//   15-03-2106  Siva         TR-CP-287061  Issues found in Anvil2010 parts
//   22-04-2016	 PVK		  Resolved Coverity Issues
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
using System.Collections.Generic;
using Ingr.SP3D.Structure.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle.Services;

namespace Ingr.SP3D.Content.Support.Symbols
{

    [CacheOption(CacheOptionType.NonCached)]
    [SymbolVersion("1.0.0.0")]
    [VariableOutputs]
    public class DummyLeg : SmartPartComponentDefinition, ICustomWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "HSSmartPart,Ingr.SP3D.Content.Support.Symbols.DummyLeg"
        //----------------------------------------------------------------------------------


        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;

        #region "Definition Of AdditionalInputs"
        public override IEnumerable<Input> AdditionalInputs
        {
            get
            {
                int endIndex;
                List<Input> additionalInputs = new List<Input>();
                AddDummyLegInputs(2, out endIndex, additionalInputs);
                additionalInputs.Add(new InputDouble(++endIndex, "BOMLenUnits", "BOMLenUnits", 0, false));
                additionalInputs.Add(new InputDouble(++endIndex, "InsulationTh", "InsulationTh", 0, true));
                additionalInputs.Add(new InputDouble(++endIndex, "CreateInsulationAspect", "CreateInsulationAspect", 0, true));
                additionalInputs.Add(new InputDouble(++endIndex, "InsulationLength", "InsulationLength", 0, true));
                additionalInputs.Add(new InputDouble(++endIndex, "PipeRunInsulationThickness", "PipeRunInsulationThickness", 0, true));
                return additionalInputs;
            }
        }
        #endregion
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Route", "Route")]
        [SymbolOutput("Steel", "Steel")]
        [SymbolOutput("Surface", "Surface")]
        public AspectDefinition m_Symbolic;

        [Aspect("Insulation", "Insulation Aspect", AspectID.Insulation)]
        public AspectDefinition m_Insulation;


        #endregion

        #region "Definition of Additional Outputs"
        public override IEnumerable<OutputDefinition> AdditionalOutputs(string aspectName)
        {
            List<OutputDefinition> additionalOutputs = new List<OutputDefinition>();
            if (aspectName == "Symbolic")
            {
                AddDummyLegOutpus(additionalOutputs);
            }
            return additionalOutputs;
        }
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
                int startIndex;
                //Load DummyLeg/Stanchion Data
                DummyLegInputs dummyLeg = LoadDummyLegData(2, out startIndex);
                
                Double bomUnits = (GetDoubleInputValue(++startIndex));
                Double insulationThickness = (GetDoubleInputValue(++startIndex));
                bool createInsulationAspect = false;
                createInsulationAspect=Convert.ToBoolean((GetDoubleInputValue(++startIndex)));
                Double insulationLength = (GetDoubleInputValue(++startIndex));
                Double runInsulationThickness = (GetDoubleInputValue(++startIndex));
                double locateB = 0, locateT = 0, plt1Thickness = 0, plt2Thickness = 0, plt3Thickness = 0, plt4Thickness = 0, steelPort = 0, surfacePort = 0, botHeight = 0, topHeight = 0;
                int count = 0;
                bool isFixedLength = true;
                StanchionShapeInputs botStanchionShape = new StanchionShapeInputs();
                StanchionShapeInputs topStanchionShape = new StanchionShapeInputs();
                DummyLegShapeInputs botDummyShape = new DummyLegShapeInputs();
                DummyLegShapeInputs topDummyShape = new DummyLegShapeInputs();

                string maxLength = "";
                string minLength = "";
                string value = "";
                topStanchionShape.stanShape = 0;
                //Intialise plate dimensions
                PlateInputs plate1 = new PlateInputs();
                PlateInputs plate2 = new PlateInputs();
                Matrix4X4 matrix = new Matrix4X4();
                int isStanchion = 1;
                isStanchion= dummyLeg.isStanchion;

                if (isStanchion == 1) //Part is completly of Stanchion shapes
                {
                    if (!(dummyLeg.botShape == "" || dummyLeg.botShape == "No Value"))
                    {
                        //Load Bottom and Top Stanchion Shape's Data
                        botStanchionShape = LoadStanchShapeDataByQuery(dummyLeg.botShape);
                        topStanchionShape = LoadStanchShapeDataByQuery(dummyLeg.topShape);
                        botHeight = botStanchionShape.stanHeight;
                        topHeight = topStanchionShape.stanHeight;

                        //Load Bottom Shape Plates Data
                        if (!(botStanchionShape.plate1Shape == "" || botStanchionShape.plate1Shape == "No Value"))
                        {
                            plate1 = LoadPlateDataByQuery(botStanchionShape.plate1Shape);
                            plt1Thickness = plate1.thickness1;
                        }
                        if (!(botStanchionShape.plate2Shape == "" || botStanchionShape.plate2Shape == "No Value"))
                        {
                            plate2 = LoadPlateDataByQuery(botStanchionShape.plate2Shape);
                            plt2Thickness = plate2.thickness1;
                        }
                        //Load Top Shape Plates Data
                        if (!(topStanchionShape.plate1Shape == "" || topStanchionShape.plate1Shape == "No Value"))
                        {
                            plate1 = LoadPlateDataByQuery(topStanchionShape.plate1Shape);
                            plt3Thickness = plate1.thickness1;
                        }
                        if (!(topStanchionShape.plate2Shape == "" || topStanchionShape.plate2Shape == "No Value"))
                        {
                            plate2 = LoadPlateDataByQuery(topStanchionShape.plate2Shape);
                            plt4Thickness = plate2.thickness1;
                        }
                    }
                    else
                    {
                        RaiseSmartPartTDLError(SmartPartSymbolResourceIDs.ErrInvalidTypeStanchion, "Bottom Shape is required to create Stanchion");
                        return;
                    }
                }
                if (isStanchion == 2)  //Part has Dummy Leg shape
                {
                    if (!(dummyLeg.botShape == "" || dummyLeg.botShape == "No Value"))
                    {
                        dummyLeg.stanGap = 0;   //Dummy Leg should attach to the pipe
                        if (!(dummyLeg.topShape == "" || dummyLeg.topShape == "No Value"))   //Part has both Top and Bottom Shapes
                        {
                            //Load Bottom and Top Shape's Data
                            botStanchionShape = LoadStanchShapeDataByQuery(dummyLeg.botShape);
                            topDummyShape = LoadDummyLegShapeDataByQuery(dummyLeg.topShape);

                            if (HgrCompareDoubleService.cmpdbl(dummyLeg.elbowRadius, 0) == true && HgrCompareDoubleService.cmpdbl(dummyLeg.faceToCenter, 0) == true) 
                            {
                                //Pipe Center to DummyLeg bottom distance
                                topDummyShape.dummyHeight = dummyLeg.diameter1 / 2.0 + topDummyShape.dummyHeight;
                            }
                            else
                            {
                                topDummyShape.dummyHeight = topDummyShape.dummyHeight + dummyLeg.diameter1 / 2.0 + dummyLeg.faceToCenter;
                            }

                            //Set if DummyLeg supporting pipe dia is not given use dia from partclass
                            if (HgrCompareDoubleService.cmpdbl(topDummyShape.diameter, 0) == true && HgrCompareDoubleService.cmpdbl(dummyLeg.diameter1, 0) == false)
                            {
                                topDummyShape.diameter = dummyLeg.diameter1;
                            }
                            else if (HgrCompareDoubleService.cmpdbl(topDummyShape.diameter, 0) == true && HgrCompareDoubleService.cmpdbl(dummyLeg.diameter1, 0) == true)
                            {
                                RaiseSmartPartTDLError(SmartPartSymbolResourceIDs.ErrInvalidPipeDiaStanchion, "Pipe Diameter is required to create Dummy Leg");
                                return;
                            }

                            botHeight = botStanchionShape.stanHeight;
                            topHeight = topDummyShape.dummyHeight;

                            //Load Bottom Shape Plates Data
                            if (!(botStanchionShape.plate1Shape == "" || botStanchionShape.plate1Shape == "No Value"))
                            {
                                plate1 = LoadPlateDataByQuery(botStanchionShape.plate1Shape);
                                plt1Thickness = plate1.thickness1;
                            }
                            if (!(botStanchionShape.plate2Shape == "" || botStanchionShape.plate2Shape == "No Value"))
                            {
                                plate2 = LoadPlateDataByQuery(botStanchionShape.plate2Shape);
                                plt2Thickness = plate2.thickness1;
                            }
                            if (!(topDummyShape.plateShape == "" || topDummyShape.plateShape == "No Value"))
                            {
                                plate1 = LoadPlateDataByQuery(topDummyShape.plateShape);
                                plt3Thickness = plate1.thickness1;
                            }

                        }

                        else  //Part has only Bootom DummyLeg Shape
                        {
                            botDummyShape = LoadDummyLegShapeDataByQuery(dummyLeg.botShape);

                            if (HgrCompareDoubleService.cmpdbl(dummyLeg.elbowRadius, 0) == true )  //Pipe Center to DummyLeg bottom distance
                            {
                                botDummyShape.dummyHeight = (dummyLeg.diameter1 / 2.0) + botDummyShape.dummyHeight;
                            }
                            else
                            {
                                botDummyShape.dummyHeight = botDummyShape.dummyHeight + dummyLeg.diameter1 / 2.0 + dummyLeg.faceToCenter;
                            }

                            if (HgrCompareDoubleService.cmpdbl(botDummyShape.diameter, 0) == true  && HgrCompareDoubleService.cmpdbl(dummyLeg.diameter1, 0) == false )  //Set if DummyLeg supporting pipe dia is not given use dia from partclass
                            {
                                botDummyShape.diameter = dummyLeg.diameter1;
                            }
                            else if (HgrCompareDoubleService.cmpdbl(botDummyShape.diameter , 0) == true && HgrCompareDoubleService.cmpdbl(dummyLeg.diameter1, 0) == true)
                            {
                                RaiseSmartPartTDLError(SmartPartSymbolResourceIDs.ErrInvalidPipeDiaStanchion, "Pipe Diameter is required to create Dummy Leg");
                                return;
                            }
                            botHeight = botDummyShape.dummyHeight;

                            //Load DummyLeg Plate Data
                            if (!(botDummyShape.plateShape == "" || botDummyShape.plateShape == "No Value"))
                            {
                                plate1 = LoadPlateDataByQuery(botDummyShape.plateShape);
                                plt1Thickness = plate1.thickness1;
                            }
                        }

                    }
                    else
                    {
                        RaiseSmartPartTDLError(SmartPartSymbolResourceIDs.ErrInvalidBotShapeStanchion, "Bottom Shape is required to create Dummy Leg");
                        return;
                    }

                }
                double variableLength = 0, fixedLength = 0;

                isFixedLength = true;
                if (part.SupportsInterface("IJUAhsLength"))
                {
                    try
                    {
                        fixedLength = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAhsLength", "Length")).PropValue;
                    }
                    catch
                    {
                        fixedLength = 0;
                    }
                }
                else if (part.SupportsInterface("IJOAHgrOccLength"))
                {
                    try
                    {
                        //No need to get this value from Occurence Attribute.
                        variableLength = (double)((PropertyValueDouble)part.GetPropertyValue("IJOAHgrOccLength", "Length")).PropValue;
                    }
                    catch
                    {
                        variableLength = 0;
                    }
                    isFixedLength = false;
                }
                else
                {
                    fixedLength = 0;
                    variableLength = 0;
                }

                if (isFixedLength == true)
                {
                    if (fixedLength <= 0) //if Length value is not provided, it must have atleast both Bottom and Top Shape's(if exists) height.
                    {
                        if (!(dummyLeg.botShape == "" || dummyLeg.botShape == "No Value"))
                        {
                            if (!(dummyLeg.topShape == "" || dummyLeg.topShape == "No Value"))
                            {
                                if (HgrCompareDoubleService.cmpdbl(botHeight , 0) == true || HgrCompareDoubleService.cmpdbl(topHeight , 0) == true) 
                                {
                                    RaiseSmartPartTDLError(SmartPartSymbolResourceIDs.ErrInvalidBotShapeLengthStanchion, "For Fixed Length Part, either Length or BottomShape and TopShape(if exists) Heights are to be provided");
                                    return;
                                }
                            }
                            else if (HgrCompareDoubleService.cmpdbl(botHeight, 0) == true)
                            {
                                RaiseSmartPartTDLError(SmartPartSymbolResourceIDs.ErrInvalidBotShapeLengthStanchion, "For Fixed Length Part, either Length or BottomShape and TopShape(if exists) Heights are to be provided");
                                return;
                            }
                            else
                            {
                                surfacePort = dummyLeg.stanGap + (dummyLeg.diameter1 / 2.0);
                                dummyLeg.length = surfacePort + dummyLeg.stanGap + botHeight + plt1Thickness + plt2Thickness + topHeight + plt3Thickness + plt4Thickness;
                                steelPort = dummyLeg.length;
                            }


                        }
                        else
                        {
                            RaiseSmartPartTDLError(SmartPartSymbolResourceIDs.ErrInvalidReqBotShapeStanchion, "BottomShape is reqiured for part placement");
                            return;
                        }
                    }
                    else
                    {
                        surfacePort = dummyLeg.stanGap + (dummyLeg.diameter1 / 2.0);
                        steelPort = dummyLeg.length;
                    }
                }

                else
                {
                    if (HgrCompareDoubleService.cmpdbl(variableLength, 0) == false)        //In case of Variable length interface, default Length should not be provided in the workbook, if provided then the part will be a fixeed length part
                    {
                        surfacePort = dummyLeg.stanGap + (dummyLeg.diameter1 / 2.0);
                        dummyLeg.length = dummyLeg.length + dummyLeg.stanGap;
                        steelPort = dummyLeg.length;
                        isFixedLength = true;
                    }
                }

                try
                {
                    PropertyValueCodelist bomList = (PropertyValueCodelist)part.GetPropertyValue("IJUAhsBOMLenUnits", "BOMLenUnits");
                    value = bomList.PropertyInfo.CodeListInfo.GetCodelistItem((int)bomUnits).DisplayName;
                }
                catch
                {
                    value = "in";
                }
                if (value.ToUpper() == "IN")
                {
                    minLength = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, dummyLeg.minLen, UnitName.DISTANCE_INCH);
                    maxLength = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, dummyLeg.maxLen, UnitName.DISTANCE_INCH);
                }
                else if (value.ToUpper() == "FT")
                {
                    minLength = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, dummyLeg.minLen, UnitName.DISTANCE_FOOT);
                    maxLength = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, dummyLeg.maxLen, UnitName.DISTANCE_FOOT);
                }
                else if (value.ToUpper() == "MM")
                {
                    minLength = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, dummyLeg.minLen, UnitName.DISTANCE_MILLIMETER);
                    maxLength = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, dummyLeg.maxLen, UnitName.DISTANCE_MILLIMETER);
                }
                else if (value.ToUpper() == "M")
                {
                    minLength = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, dummyLeg.minLen, UnitName.DISTANCE_METER);
                    maxLength = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, dummyLeg.maxLen, UnitName.DISTANCE_METER);
                }

                if (isFixedLength == false && isStanchion == 1)
                {
                    if (dummyLeg.length <= 0) //At first part placement
                    {
                        if (dummyLeg.topShape == "" || dummyLeg.topShape == "No Value")
                        {
                            if (botHeight <= 0) //BotShape is Stretchable
                            {
                                botStanchionShape.stanHeight = 0.5;
                                dummyLeg.length = (dummyLeg.diameter1 / 2.0) + botStanchionShape.stanHeight + plt1Thickness + plt2Thickness;
                            }
                            else
                            {
                                isFixedLength = true;
                            }
                        }
                        else
                        {
                            if (botHeight <= 0 && topHeight <= 0)
                            {
                                RaiseSmartPartTDLError(SmartPartSymbolResourceIDs.ErrInvalidBotTopHeightStanchion, "Need to Provide Height for atleast one of Bottom and Top shapes");
                                return;
                            }
                            else if (botHeight <= 0 && HgrCompareDoubleService.cmpdbl(topHeight, 0) == false) //Bottom shape is stretchable
                            {
                                botStanchionShape.stanHeight = 0.5;
                                topStanchionShape.stanHeight = topStanchionShape.stanHeight + dummyLeg.offset2;
                                dummyLeg.length = (dummyLeg.diameter1 / 2.0) + botStanchionShape.stanHeight + plt1Thickness + plt2Thickness + topStanchionShape.stanHeight + plt3Thickness + plt4Thickness - dummyLeg.offset2;
                            }
                            else if (topHeight <= 0 && HgrCompareDoubleService.cmpdbl(botHeight, 0) == false) //Top shape is Strechable
                            {
                                topStanchionShape.stanHeight = 0.5 + dummyLeg.offset2;
                                dummyLeg.length = (dummyLeg.diameter1 / 2.0) + botStanchionShape.stanHeight + plt1Thickness + plt2Thickness + topStanchionShape.stanHeight + plt3Thickness + plt4Thickness - dummyLeg.offset2;
                                locateB = -dummyLeg.length;
                                locateT = (dummyLeg.diameter1 / 2.0) + dummyLeg.stanGap + plt3Thickness + plt4Thickness + topStanchionShape.stanHeight;
                            }
                        }
                        steelPort = dummyLeg.length;
                        surfacePort = dummyLeg.stanGap + dummyLeg.diameter1 / 2.0;
                        count = 1; //Used for Stanchion Placement case iCount = 4
                        isFixedLength = true;

                    }
                    else
                    {
                        if (dummyLeg.topShape == "" || dummyLeg.topShape == "No Value")
                        {
                            if (botHeight <= 0) //BotShape is Stretchable
                            {
                                botHeight = -dummyLeg.length + dummyLeg.stanGap + (dummyLeg.diameter1 / 2.0) + plt1Thickness + plt2Thickness;
                                botStanchionShape.stanHeight = botHeight;
                            }
                            else
                            {
                                isFixedLength = true;
                                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrInvalidBotHeightStanchion, "Bottom shape has given Height value. So part cannot be a Streached"));

                            }
                        }
                        else
                        {
                            if (botHeight <= 0 && topHeight <= 0)
                            {
                                RaiseSmartPartTDLError(SmartPartSymbolResourceIDs.ErrNoHeightStanchion, "Need to Provide Height for atleast one of Bottom and Top shapes");
                                return;
                            }
                            else if (botHeight <= 0 && HgrCompareDoubleService.cmpdbl(topHeight, 0) == false)  //Bottom shape is stretchable
                            {
                                botHeight = -dummyLeg.length + dummyLeg.stanGap + dummyLeg.diameter1 / 2 + plt1Thickness + plt2Thickness + plt3Thickness + plt4Thickness + topHeight - dummyLeg.offset2;
                                botStanchionShape.stanHeight = botHeight;
                                topStanchionShape.stanHeight = -topHeight;
                            }
                            else if (topHeight <= 0 && HgrCompareDoubleService.cmpdbl(botHeight, 0) == false) //Top shape is Strechable
                            {
                                topHeight = -dummyLeg.length + dummyLeg.stanGap + dummyLeg.diameter1 / 2 + plt1Thickness + plt2Thickness + plt3Thickness + plt4Thickness + botHeight;
                                topStanchionShape.stanHeight = topHeight - dummyLeg.offset2;
                                botStanchionShape.stanHeight = -botHeight;
                            }
                            else if (botHeight > 0 && topHeight > 0)
                            {
                                count = 2;                            //iCount = 5
                                isFixedLength = true;
                                steelPort = -steelPort;
                                topStanchionShape.stanHeight = -topHeight;
                                botStanchionShape.stanHeight = -botHeight;
                                dummyLeg.offset2 = 0;
                            }
                        }

                        dummyLeg.length = -(dummyLeg.diameter1 / 2.0) - dummyLeg.stanGap - plt3Thickness - plt4Thickness + topStanchionShape.stanHeight - plt1Thickness - plt2Thickness + botStanchionShape.stanHeight + dummyLeg.offset2;
                        locateB = -dummyLeg.length;
                        locateT = (dummyLeg.diameter1 / 2.0) + dummyLeg.stanGap + plt3Thickness + plt4Thickness - topStanchionShape.stanHeight;
                        surfacePort = -dummyLeg.stanGap - dummyLeg.diameter1 / 2.0;
                        steelPort = dummyLeg.length;
                    }
                }
                else if (isFixedLength == false && isStanchion == 2)  //Dummy Leg
                {
                    if (dummyLeg.length <= 0)      //At first part placement
                    {
                        double b = (dummyLeg.diameter1 / 2.0 + dummyLeg.faceToCenter);
                        if (dummyLeg.topShape == "" || dummyLeg.topShape == "No Value")
                        {
                            if (botHeight <= b)    //BotShape is Stretchable
                            {
                                botDummyShape.dummyHeight = 0.5 + dummyLeg.faceToCenter + dummyLeg.diameter1 / 2.0;
                                dummyLeg.length = botDummyShape.dummyHeight + plt1Thickness;
                                steelPort = dummyLeg.length;
                            }
                            else
                            {
                                isFixedLength = true;
                            }
                        }
                        else
                        {
                            if (botHeight <= 0 && topHeight <= 0)
                            {
                                RaiseSmartPartTDLError(SmartPartSymbolResourceIDs.ErrNoHeightStanchion, "Need to Provide Height for atleast one of Bottom and Top shapes");
                                return;
                            }
                            else if (botHeight <= 0 && HgrCompareDoubleService.cmpdbl(topHeight, 0) == false)    //Bottom shape is stretchable
                            {
                                count = 3;
                                topDummyShape.dummyHeight = topDummyShape.dummyHeight + dummyLeg.offset2;
                                botStanchionShape.stanHeight = 0.5;
                                dummyLeg.length = -botStanchionShape.stanHeight - plt1Thickness - plt2Thickness - topDummyShape.dummyHeight + plt3Thickness + dummyLeg.offset2;
                                steelPort = -dummyLeg.length;
                            }


                            else if ((topHeight <= b) && HgrCompareDoubleService.cmpdbl(botHeight, 0) == false)   //Top shape is Strechable
                            {
                                topDummyShape.dummyHeight = 0.5 + dummyLeg.offset2 + dummyLeg.faceToCenter + dummyLeg.diameter1 / 2.0;
                                dummyLeg.length = (botStanchionShape.stanHeight + plt1Thickness + plt2Thickness + topDummyShape.dummyHeight + plt3Thickness - dummyLeg.offset2);
                                steelPort = dummyLeg.length;
                            }
                        }
                        isFixedLength = true;
                    }
                    else
                    {
                        double b = (dummyLeg.diameter1 / 2.0 + dummyLeg.faceToCenter);
                        if (dummyLeg.topShape == "" || dummyLeg.topShape == "No Value")
                        {
                            if (botHeight <= b)   //BotShape is Stretchable
                            {
                                botHeight = -dummyLeg.length + plt1Thickness;
                                botDummyShape.dummyHeight = botHeight;
                            }
                            else
                            {
                                isFixedLength = true;
                                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrInvalidBotHeightStanchion, "Bottom shape has given Height value. So part cannot be a Streached"));
                            }
                        }
                        else
                        {
                            if (botHeight <= 0 && topHeight <= b)
                            {
                                RaiseSmartPartTDLError(SmartPartSymbolResourceIDs.ErrNoHeightStanchion, "Need to Provide Height for atleast one of Bottom and Top shapes");
                                return;
                            }
                            else if (botHeight <= 0 && HgrCompareDoubleService.cmpdbl(topHeight, 0) == false) //Bottom shape is stretchable
                            {
                                botHeight = dummyLeg.length - plt1Thickness - plt2Thickness + plt3Thickness - topHeight - dummyLeg.offset2;
                                botStanchionShape.stanHeight = -botHeight;
                                topDummyShape.dummyHeight = -topHeight;
                            }
                            else if (topHeight <= b && HgrCompareDoubleService.cmpdbl(botHeight, 0) == false)    //Top shape is Strechable
                            {
                                topHeight = -(dummyLeg.length - plt1Thickness - plt2Thickness - plt3Thickness - botHeight);
                                topDummyShape.dummyHeight = topHeight - dummyLeg.offset2;
                                botStanchionShape.stanHeight = -botHeight;
                            }
                            else if (botHeight > 0 && topHeight > 0)
                            {
                                count = 4;
                                isFixedLength = true;
                                topDummyShape.dummyHeight = -topHeight;
                                botStanchionShape.stanHeight = -botHeight;
                                dummyLeg.offset2 = 0;
                            }
                        }
                        steelPort = -dummyLeg.length;

                    }



                }

                if (isFixedLength == false)
                {
                    if (dummyLeg.length < dummyLeg.minLen || dummyLeg.length > dummyLeg.maxLen)
                    {
                        ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrInvalidLengthOfDummyLeg, "Lenth of the DummyLeg must be between" + "" + minLength + "and" + maxLength));
                    }
                }
                if (isFixedLength == false && HgrCompareDoubleService.cmpdbl(dummyLeg.length, 0) == false)    //Add Route Port
                {
                    Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                    m_Symbolic.Outputs["Route"] = port1;

                }
                else
                {
                    Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, -1));
                    m_Symbolic.Outputs["Route"] = port1;

                }
                if (isStanchion == 2 && isFixedLength == false)
                {
                    if (Math.Abs(dummyLeg.length) < (dummyLeg.diameter1 / 2.0) || Math.Abs(dummyLeg.length) < (dummyLeg.elbowRadius + dummyLeg.diameter1 / 2.0))
                    {
                        steelPort = -(dummyLeg.elbowRadius + dummyLeg.diameter1 + plt2Thickness + plt1Thickness + plt3Thickness - botStanchionShape.stanHeight);
                    }

                }
                if (isFixedLength == false && HgrCompareDoubleService.cmpdbl(dummyLeg.length, 0) == false)   //Add Steel port
                {
                    Port port2 = new Port(OccurrenceConnection, part, "Steel", new Position(0, 0, -steelPort), new Vector(1, 0, 0), new Vector(0, 0, -1));
                    m_Symbolic.Outputs["Steel"] = port2;

                }
                else
                {
                    Port port2 = new Port(OccurrenceConnection, part, "Steel", new Position(0, 0, -steelPort), new Vector(1, 0, 0), new Vector(0, 0, 1));
                    m_Symbolic.Outputs["Steel"] = port2;


                }
                if (isStanchion == 1)    //Add Surface port for Stanchion Shape only
                {
                    if (isFixedLength == false && HgrCompareDoubleService.cmpdbl(dummyLeg.length, 0) == false)
                    {
                        Port port3 = new Port(OccurrenceConnection, part, "Surface", new Position(0, 0, -surfacePort), new Vector(1, 0, 0), new Vector(0, 0, -1));
                        m_Symbolic.Outputs["Surface"] = port3;

                    }
                    else
                    {
                        Port port3 = new Port(OccurrenceConnection, part, "Surface", new Position(0, 0, -surfacePort), new Vector(1, 0, 0), new Vector(0, 0, 1));
                        m_Symbolic.Outputs["Surface"] = port3;

                    }
                }

                //Load Bottom Dummy/Stanchion Shape
                if (!(dummyLeg.botShape == "" || dummyLeg.botShape == "No Value"))
                {
                    if ((dummyLeg.topShape == "" || dummyLeg.topShape == "No Value") && isStanchion == 2)
                    {
                        //Create DummyLeg Shape
                        matrix = new Matrix4X4();
                        matrix.SetIdentity();
                        matrix.Translate(new Vector(0, 0, 0));
                        AddDummyLegShape(botDummyShape, botDummyShape.diameter, dummyLeg.elbowRadius, dummyLeg.offset1, matrix, m_Symbolic.Outputs, "DummyStanchBot", isFixedLength);
                    }
                    else if (!(dummyLeg.topShape == "" || dummyLeg.topShape == "No Value") && isStanchion == 2)
                    {
                        if (isFixedLength == true && count != 3)
                        {
                            locateB = -dummyLeg.length;
                        }
                        else if (count == 3)
                        {
                            locateB = dummyLeg.length;
                            //isFixedLength = true;
                        }
                        else
                        {
                            locateB = dummyLeg.length;
                        }
                        if (isStanchion == 2 && isFixedLength == false)
                        {
                            if (Math.Abs(dummyLeg.length) < (dummyLeg.diameter1 / 2.0) || Math.Abs(dummyLeg.length) < (dummyLeg.elbowRadius + dummyLeg.diameter1 / 2.0))
                            {
                                locateB = dummyLeg.elbowRadius + dummyLeg.diameter1 + plt2Thickness + plt1Thickness + plt3Thickness - botStanchionShape.stanHeight;
                            }
                        }


                        matrix = new Matrix4X4();
                        matrix.SetIdentity();
                        matrix.Translate(new Vector(0, 0, locateB));
                        AddStanchionShape(botStanchionShape, matrix, m_Symbolic.Outputs, "DummyStanchBot", isFixedLength);
                    }
                    else
                    {
                        if (isStanchion == 1 && isFixedLength == true)
                        {
                            if (HgrCompareDoubleService.cmpdbl(Math.Round(botStanchionShape.stanHeight - dummyLeg.stanGap, 4), 0) == true)
                            {
                                RaiseSmartPartTDLError(SmartPartSymbolResourceIDs.ErrInvalidStanchionGapAndHeight, "Stanchion Height and Stanchion Gap should not be equal");
                                return;
                            }
                            botStanchionShape.stanHeight = botStanchionShape.stanHeight - dummyLeg.stanGap;
                        }
                        if (isFixedLength == true)
                        {
                            locateB = -(dummyLeg.stanGap + dummyLeg.diameter1 / 2.0 + topStanchionShape.stanHeight - dummyLeg.offset2 + plt3Thickness + plt4Thickness + botStanchionShape.stanHeight + plt2Thickness + plt1Thickness);
                            if (count == 2)
                            {
                                locateB = -(-dummyLeg.stanGap - dummyLeg.diameter1 / 2.0 + topStanchionShape.stanHeight - dummyLeg.offset2 + plt3Thickness + plt4Thickness + botStanchionShape.stanHeight + plt2Thickness + plt1Thickness);
                            }

                        }
                        matrix = new Matrix4X4();
                        matrix.SetIdentity();
                        matrix.Translate(new Vector(0, 0, locateB));
                        AddStanchionShape(botStanchionShape, matrix, m_Symbolic.Outputs, "DummyStanchBot", isFixedLength);

                    }

                    //Load Top Stanchion
                    if (!(dummyLeg.topShape == "" || dummyLeg.topShape == "No Value") && !(dummyLeg.botShape == "" || dummyLeg.botShape == "No Value"))
                    {
                        if (isStanchion == 1)
                        {
                            if (isFixedLength == true)
                            {
                                locateT = dummyLeg.diameter1 / 2.0 + dummyLeg.stanGap + topStanchionShape.stanHeight + plt3Thickness + plt4Thickness;
                                if (count == 1)
                                    locateT = -(dummyLeg.diameter1 / 2.0 + dummyLeg.stanGap + topStanchionShape.stanHeight + plt3Thickness + plt4Thickness);
                            }
                            if (topStanchionShape.stanShape == 3 && botStanchionShape.plate2Shape == "" && topStanchionShape.plate1Shape == "")
                            {
                                RaiseSmartPartTDLError(SmartPartSymbolResourceIDs.ErrNoStanchforStandardsteel, "Top stanchion shape cannot be created for Standard steel without Baseplate");
                                return;
                            }
                            else
                            {
                                if (isFixedLength == false)
                                    locateT = -locateT;
                                if (count == 1)
                                {
                                    isFixedLength = true;
                                    locateT = -locateT;
                                }

                                matrix = new Matrix4X4();
                                matrix.SetIdentity();
                                matrix.Translate(new Vector(0, 0, -locateT));
                                AddStanchionShape(topStanchionShape, matrix, m_Symbolic.Outputs, "DummyStanchTop", isFixedLength);  //Create Shape
                            }
                        }
                        else
                        {
                            matrix = new Matrix4X4();
                            matrix.SetIdentity();
                            matrix.Translate(new Vector(0, 0, 0));
                            AddDummyLegShape(topDummyShape, topDummyShape.diameter, dummyLeg.elbowRadius, dummyLeg.offset1, matrix, m_Symbolic.Outputs, "DummyLegTop", isFixedLength);  //DummyLeg Shape
                        }
                    }
                }

                if (createInsulationAspect)// create insulation aspect
                {
                    if (!(dummyLeg.botShape == "" || dummyLeg.botShape == "No Value"))
                    {
                        if ((dummyLeg.topShape == "" || dummyLeg.topShape == "No Value") && isStanchion == 2)
                        {
                            matrix = new Matrix4X4();
                            matrix.SetIdentity();
                            matrix.Translate(new Vector(0, 0, 0));
                            AddDummyLegInsulationShape(botDummyShape, botDummyShape.diameter, dummyLeg.elbowRadius, dummyLeg.offset1, matrix, m_Insulation.Outputs, "InsulationDummyStanchBot", isFixedLength, insulationThickness, ref insulationLength, runInsulationThickness);
                        }
                    }
                    if (!(dummyLeg.topShape == "" || dummyLeg.topShape == "No Value") && !(dummyLeg.botShape == "" || dummyLeg.botShape == "No Value"))
                    {
                        if (isStanchion == 2)
                        {
                            matrix = new Matrix4X4();
                            matrix.SetIdentity();
                            matrix.Translate(new Vector(0, 0, 0));
                            AddDummyLegInsulationShape(topDummyShape, topDummyShape.diameter, dummyLeg.elbowRadius, dummyLeg.offset1, matrix, m_Insulation.Outputs, "InsulationDummyLegTop", isFixedLength, insulationThickness, ref insulationLength, runInsulationThickness); 
                        }
                    }
                    SupportComponent supportComponnet = (SupportComponent)Occurrence;
                    supportComponnet.SetPropertyValue(insulationLength, "IJOAhsInsulationL", "InsulationLength");//set the reset value of insulation length on part occurance

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
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of DummyLeg"));
                    return;
                }
            }
        }
        #endregion

        #region "ICustomHgrWeightCG Members"

        public void EvaluateWeightCG(BusinessObject supportComponentBO)
        {
            try
            {
                CatalogStructHelper catalogStructHelper = new CatalogStructHelper();
                string botshape, topShape;
                double dryWeight, dryCogX, dryCogy, dryCogZ, botomHt = 0,stanGap, topHT = 0,length = 0, elbowRadius, faceToCenter = 0, diameter1, areaCS = 0;
                
                int isStanchion, topDummyType = 0, bottomStanType = 0, bottomDummyType = 0;
                DummyLegShapeInputs topDummyLegshape, botDummylegShape;
                StanchionShapeInputs botStanShape, topStanShape;
                PlateInputs plate1, plate2;
                double plate1Thk = 0, plate2Thk = 0, plate3Thk = 0, plate4Thk = 0, plate1Width = 0, plate2Width = 0, plate3Width = 0, plate4Width = 0, plate1Len = 0, plate2len = 0, plate3Len = 0, plate4Len = 0, plate1Vol = 0, plate2Vol = 0, plate3Vol = 0, plate4Vol = 0, totPlateThk = 0, totPlateVol = 0, totVol = 0;
                double bottomStanhgt = 0, bottomDummyHt = 0, topStanHt = 0, bottomStanWdth = 0, bottomDummyWdth = 0, topStanWdth = 0, bottomStanDpth = 0, bottomDummyDpth = 0, topStanDpth = 0, topDummyDpth = 0, topDummyWdth = 0, topDummyHt = 0, topLegVol = 0, bottomLegVol = 0, totLegVol = 0;
                string bottomSteelName = "", bottomSteelType = "", bottomSteelStd = "", topSteelName = "", topSteelType = "", topSteelStd = "";
                int topStanType = 0;
                string materialType = "", materialGrade = "";
                double materialDensity = 0;
                Part catalogPart = (Part)supportComponentBO.GetRelationship("madeFrom", "part").TargetObjects[0];
                double weight = 0, cogX, cogY, cogZ;

                try
                {
                    diameter1 = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJOAhsDiameter1", "Diameter1")).PropValue;
                }
                catch
                {
                    diameter1 = 0;
                }
                botshape = (string)((PropertyValueString)catalogPart.GetPropertyValue("IJUAhsBotShape", "BotShape")).PropValue;
                if (botshape == null)
                {
                    botshape = "";
                }
                topShape = (string)((PropertyValueString)catalogPart.GetPropertyValue("IJUAhsTopShape", "TopShape")).PropValue;
                if (topShape == null)
                {
                    topShape = "";
                }

                isStanchion = (int)((PropertyValueCodelist)catalogPart.GetPropertyValue("IJUAhsStanchion", "IsStanchion")).PropValue;
                try
                {
                    stanGap = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJOAhsStanGap", "StanGap")).PropValue;

                }
                catch
                {
                    stanGap = 0;
                }
                try
                {
                    elbowRadius = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJOAhsElbow", "ElbowRadius")).PropValue;
                }
                catch
                {
                    elbowRadius = 0;
                }
                try
                {
                    faceToCenter = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJOAhsElbow", "FaceToCenter")).PropValue;
                }
                catch
                {
                    faceToCenter = 0;
                }


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
                else if (catalogPart.SupportsInterface("IJOAhsMaterialEx"))
                {
                    try
                    {
                        materialType = (((PropertyValueString)catalogPart.GetPropertyValue("IJOAhsMaterialEx", "MaterialType")).PropValue);
                        materialGrade = (((PropertyValueString)catalogPart.GetPropertyValue("IJOAhsMaterialEx", "MaterialGrade")).PropValue);
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
                    dryWeight = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryWeight")).PropValue;
                }
                catch
                {
                    dryWeight = 0;
                }
                //Center of Gravity
                try
                {
                    dryCogX = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryCogX")).PropValue;
                }
                catch
                {
                    dryCogX = 0;
                }
                try
                {
                    dryCogy = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryCogY")).PropValue;
                }
                catch
                {
                    dryCogy = 0;
                }
                try
                {
                    dryCogZ = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryCogZ")).PropValue;
                }
                catch
                {
                    dryCogZ = 0;
                }

                bool isFixedLeng = true;
                if (supportComponentBO.SupportsInterface("IJOAHgrOccLength"))
                {
                    isFixedLeng = false;
                }

                if (isStanchion == 1)
                {
                    if (!(botshape == "" || botshape == "No Value"))  //Load Bottom and Top Stanchion Shape's Data
                    {
                        botStanShape = LoadStanchShapeDataByQuery(botshape);
                        topStanShape = LoadStanchShapeDataByQuery(topShape);

                        //Setting the width,height, and depth in case of a square or rectangular crosssection
                        if (isFixedLeng == false)
                        {
                            if (!(botshape == "" || botshape == "No Value"))
                            {
                                if (HgrCompareDoubleService.cmpdbl(botStanShape.stanHeight, 0) == true)
                                {
                                    botStanShape.stanHeight = 0.5;
                                }
                            }
                        }
                        if (isFixedLeng == false)
                        {
                            if (!(topShape == "" || topShape == "No Value"))
                            {
                                if (HgrCompareDoubleService.cmpdbl(topStanShape.stanHeight , 0) == true)
                                {
                                    topStanShape.stanHeight = 0.5;
                                }
                            }
                        }
                        

                        bottomStanWdth = botStanShape.stanWidth;
                        topStanWdth = topStanShape.stanWidth;
                        bottomStanDpth = botStanShape.stanDepth;
                        topStanDpth = topStanShape.stanDepth;
                        topSteelName = topStanShape.steelName;
                        topSteelType = topStanShape.steelType;
                        topSteelStd = topStanShape.steelStandard;
                        topStanType = Convert.ToInt32(topStanShape.stanShape);
                        bottomStanType = Convert.ToInt32(botStanShape.stanShape);

                        //Load Bottom Shape Plates Data
                        if (!(botStanShape.plate1Shape == "" || botStanShape.plate1Shape == "No Value"))
                        {
                            plate1 = LoadPlateDataByQuery(botStanShape.plate1Shape);
                            plate1Thk = plate1.thickness1;
                            plate1Width = plate1.width1;
                            plate1Len = plate1.length1;

                        }

                        if (!(botStanShape.plate2Shape == "" || botStanShape.plate2Shape == "No Value"))
                        {
                            plate2 = LoadPlateDataByQuery(botStanShape.plate2Shape);
                            plate2Thk = plate2.thickness1;
                            plate2Width = plate2.width1;
                            plate2len = plate2.length1;
                        }
                        //Load Top Shape Plates Data
                        if (!(topStanShape.plate1Shape == "" || topStanShape.plate1Shape == "No Value"))
                        {
                            plate1 = LoadPlateDataByQuery(topStanShape.plate1Shape);
                            plate3Thk = plate1.thickness1;
                            plate3Width = plate1.width1;
                            plate3Len = plate1.length1;

                        }
                        botStanShape.plate2Shape = "";
                        if (!(topStanShape.plate2Shape == "" || topStanShape.plate2Shape == "No Value"))
                        {
                            plate2 = LoadPlateDataByQuery(topStanShape.plate2Shape);
                            plate4Thk = plate2.thickness1;
                            plate4Width = plate2.width1;
                            plate4Len = plate2.length1;
                        }

                        if (HgrCompareDoubleService.cmpdbl(elbowRadius, 0) == true && HgrCompareDoubleService.cmpdbl(faceToCenter, 0) == true)
                        {
                            //Pipe Center to DummyLeg bottom distance
                            topStanShape.stanHeight = diameter1 / 2.0 + topStanShape.stanHeight;
                        }
                        else
                        {
                            topStanShape.stanHeight = topStanShape.stanHeight + diameter1 / 2.0 + faceToCenter;
                        }

                        botomHt = botStanShape.stanHeight;
                        topHT = topStanShape.stanHeight;
                    }
                }
                else if (isStanchion == 2)  //Part has Dummy Leg shape
                {
                    if (!(botshape == "" || botshape == "No Value"))
                    {
                        stanGap = 0;  //Dummy Leg should attach to the pipe

                        if (!(topShape == "" || topShape == "No Value")) //Part has both Top and Bottom Shapes
                        {
                            //Load Bottom and Top Shape's Data
                            botStanShape = LoadStanchShapeDataByQuery(botshape);
                            topDummyLegshape = LoadDummyLegShapeDataByQuery(topShape);

                            //Setting the width,height, and depth in case of a square or rectangular crosssection
                            
                            //If bot and top shapes are not having shape data
                            if (isFixedLeng == false)
                            {
                                if (!(botshape == "" || botshape == "No Value"))
                                {
                                    if (HgrCompareDoubleService.cmpdbl(botStanShape.stanHeight , 0) == true)
                                    {
                                        botStanShape.stanHeight = 0.5;
                                    }
                                }
                            }

                            if (isFixedLeng == false)
                            {
                                if (!(topShape == "" || topShape == "No Value"))
                                {
                                    if (HgrCompareDoubleService.cmpdbl(topDummyLegshape.dummyHeight , 0) == true)
                                    {
                                        topDummyLegshape.dummyHeight = 0.5;
                                    }
                                }
                            }

                            
                            bottomStanhgt = botStanShape.stanHeight;
                            bottomStanWdth = topDummyLegshape.dummyWidth;
                            topDummyDpth = topDummyLegshape.dummyDepth;
                            topSteelName = topDummyLegshape.steelName;
                            topSteelType = topDummyLegshape.steelType;
                            topSteelStd = topDummyLegshape.steelStandard;
                            bottomSteelName = botStanShape.steelName;
                            bottomSteelType = botStanShape.steelType;
                            bottomSteelStd = botStanShape.steelStandard;
                            topDummyType = topDummyLegshape.dummyShape;
                            bottomStanType = botStanShape.stanShape;

                            if (HgrCompareDoubleService.cmpdbl(elbowRadius, 0) == true && HgrCompareDoubleService.cmpdbl(faceToCenter, 0) == true)
                            {
                                //Pipe Center to DummyLeg bottom distance
                                topDummyLegshape.dummyHeight = diameter1 / 2.0 + topDummyLegshape.dummyHeight;
                            }
                            else
                            {
                                topDummyLegshape.dummyHeight = topDummyLegshape.dummyHeight + diameter1 / 2.0 + faceToCenter;
                            }
                            //Set if DummyLeg supporting pipe dia is not given use dia from partclass
                            if (HgrCompareDoubleService.cmpdbl(topDummyLegshape.diameter, 0) == true && HgrCompareDoubleService.cmpdbl(diameter1, 0) == true)
                            {
                                topDummyLegshape.diameter = diameter1;
                            }
                            else if (HgrCompareDoubleService.cmpdbl(topDummyLegshape.diameter, 0) == true && HgrCompareDoubleService.cmpdbl(diameter1, 0) == true)
                            {
                                //MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", "WARNING: " + "Pipe Diameter is required to Calculate DummyLeg Weight", "", "DummyLeg.cs", 1);
                                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrPipeDiaReq, "Pipe Diameter is required to Calculate DummyLeg Weight"));
                            }

                            botomHt = botStanShape.stanHeight;
                            topDummyHt = topDummyLegshape.dummyHeight;

                            //Load Bottom Shape Plates Data
                            if (!(botStanShape.plate1Shape == "" || botStanShape.plate1Shape == "No Value"))
                            {
                                plate1 = LoadPlateDataByQuery(botStanShape.plate1Shape);
                                plate1Thk = plate1.thickness1;
                                plate1Width = plate1.width1;
                                plate1Len = plate1.length1;
                            }
                            if (!(botStanShape.plate2Shape == "" || botStanShape.plate2Shape == "No Value"))
                            {
                                plate2 = LoadPlateDataByQuery(botStanShape.plate2Shape);
                                plate2Thk = plate2.thickness1;
                                plate2Width = plate2.width1;
                                plate2len = plate2.length1;
                            }
                            //Load Top Shape Plate Data
                            if (!(topDummyLegshape.plateShape == "" || topDummyLegshape.plateShape == "No Value"))
                            {
                                plate1 = LoadPlateDataByQuery(topDummyLegshape.plateShape);
                                plate3Thk = plate1.thickness1;
                                plate3Width = plate1.width1;
                                plate3Len = plate1.length1;
                            }
                        }
                        else
                        {
                            botDummylegShape = LoadDummyLegShapeDataByQuery(botshape);

                            if (isFixedLeng == false)
                            {
                                if (!(botshape == "" || botshape == "No Value"))
                                {
                                    if (HgrCompareDoubleService.cmpdbl(botDummylegShape.dummyHeight , 0) == true)
                                    {
                                        botDummylegShape.dummyHeight = 0.5;
                                    }
                                }
                            }

                            bottomDummyWdth = botDummylegShape.dummyWidth;
                            bottomDummyDpth = botDummylegShape.dummyDepth;
                            bottomSteelName = botDummylegShape.steelName;
                            bottomSteelType = botDummylegShape.steelType;
                            bottomSteelStd = botDummylegShape.steelStandard;
                            bottomDummyType = botDummylegShape.dummyShape;

                            if (HgrCompareDoubleService.cmpdbl(elbowRadius, 0) == true)
                            {
                                //Pipe Center to DummyLeg bottom distance
                                botDummylegShape.dummyHeight = (diameter1 / 2) + botDummylegShape.dummyHeight;
                            }
                            else
                            {
                                botDummylegShape.dummyHeight = botDummylegShape.dummyHeight + diameter1 / 2 + faceToCenter;
                            }
                            //Set if DummyLeg supporting pipe dia is not given use dia from partclass
                            if (HgrCompareDoubleService.cmpdbl(botDummylegShape.diameter, 0) == true && HgrCompareDoubleService.cmpdbl(diameter1, 0) == false)
                            {
                                botDummylegShape.diameter = diameter1; //hs_Parts_LoadStanchShapeDataByQuery myBotDummyShape, sBotShape
                            }

                            //Load DummyLeg Plate Data
                            if (!(botDummylegShape.plateShape == "" || botDummylegShape.plateShape == "No Value"))
                            {
                                plate1 = LoadPlateDataByQuery(botDummylegShape.plateShape);
                                plate1Thk = plate1.thickness1;
                                plate1Width = plate1.width1;
                                plate1Len = plate1.length1;
                            }

                            bottomDummyHt = botDummylegShape.dummyHeight; 

                        }
                    }
                }


                bool isFixedLen = true;
                double minLen = 0, maxLen = 0;
                try
                {
                    minLen = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJUAhsMinLen", "MinLen")).PropValue;
                }
                catch
                {
                    minLen = 0;
                }
                try
                {
                    maxLen = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJUAhsMaxLen", "MaxLen")).PropValue;
                }
                catch
                {
                    maxLen = 0;
                }
                try
                {
                    if (catalogPart.SupportsInterface("IJOAHgrOccLength"))
                    {
                        PropertyValueDouble catalogLen = null;
                        try
                        {
                            catalogLen = (PropertyValueDouble)catalogPart.GetPropertyValue("IJOAHgrOccLength", "Length");
                        }
                        catch
                        {
                            catalogLen = null;
                        }
                        if (catalogLen.PropValue == null)
                            length = 0;
                        else
                            length = (double)catalogLen.PropValue;

                        if (HgrCompareDoubleService.cmpdbl(length, 0) == true && (supportComponentBO.SupportsInterface("IJOAHgrOccLength")))
                        {
                            PropertyValueDouble occurenceLen = null;
                            try
                            {
                                occurenceLen = (PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrOccLength", "Length");
                            }
                            catch
                            {
                                occurenceLen = null;
                            }
                            isFixedLen = false;
                            if (occurenceLen.PropValue == null)
                            {
                                length = topDummyHt + topHT + botomHt + bottomDummyHt + plate1Thk + plate2Thk + plate3Thk + plate4Thk;
                            }

                            else
                                length = (double)occurenceLen.PropValue;
                        }
                    }
                    else if (catalogPart.SupportsInterface("IJUAhsLength"))
                    {
                        length = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJUAhsLength", "Length")).PropValue;
                    }
                    else
                    {
                        length = 0;
                    }

                }
                catch
                {
                    length = 0;
                }

                #region // Implementation for Minimum and Maximun Length
                string maxLength = "";
                string minLength = "";
                string value = "";

                //int bomUnits = 0;
                try
                {
                    PropertyValueCodelist bomList = (PropertyValueCodelist)catalogPart.GetPropertyValue("IJUAhsBOMLenUnits", "BOMLenUnits");
                    value = bomList.PropertyInfo.CodeListInfo.GetCodelistItem(bomList.PropValue).DisplayName;
                }
                catch
                {
                    value = "in";
                }

                if (value.ToUpper() == "IN")
                {
                    minLength = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, minLen, UnitName.DISTANCE_INCH);
                    maxLength = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, maxLen, UnitName.DISTANCE_INCH);
                }
                else if (value.ToUpper() == "FT")
                {
                    minLength = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, minLen, UnitName.DISTANCE_FOOT);
                    maxLength = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, maxLen, UnitName.DISTANCE_FOOT);
                }
                else if (value.ToUpper() == "MM")
                {
                    minLength = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, minLen, UnitName.DISTANCE_MILLIMETER);
                    maxLength = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, maxLen, UnitName.DISTANCE_MILLIMETER);
                }
                else if (value.ToUpper() == "M")
                {
                    minLength = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, minLen, UnitName.DISTANCE_METER);
                    maxLength = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, maxLen, UnitName.DISTANCE_METER);
                }

                if (isFixedLen == false)
                {
                    if (length < minLen || length > maxLen)
                    {
                        ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrInvalidLengthOfDummyLeg, "Length of the DummyLeg must be between" + "" + minLength + "and" + maxLength));
                    }
                }

                if (length < minLen)
                    length = minLen;
                if (length > maxLen)
                    length = maxLen;

                try
                {
                    supportComponentBO.SetPropertyValue(length, "IJOAHgrOccLength", "Length");
                }
                catch
                {
                }

                #endregion
                //Obtaining the Cross-section Area in case a Steel Member is used as Leg
                if (bottomSteelName != "" || topSteelName != "")
                {
                    CrossSectionServices crossSectionServices = new CrossSectionServices();

                    CrossSection crossSection;

                    if (bottomSteelName != "")
                    {

                        crossSection = catalogStructHelper.GetCrossSection(bottomSteelStd, bottomSteelType, bottomSteelName);
                    }
                    else
                    {
                        crossSection = catalogStructHelper.GetCrossSection(topSteelStd, topSteelType, topSteelName);
                    }
                    try
                    {
                        areaCS = (double)((PropertyValueDouble)crossSection.GetPropertyValue("IStructCrossSectionDimensions", "Area")).PropValue;
                    }
                    catch
                    {
                        areaCS = 0;
                    }
                }

                //Obtaining the Material Density
                Material material;
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


                //Calculating the Plate Volumes
                if (HgrCompareDoubleService.cmpdbl(plate1Thk , 0)== false && HgrCompareDoubleService.cmpdbl(plate1Width , 0)== false && HgrCompareDoubleService.cmpdbl(plate1Len , 0)== false)
                {
                    plate1Vol = plate1Thk * plate1Width * plate1Len;
                }
                if (HgrCompareDoubleService.cmpdbl(plate2Thk, 0) == false && HgrCompareDoubleService.cmpdbl(plate2Width, 0) == false && HgrCompareDoubleService.cmpdbl(plate2len, 0) == false)
                {
                    plate2Vol = plate2Thk * plate2Width * plate2len;
                }
                if (HgrCompareDoubleService.cmpdbl(plate3Thk, 0) == false && HgrCompareDoubleService.cmpdbl(plate3Width, 0) == false && HgrCompareDoubleService.cmpdbl(plate3Len, 0) == false)
                {
                    plate3Vol = plate3Thk * plate3Width * plate3Len;
                }
                if (HgrCompareDoubleService.cmpdbl(plate4Thk, 0) == false && HgrCompareDoubleService.cmpdbl(plate4Width, 0) == false && HgrCompareDoubleService.cmpdbl(plate4Len, 0) == false)
                {
                    plate4Vol = plate4Thk * plate4Width * plate4Len;
                }
                totPlateThk = plate1Thk + plate2Thk + plate3Thk + plate4Thk;
                totPlateVol = plate1Vol + plate2Vol + plate3Vol + plate4Vol;

                if (isFixedLen == false)
                {
                    //Calculating the Leg Volume in case of a Variable Height DummyLeg, and a Square or Rectangular Leg cross-section
                    if (topStanType == 2)
                    {
                        if (HgrCompareDoubleService.cmpdbl(topStanDpth , 0)==false  && isStanchion == 1)
                        {
                            topLegVol = (topStanDpth * topStanWdth * length);
                        }
                        else if (HgrCompareDoubleService.cmpdbl(topStanDpth , 0)==true  && isStanchion == 1)
                        {
                            topLegVol = (topStanWdth * topStanWdth * length);
                        }
                    }
                    if (bottomStanType == 2)
                    {
                        if (HgrCompareDoubleService.cmpdbl(bottomStanDpth, 0)==false )
                        {
                            bottomLegVol = (bottomStanDpth * bottomStanWdth * length);
                        }
                        else
                        {
                            bottomLegVol = (bottomStanWdth * bottomStanWdth * length);
                        }
                    }
                    if (topDummyType == 2)
                    {
                        if (HgrCompareDoubleService.cmpdbl(topDummyDpth, 0)==false  && isStanchion == 2)
                        {
                            topLegVol = (bottomStanDpth * bottomStanWdth * length);
                        }
                        else if (HgrCompareDoubleService.cmpdbl(topDummyDpth, 0)==true && isStanchion == 1)
                        {
                            topLegVol = (bottomStanWdth * bottomStanWdth * length);
                        }
                    }
                    if (bottomDummyType == 2)
                    {
                        if (HgrCompareDoubleService.cmpdbl(bottomDummyDpth, 0)==false  && isStanchion == 2)
                        {
                            bottomLegVol = (bottomDummyDpth * bottomDummyWdth * length);
                        }
                        else if (HgrCompareDoubleService.cmpdbl(bottomDummyDpth, 0)==true && isStanchion == 1)
                        {
                            bottomLegVol = (bottomDummyWdth * bottomDummyWdth * length);
                        }
                    }

                }
                else
                {
                    if (topStanType == 2)
                    {
                        //Calculating the Leg Volume in case of a Constant Height Dummy Leg, and a Square or Rectangular Leg cross-section
                        if (HgrCompareDoubleService.cmpdbl(topStanDpth , 0)==false  && isStanchion == 1)
                        {
                            topLegVol = (topStanDpth * topStanWdth * topStanHt);
                        }
                        else if (HgrCompareDoubleService.cmpdbl(topStanDpth , 0)==true  && isStanchion == 1)
                        {
                            topLegVol = (topStanWdth * topStanWdth * topStanHt);
                        }
                    }
                    if (bottomStanType == 2)
                    {
                        if (HgrCompareDoubleService.cmpdbl(bottomStanDpth, 0)==false )
                        {
                            bottomLegVol = (bottomStanDpth * bottomStanWdth * bottomStanhgt);
                        }
                        else
                        {
                            bottomLegVol = (bottomStanWdth * bottomStanWdth * bottomStanhgt);
                        }
                    }
                    if (topDummyType == 2)
                    {
                        if (HgrCompareDoubleService.cmpdbl(topDummyDpth, 0)==false  && isStanchion == 2)
                        {
                            topLegVol = (bottomStanDpth * bottomStanWdth * topDummyHt);
                        }
                        else if (HgrCompareDoubleService.cmpdbl(topDummyDpth, 0)==true && isStanchion == 1)
                        {
                            topLegVol = (bottomStanWdth * bottomStanWdth * topDummyHt);
                        }
                    }
                    if (bottomDummyType == 2)
                    {
                        if (HgrCompareDoubleService.cmpdbl(bottomDummyDpth, 0)==false  && isStanchion == 2)
                        {
                            bottomLegVol = (bottomDummyDpth * bottomDummyWdth * bottomDummyHt);
                        }
                        else if (HgrCompareDoubleService.cmpdbl(bottomDummyDpth, 0)==true && isStanchion == 1)
                        {
                            bottomLegVol = (bottomDummyWdth * bottomDummyWdth * bottomDummyHt);
                        }
                    }
                }

                if (isFixedLen == false)
                {
                    //Calculating the Leg Volume in case of a Variable Height DummyLeg, and a Circular Leg cross-section, Steel Member as a Leg
                    if (topStanType == 1 && HgrCompareDoubleService.cmpdbl(topStanDpth , 0)==true  && HgrCompareDoubleService.cmpdbl(topStanWdth, 0)==false && isStanchion == 1)
                    {
                        topLegVol = (length * (Math.PI) * topStanWdth * topStanWdth / 4);
                    }
                    if (bottomStanType == 1 && HgrCompareDoubleService.cmpdbl(bottomStanDpth, 0)==true && HgrCompareDoubleService.cmpdbl(bottomStanWdth, 0)==false )
                    {
                        bottomLegVol = (length * Math.PI * bottomStanWdth * bottomStanWdth / 4);
                    }

                    if (topDummyType == 1 && HgrCompareDoubleService.cmpdbl(topDummyDpth, 0)==true && HgrCompareDoubleService.cmpdbl(topDummyWdth, 0)==true && isStanchion == 2)
                    {
                        topLegVol = (length * Math.PI * topDummyWdth * topDummyWdth / 4);
                    }

                    if (bottomDummyType == 1 && HgrCompareDoubleService.cmpdbl(bottomDummyDpth, 0)==true && HgrCompareDoubleService.cmpdbl(bottomDummyWdth, 0)==false  && isStanchion == 2)
                    {
                        bottomLegVol = (length * Math.PI * bottomDummyWdth * bottomDummyWdth / 4);
                    }


                    if (topStanType == 3 && HgrCompareDoubleService.cmpdbl(topStanDpth , 0)==true  && HgrCompareDoubleService.cmpdbl(topStanWdth, 0)==true && isStanchion == 1)
                    {
                        topLegVol = (length * areaCS);
                    }

                    if (bottomStanType == 3 && HgrCompareDoubleService.cmpdbl(bottomStanDpth, 0)==true && HgrCompareDoubleService.cmpdbl(bottomStanWdth, 0)==true)
                    {
                        bottomLegVol = (length * areaCS);
                    }

                    if (topDummyType == 3 && HgrCompareDoubleService.cmpdbl(topDummyDpth, 0)==true && HgrCompareDoubleService.cmpdbl(topDummyWdth, 0)==false  && isStanchion == 2)
                    {
                        topLegVol = (length * areaCS);
                    }

                    if (bottomDummyType == 3 && HgrCompareDoubleService.cmpdbl(bottomDummyDpth, 0)==true && HgrCompareDoubleService.cmpdbl(bottomDummyWdth, 0)==false  && isStanchion == 2)
                    {
                        bottomLegVol = (length * areaCS);
                    }
                }
                else
                {
                    //Calculating the Leg Volume in case of a Constant Height DummyLeg, and a Circular Leg cross-section, Steel Member as a Leg
                    if (topStanType == 1 && HgrCompareDoubleService.cmpdbl(topStanDpth , 0)==true  && HgrCompareDoubleService.cmpdbl(topStanWdth, 0)==false && isStanchion == 1)
                    {
                        topLegVol = (topStanHt * (Math.PI) * topStanWdth * topStanWdth / 4);
                    }
                    if (bottomStanType == 1 && HgrCompareDoubleService.cmpdbl(bottomStanDpth, 0)==true && HgrCompareDoubleService.cmpdbl(bottomStanWdth, 0)==false )
                    {
                        bottomLegVol = (bottomStanhgt * Math.PI * bottomStanWdth * bottomStanWdth / 4);
                    }

                    if (topDummyType == 1 && HgrCompareDoubleService.cmpdbl(topDummyDpth, 0)==true && HgrCompareDoubleService.cmpdbl(topDummyWdth, 0)==true && isStanchion == 2)
                    {
                        topLegVol = (topDummyHt * Math.PI * topDummyWdth * topDummyWdth / 4);
                    }

                    if (bottomDummyType == 1 && HgrCompareDoubleService.cmpdbl(bottomDummyDpth, 0)==true && HgrCompareDoubleService.cmpdbl(bottomDummyWdth, 0)==false  && isStanchion == 2)
                    {
                        bottomLegVol = (bottomDummyHt * Math.PI * bottomDummyWdth * bottomDummyWdth / 4);
                    }

                    //Leg Volume in Case of a Steel Member
                    if (topStanType == 3 && HgrCompareDoubleService.cmpdbl(topStanDpth , 0)==true  && HgrCompareDoubleService.cmpdbl(topStanWdth, 0)==true && isStanchion == 1)
                    {
                        topLegVol = (topStanHt * areaCS);
                    }

                    if (bottomStanType == 3 && HgrCompareDoubleService.cmpdbl(bottomStanDpth, 0)==true && HgrCompareDoubleService.cmpdbl(bottomStanWdth, 0)==true)
                    {
                        bottomLegVol = (bottomStanhgt * areaCS);
                    }

                    if (topDummyType == 3 && HgrCompareDoubleService.cmpdbl(topDummyDpth, 0)==true && HgrCompareDoubleService.cmpdbl(topDummyWdth, 0)==false  && isStanchion == 2)
                    {
                        topLegVol = (topDummyHt * areaCS);
                    }

                    if (bottomDummyType == 3 && HgrCompareDoubleService.cmpdbl(bottomDummyDpth, 0)==true && HgrCompareDoubleService.cmpdbl(bottomDummyWdth, 0)==false  && isStanchion == 2)
                    {
                        bottomLegVol = (bottomDummyHt * areaCS);
                    }
                }

                totLegVol = topLegVol + bottomLegVol;

                //The Total Volume is the Sum of Leg Volume and the Plate Volume of the Dummy Leg
                totVol = totPlateVol + totLegVol;

                double wtPerLen = 0;
                try
                {
                    if (catalogPart.SupportsInterface("IJUAhsWtPerLen"))
                    {
                        wtPerLen = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJUAhsWtPerLen", "WtPerLen")).PropValue;
                    }
                }
                catch
                {
                    wtPerLen = 0;
                }
                
                if (HgrCompareDoubleService.cmpdbl(dryWeight, 0)==false )
                {
                    //Weight
                    weight = dryWeight;
                }
                else if (wtPerLen > 0)
                {
                    double extraLenToPipeCL = 0;
                    if (HgrCompareDoubleService.cmpdbl(elbowRadius, 0)==true)
                    {
                        //Pipe Center to DummyLeg bottom distance
                        extraLenToPipeCL = (diameter1 / 2);
                    }
                    else
                    {
                        extraLenToPipeCL =  diameter1 / 2 + faceToCenter;
                    }

                    weight = (totPlateVol * materialDensity) + ((length - totPlateThk - extraLenToPipeCL) * wtPerLen);
                }
                else
                {
                    //The weight of the Dummy Leg is obtained by multiplying Volume with Material Density
                    weight = (totVol * materialDensity);
                }

                //Center of Gravity
                if (HgrCompareDoubleService.cmpdbl(dryCogX, 0)==false )
                {
                    cogX = dryCogX;
                }
                else
                {
                    cogX = 0;
                }
                if (HgrCompareDoubleService.cmpdbl(dryCogy, 0)==false )
                {
                    cogY = dryCogy;
                }
                else
                {
                    cogY = 0;
                }
                if (HgrCompareDoubleService.cmpdbl(dryCogZ, 0)==false )
                {
                    cogZ = dryCogZ;
                }
                else
                {
                    cogZ = 0;
                }

                SupportComponent supportComponent = (SupportComponent)supportComponentBO;
                supportComponent.SetWeightAndCOG(weight, cogX, cogY, cogZ);
            }

            catch
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrWeightCG, "Error in Weight CG of DummyLeg"));
                }
            }

        }
        
    }
}
        #endregion
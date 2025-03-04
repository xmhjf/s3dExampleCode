//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   ConstantTypeG.cs
//    HSSmartPart,Ingr.SP3D.Content.Support.Symbols.ConstantTypeG
//   Author       :  Vijay
//   Creation Date: 06.June.2013 
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   06.June.2013    Vijay   CR-CP-222480  Convert HS_S3DConstant Smartpart VB Project to C# .Net  
//   03-09-2014     B Chethan DI-CP-253819  Create the missing Constant and Variable springs for Seonghwa / Wookwang catalog  
//   12-12-2014      PVK     TR-CP-264951	Resolve P3 coverity issues found in November 2014 report
//   28-04-2015      PVK	 Resolve Coverity issues found in April
//   26-05-2015      PVK	 TR-CP-262762	Constant Spring SmartPart fails with incorrect message
//   10-06-2015      PVK	 TR-CP-274155	SmartPart TDL Errors should be corrected.
//   16-07-2015      PVK     Resolve coverity issues found in July 2015 report
//   30-11=2015      VDP     Integrate the newly developed SmartParts into Product(DI-CP-282644)
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.ReferenceData.Middle.Services;
using Ingr.SP3D.Structure.Middle.Services;
using Ingr.SP3D.Support.Middle;

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

    [CacheOption(CacheOptionType.NonCached)]
    [SymbolVersion("1.0.0.0")]
    [VariableOutputs]
    public class ConstantTypeG : SmartPartComponentDefinition, ICustomWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "HSSmartPart,Ingr.SP3D.Content.Support.Symbols.ConstantTypeG"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        #region "Definition of Additional Inputs"

        public override IEnumerable<Input> AdditionalInputs
        {
            get
            {
                int endIndex = 2;
                List<Input> additionalInputs = new List<Input>();

                //Add Reference Table Names as Inputs
                additionalInputs.Add(new InputString(endIndex, "SizeTable", "SizeTable", "No Value", false));
                additionalInputs.Add(new InputString(++endIndex, "TravelTable", "TravelTable", "No Value", false));
                additionalInputs.Add(new InputString(++endIndex, "LoadTable", "LoadTable", "No Value", false));
                additionalInputs.Add(new InputString(++endIndex, "SelectionTable", "SelectionTable", "No Value", false));

                //Add Data Inputs
                additionalInputs.Add(new InputDouble(++endIndex, "Load", "Load", 0, false));
                additionalInputs.Add(new InputDouble(++endIndex, "Movement", "Movement", 0, false));
                additionalInputs.Add(new InputDouble(++endIndex, "MovementDirection", "MovementDirection", 1, false));
                additionalInputs.Add(new InputDouble(++endIndex, "TTIncrement", "TTIncrement", 0, false));
                additionalInputs.Add(new InputString(++endIndex, "TTUnits", "TTUnits", "No Value", false));
                additionalInputs.Add(new InputDouble(++endIndex, "MinTT", "MinTT", 0, false));
                additionalInputs.Add(new InputDouble(++endIndex, "MaxTT", "MaxTT", 0, false));
                additionalInputs.Add(new InputDouble(++endIndex, "OverTravelMin", "OverTravelMin", 0, false));
                additionalInputs.Add(new InputDouble(++endIndex, "OverTravelPercent", "OverTravelPercent", 0, false));
                additionalInputs.Add(new InputString(++endIndex, "Size", "Size", "No Value", false));
                additionalInputs.Add(new InputDouble(++endIndex, "PipeOD", "PipeOD", 0, false));
                additionalInputs.Add(new InputDouble(++endIndex, "ShoeHeight", "ShoeHeight", 0, false));
                additionalInputs.Add(new InputDouble(++endIndex, "BBGap", "BBGap", 0, false));
                additionalInputs.Add(new InputDouble(++endIndex, "CC", "CC", 0, false));

                //Add Configuration Input
                additionalInputs.Add(new InputDouble(++endIndex, "ConstantConfig", "ConstantConfig", 1, false));
                
                //Add Travel Increment Rule Input
                additionalInputs.Add(new InputString(++endIndex, "TTIncrementByRule", "TTIncrementByRule", "No Value",true));

                //Add Valid Travel Table Input
                additionalInputs.Add(new InputString(++endIndex, "ValidTravelList", "ValidTravelList", "No Value", true));
                
                return additionalInputs;
            }
        }
        #endregion
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Route", "Route")]
        [SymbolOutput("Steel", "Steel")]
        [SymbolOutput("RodEnd1", "RodEnd1")]
        [SymbolOutput("RodEnd2", "RodEnd2")]
        [SymbolOutput("Surface1", "Surface1")]
        [SymbolOutput("Surface2", "Surface2")]
        public AspectDefinition m_PhysicalAspect;

        #endregion

        #region "Definition of Additional Outputs"

        public override IEnumerable<OutputDefinition> AdditionalOutputs(string aspectName)
        {
            List<OutputDefinition> additionalOutputs = new List<OutputDefinition>();
            if (aspectName == "Symbolic")
            {
                //Add Outputs for the Plates
                AddPlateOutputs(3, additionalOutputs);

                //Add Outputs for the Casing and Frame
                additionalOutputs.Add(new OutputDefinition("Casing", "Casing"));
                additionalOutputs.Add(new OutputDefinition("FrameA", "FrameA"));
                additionalOutputs.Add(new OutputDefinition("FrameB", "FrameB"));
                additionalOutputs.Add(new OutputDefinition("LeverA", "LeverA"));
                additionalOutputs.Add(new OutputDefinition("LeverB", "LeverB"));

                //Add Outputs for the Load Column, and the Column End Shape
                AddRod1Outputs("Column", additionalOutputs);
                additionalOutputs.Add(new OutputDefinition("ColumnPlate", "ColumnPlate"));
                additionalOutputs.Add(new OutputDefinition("ColumnEnd", "ColumnEnd"));
                additionalOutputs.Add(new OutputDefinition("ColumnNut", "ColumnNut"));

                AddClevisOutputs(aspectName, additionalOutputs);
                AddTurnbuckleOutputs(additionalOutputs);
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
                StandardComponent standardComponent = Occurrence as StandardComponent;
                if (standardComponent != null)
                {
                    if ((standardComponent).Name != null)
                        SmartPartComponentDefinition.isSpringPlacedFirstTime = false;
                    else
                        SmartPartComponentDefinition.isSpringPlacedFirstTime = true;
                }

                int startIndex = 2;
                Matrix4X4 matrix = new Matrix4X4();
                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();

                string sizeTable = GetStringInputValue(startIndex);
                string travelTable = GetStringInputValue(++startIndex);
                string loadTable = GetStringInputValue(++startIndex);
                string selectionTable = GetStringInputValue(++startIndex);
                Double load = GetDoubleInputValue(++startIndex);
                Double movement = GetDoubleInputValue(++startIndex);
                int movementDirection = (int)GetDoubleInputValue(++startIndex);
                Double totalTravelIncrement = GetDoubleInputValue(++startIndex);
                string totalTravelUnits = GetStringInputValue(++startIndex);
                Double minTT = GetDoubleInputValue(++startIndex);
                Double maxTT = GetDoubleInputValue(++startIndex);
                Double minOT = GetDoubleInputValue(++startIndex);
                Double percentOT = GetDoubleInputValue(++startIndex);
                string size = GetStringInputValue(++startIndex);
                Double pipeOuterDiameter = GetDoubleInputValue(++startIndex);
                Double shoeHeight = GetDoubleInputValue(++startIndex);
                Double bBGap = GetDoubleInputValue(++startIndex);
                Double CC = GetDoubleInputValue(++startIndex);
                int config = (int)GetDoubleInputValue(++startIndex);

                string TTIncrementByRule = String.Empty;

                if (part.SupportsInterface("IJUAhsTTIncrementByRule"))
                    TTIncrementByRule = GetStringInputValue(++startIndex);
                else
                {
                    TTIncrementByRule = "No Value";
                    ++startIndex;
                }

                string ValidTravelTable = String.Empty;

                if (part.SupportsInterface("IJUAhsValidTravels"))
                    ValidTravelTable = GetStringInputValue(++startIndex);
                else
                {
                    ValidTravelTable = "No Value";
                    ++startIndex;
                }

                if (TTIncrementByRule != "No Value")
                {
                    if (standardComponent != null)
                    {
                        if ((standardComponent).Name != null)
                        {
                            RelationCollection hgrRelation = standardComponent.GetRelationship("SupportHasComponents", "Support");
                            BusinessObject businessObject = hgrRelation.TargetObjects[0];
                            GenericHelper genericHelper = new GenericHelper((Ingr.SP3D.Support.Middle.Support)businessObject);
                            Collection<object> collection = new Collection<object>();

                            genericHelper.GetDataByRule(TTIncrementByRule, standardComponent, out collection);
                            totalTravelIncrement = (double)collection[0];
                        }
                        else
                            totalTravelIncrement = 10;
                    }
                }

                //Determine the Total Travel
                Double totalTravel;
                percentOT = percentOT / 100;
                if (movement * percentOT > minOT)
                    totalTravel = movement * (1 + percentOT);
                else
                    totalTravel = movement + minOT;

                //Round to the nearest increment
                Double totalTravelInUnits;
                if (totalTravelUnits.Equals("in"))
                {
                    totalTravel = MiddleServiceProvider.UOMMgr.ConvertDBUtoUnit(UnitType.Distance, totalTravel, UnitName.DISTANCE_INCH);
                    totalTravelInUnits = Math.Ceiling(totalTravel / totalTravelIncrement) * totalTravelIncrement;
                }
                else
                {
                    totalTravel = MiddleServiceProvider.UOMMgr.ConvertDBUtoUnit(UnitType.Distance, totalTravel, UnitName.DISTANCE_MILLIMETER);
                    totalTravelInUnits = Math.Ceiling(totalTravel / totalTravelIncrement) * totalTravelIncrement;
                }

                if (ValidTravelTable != "No Value")
                {
                    // Get the list of valid Total Travels for the constant
                    try
                    {
                        // Get Available Total Travels from codelist table
                        MetadataManager metaDataMgr = MiddleServiceProvider.SiteMgr.ActiveSite.ActivePlant.PlantCatalog.MetadataMgr;

                        CodelistInformation codeListInfo = metaDataMgr.GetCodelistInfo(ValidTravelTable, "UDP");
                        List<CodelistItem> validTravels = codeListInfo.CodelistMembers;

                        // Removed the Undefined codelist item from the list
                        validTravels.Remove(codeListInfo.GetCodelistItem(-1));

                        // Make sure codelist is sorted in ascending order
                        validTravels.Sort((x, y) => Convert.ToDouble(x.ShortDisplayName).CompareTo(Convert.ToDouble(y.ShortDisplayName)));

                        foreach (CodelistItem travel in validTravels)
                        {
                            if (Math.Round(totalTravel, 5) <= Math.Round(Convert.ToDouble(travel.ShortDisplayName), 5))
                            {
                                totalTravelInUnits = Convert.ToDouble(travel.ShortDisplayName);
                                break;
                            }
                        }
                    }
                    catch
                    { }
                }

                if (Math.Round(totalTravelInUnits, 1) < minTT)
                {
                    totalTravelInUnits = minTT;
                }
                if (Math.Round(totalTravelInUnits, 1) > maxTT)
                {
                    RaiseSmartPartTDLError(SmartPartSymbolResourceIDs.ErrInvTotalTravelUnits, "Total Travel is greater then the maximum allowable value for a size " + size + " Constant.");
                    return;
                }
                if (totalTravelUnits.Equals("in"))
                    totalTravel = MiddleServiceProvider.UOMMgr.ConvertUnitToDBU(UnitType.Distance, totalTravelInUnits, UnitName.DISTANCE_INCH);
                else
                    totalTravel = MiddleServiceProvider.UOMMgr.ConvertUnitToDBU(UnitType.Distance, totalTravelInUnits, UnitName.DISTANCE_MILLIMETER);

                //Load the Constant Reference Data
                ConstantInputs constant = LoadConstantData(sizeTable, travelTable, loadTable, size, Math.Round(totalTravelInUnits, 1), load);
                //Determine the Angle of the Lever Arm and the Rod Take Out
                Double leverAngle, rodTakeout;

                if(movementDirection == 1)
                {
                    leverAngle = constant.angleLow;
                    rodTakeout = constant.rodTakeOutLow;
                }
                else
                {
                    leverAngle = -constant.angleHigh;
                    rodTakeout = constant.rodTakeOutHigh;
                }

                //Determine the length of the column
                switch (config)
                {
                    case 1:
                        constant.length1 = rodTakeout - constant.offset5 + constant.pivotOffset - constant.length2 * Math.Sin(leverAngle) + constant.offset2;
                        break;
                    case 2:
                        constant.length1 = -constant.offset5 - constant.frame.length1 - constant.plate3.thickness1 + constant.pivotOffset - constant.length2 * Math.Sin(leverAngle) + constant.offset2 + rodTakeout;
                        break;
                    case 3:
                        RaiseSmartPartTDLError(SmartPartSymbolResourceIDs.ErrGConfig, "Horizontal Configuration not valid for Type G Constants.");
                        return;
                }

                //Set the Output Part Occurences
                try
                {
                    SupportComponent supportComponent = (SupportComponent)Occurrence;
                    supportComponent.SetPropertyValue(rodTakeout, "IJUAhsRodTakeOut", "RodTakeOut");
                    supportComponent.SetPropertyValue(totalTravel, "IJUAhsTotalTravel", "TotalTravel");
                    supportComponent.SetPropertyValue(constant.rodDiameter, "IJUAhsRodDiameter", "RodDiameter");
                }
                catch{ }
                //set CC Value from load or travel sheet
                if (HgrCompareDoubleService.cmpdbl(CC , 0)==true)
                    CC = constant.cCMin;

                //Add the Ports
                Port port;
                switch (config)
                {
                    case 1:
                        port = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, pipeOuterDiameter / 2 + shoeHeight), new Vector(0, 1, 0), new Vector(0, 0, 1));
                        m_PhysicalAspect.Outputs["Route"] = port;

                        port = new Port(OccurrenceConnection, part, "Steel", new Position(0, 0, 0), new Vector(0, 1, 0), new Vector(0, 0, 1));
                        m_PhysicalAspect.Outputs["Steel"] = port;

                        port = new Port(OccurrenceConnection, part, "RodEnd1", new Position(-CC / 2, 0, rodTakeout), new Vector(1, 0, 0), new Vector(0, 0, 1));
                        m_PhysicalAspect.Outputs["RodEnd1"] = port;

                        port = new Port(OccurrenceConnection, part, "Surface1", new Position(-CC / 2, 0, rodTakeout + constant.offset2), new Vector(1, 0, 0), new Vector(0, 0, 1));
                        m_PhysicalAspect.Outputs["Surface1"] = port;

                        port = new Port(OccurrenceConnection, part, "RodEnd2", new Position(CC / 2, 0, rodTakeout), new Vector(1, 0, 0), new Vector(0, 0, 1));
                        m_PhysicalAspect.Outputs["RodEnd2"] = port;

                        port = new Port(OccurrenceConnection, part, "Surface2", new Position(CC / 2, 0, rodTakeout + constant.offset2), new Vector(1, 0, 0), new Vector(0, 0, 1));
                        m_PhysicalAspect.Outputs["Surface2"] = port;
                        break;
                    case 2:
                        port = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, pipeOuterDiameter / 2 + shoeHeight), new Vector(0, 1, 0), new Vector(0, 0, 1));
                        m_PhysicalAspect.Outputs["Route"] = port;

                        port = new Port(OccurrenceConnection, part, "Steel", new Position(0, 0, 0), new Vector(0, 1, 0), new Vector(0, 0, 1));
                        m_PhysicalAspect.Outputs["Steel"] = port;

                        port = new Port(OccurrenceConnection, part, "RodEnd1", new Position(-CC / 2, 0, constant.offset5 + constant.frame.length1 + constant.plate3.thickness1 - constant.pivotOffset + constant.length2 * Math.Sin(leverAngle) + constant.length1 - constant.offset2), new Vector(1, 0, 0), new Vector(0, 0, 1));
                        m_PhysicalAspect.Outputs["RodEnd1"] = port;

                        port = new Port(OccurrenceConnection, part, "Surface1", new Position(-CC / 2, 0, constant.offset5 + constant.frame.length1 + constant.plate3.thickness1 - constant.pivotOffset + constant.length2 * Math.Sin(leverAngle) + constant.length1), new Vector(1, 0, 0), new Vector(0, 0, 1));
                        m_PhysicalAspect.Outputs["Surface1"] = port;

                        port = new Port(OccurrenceConnection, part, "RodEnd2", new Position(CC / 2, 0, constant.offset5 + constant.frame.length1 + constant.plate3.thickness1 - constant.pivotOffset + constant.length2 * Math.Sin(leverAngle) + constant.length1 - constant.offset2), new Vector(1, 0, 0), new Vector(0, 0, 1));
                        m_PhysicalAspect.Outputs["RodEnd2"] = port;

                        port = new Port(OccurrenceConnection, part, "Surface2", new Position(CC / 2, 0, constant.offset5 + constant.frame.length1 + constant.plate3.thickness1 - constant.pivotOffset + constant.length2 * Math.Sin(leverAngle) + constant.length1), new Vector(1, 0, 0), new Vector(0, 0, 1));
                        m_PhysicalAspect.Outputs["Surface2"] = port;
                        break;
                }

                CrossSection crossSection1, crossSection2;
                CatalogStructHelper catalogStructHelper = new CatalogStructHelper();
                int currentStandard;
                if (int.TryParse(constant.steel.sectionStandard, out currentStandard))
                {
                    MetadataManager metadataManager = MiddleServiceProvider.SiteMgr.ActiveSite.ActivePlant.PlantModel.MetadataMgr;
                    string strCurrentStandard = string.Empty;
                    if (metadataManager != null)
                        strCurrentStandard = metadataManager.GetCodelistInfo("hsSteelStandard", "UDP").GetCodelistItem(currentStandard).ShortDisplayName.Trim();
                    constant.steel.sectionStandard = strCurrentStandard;
                }
                if (!(constant.steel.sectionStandard == "" || constant.steel.sectionStandard == null || constant.steel.sectionStandard == "No Value" && constant.steel.sectionType == "" || constant.steel.sectionType == null || constant.steel.sectionType == "No Value" && constant.steel.sectionName == "" || constant.steel.sectionName == null || constant.steel.sectionName == "No Value"))
                {
                    try
                    {
                        crossSection1 = catalogStructHelper.GetCrossSection(constant.steel.sectionStandard, constant.steel.sectionType, constant.steel.sectionName);
                    }
                    catch
                    {
                        base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, constant.steel.sectionName + " " + SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrSectionNotFound, "section not found in catalog."));
                        return;
                    }
                }
                else
                {
                    try
                    {
                        crossSection1 = (CrossSection)part.GetRelationship("HgrCrossSection", "CrossSection").TargetObjects.First();
                    }
                    catch
                    {
                        RaiseSmartPartTDLError(SmartPartSymbolResourceIDs.ErrCrossSectionNotFound, "Could not get Cross-section object.");
                        return;
                    }
                }

                if (crossSection1 == null)
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrCrossSectionStellDetails, "Steel Details are not found."));

                Double depth = crossSection1.Depth;
                Double width = crossSection1.Width;

                Boolean withTurnbuckle;
                if (HgrCompareDoubleService.cmpdbl(constant.turnbuckle.Diameter1 , 0)==true && HgrCompareDoubleService.cmpdbl(constant.turnbuckle.Length2 , 0)==true)
                    withTurnbuckle = false;
                else
                    withTurnbuckle = true;

                Double turnbuckleLength = constant.turnbuckle.Opening1 + constant.turnbuckle.Nut.ShapeLength * 2;
                Double turnbuckleTakeOut = turnbuckleLength - 2 * constant.offset2;

                switch (config)
                {
                    case 1:
                        //Frame
                        matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0));
                        matrix.Rotate(Math.PI, new Vector(0, 1, 0));
                        matrix.Translate(new Vector(CC / 2 + constant.length2 * Math.Cos(leverAngle) + constant.caseCLOffset + constant.casing.ShapeWidth1 / 2 + constant.frameOffset, constant.frameSpacing / 2 + constant.frame.thickness1, constant.offset5 - constant.plate2.thickness1));
                        AddPlate(constant.frame, matrix, m_PhysicalAspect.Outputs, "Frame1A");

                        matrix = new Matrix4X4();
                        matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0));
                        matrix.Rotate(Math.PI, new Vector(0, 1, 0));
                        matrix.Translate(new Vector(CC / 2 + constant.length2 * Math.Cos(leverAngle) + constant.caseCLOffset + constant.casing.ShapeWidth1 / 2 + constant.frameOffset, -constant.frameSpacing / 2, constant.offset5 - constant.plate2.thickness1));
                        AddPlate(constant.frame, matrix, m_PhysicalAspect.Outputs, "Frame1B");

                        matrix = new Matrix4X4();
                        matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0));
                        matrix.Rotate(Math.PI, new Vector(0, 1, 0));
                        matrix.Rotate(Math.PI, new Vector(0, 0, 1));
                        matrix.Translate(new Vector(-CC / 2 - constant.length2 * Math.Cos(leverAngle) - constant.caseCLOffset - constant.casing.ShapeWidth1 / 2 - constant.frameOffset, constant.frameSpacing / 2, constant.offset5 - constant.plate2.thickness1));
                        AddPlate(constant.frame, matrix, m_PhysicalAspect.Outputs, "Frame2A");

                        matrix = new Matrix4X4();
                        matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0));
                        matrix.Rotate(Math.PI, new Vector(0, 1, 0));
                        matrix.Rotate(Math.PI, new Vector(0, 0, 1));
                        matrix.Translate(new Vector(-CC / 2 - constant.length2 * Math.Cos(leverAngle) - constant.caseCLOffset - constant.casing.ShapeWidth1 / 2 - constant.frameOffset, -constant.frameSpacing / 2 - constant.frame.thickness1, constant.offset5 - constant.plate2.thickness1));
                        AddPlate(constant.frame, matrix, m_PhysicalAspect.Outputs, "Frame2B");

                        //Plates
                        matrix = new Matrix4X4();
                        matrix.Translate(new Vector(-CC / 2 - constant.length2 * Math.Cos(leverAngle) - constant.caseCLOffset - constant.casing.ShapeWidth1 / 2 + constant.plate1Offset, -constant.plate1.length1 / 2, constant.offset5 - constant.plate2.thickness1 - constant.frame.length1 - constant.plate3.thickness1 + constant.length - constant.plate1.thickness1));
                        AddPlate(constant.plate1, matrix, m_PhysicalAspect.Outputs, "Plate1A");

                        matrix = new Matrix4X4();
                        matrix.Translate(new Vector(CC / 2 + constant.length2 * Math.Cos(leverAngle) + constant.caseCLOffset + constant.casing.ShapeWidth1 / 2 - constant.plate1Offset - constant.plate1.width1, -constant.plate1.length1 / 2, constant.offset5 - constant.plate2.thickness1 - constant.frame.length1 - constant.plate3.thickness1 + constant.length - constant.plate1.thickness1));
                        AddPlate(constant.plate1, matrix, m_PhysicalAspect.Outputs, "Plate1B");

                        matrix = new Matrix4X4();
                        matrix.Translate(new Vector(-CC / 2 - constant.length2 * Math.Cos(leverAngle) - constant.caseCLOffset - constant.casing.ShapeWidth1 / 2 + constant.plate2Offset, -constant.plate2.length1 / 2, constant.offset5 - constant.plate2.thickness1));
                        AddPlate(constant.plate2, matrix, m_PhysicalAspect.Outputs, "Plate2A");

                        matrix = new Matrix4X4();
                        matrix.Translate(new Vector(CC / 2 + constant.length2 * Math.Cos(leverAngle) + constant.caseCLOffset + constant.casing.ShapeWidth1 / 2 - constant.plate2Offset - constant.plate2.width1, -constant.plate2.length1 / 2, constant.offset5 - constant.plate2.thickness1));
                        AddPlate(constant.plate2, matrix, m_PhysicalAspect.Outputs, "Plate2B");

                        matrix = new Matrix4X4();
                        matrix.Translate(new Vector(-CC / 2 - constant.length2 * Math.Cos(leverAngle) - constant.caseCLOffset - constant.casing.ShapeWidth1 / 2 - constant.frameOffset + constant.plate3Offset, -constant.plate3.length1 / 2, constant.offset5 - constant.plate2.thickness1 - constant.frame.length1 - constant.plate3.thickness1));
                        AddPlate(constant.plate3, matrix, m_PhysicalAspect.Outputs, "Plate3A");

                        matrix = new Matrix4X4();
                        matrix.Translate(new Vector(CC / 2 + constant.length2 * Math.Cos(leverAngle) + constant.caseCLOffset + constant.casing.ShapeWidth1 / 2 + constant.frameOffset - constant.plate3Offset - constant.plate3.width1, -constant.plate3.length1 / 2, constant.offset5 - constant.plate2.thickness1 - constant.frame.length1 - constant.plate3.thickness1));
                        AddPlate(constant.plate3, matrix, m_PhysicalAspect.Outputs, "Plate3B");

                        //Plate 4 on the edges of the steel section
                        matrix = new Matrix4X4();
                        matrix.Rotate(Math.PI / 2, new Vector(0, 1, 0));
                        matrix.Translate(new Vector(-CC / 2 - constant.length2 * Math.Cos(leverAngle) - constant.caseCLOffset - constant.casing.ShapeWidth1 / 2 - constant.frameOffset + constant.frame.width1, -constant.plate4.length1 / 2, -depth / 2 + constant.plate4.width1 / 2));
                        AddPlate(constant.plate4, matrix, m_PhysicalAspect.Outputs, "Plate4A");

                        matrix = new Matrix4X4();
                        matrix.Rotate(Math.PI / 2, new Vector(0, 1, 0));
                        matrix.Translate(new Vector(CC / 2 + constant.length2 * Math.Cos(leverAngle) + constant.caseCLOffset + constant.casing.ShapeWidth1 / 2 + constant.frameOffset - constant.frame.width1 - constant.plate4.thickness1, -constant.plate4.length1 / 2, -depth / 2 + constant.plate4.width1 / 2));
                        AddPlate(constant.plate4, matrix, m_PhysicalAspect.Outputs, "Plate4B");

                        //Casing
                        constant.casing.ShapeLength = constant.length - constant.plate1.thickness1 - constant.plate2.thickness1 - constant.plate3.thickness1 - constant.frame.length1;
                        matrix = new Matrix4X4();
                        matrix.Translate(new Vector(-CC / 2 - constant.length2 * Math.Cos(leverAngle) - constant.caseCLOffset, 0, constant.offset5));
                        AddNut(constant.casing, matrix, m_PhysicalAspect.Outputs, "Casing1");

                        matrix = new Matrix4X4();
                        matrix.Translate(new Vector(CC / 2 + constant.length2 * Math.Cos(leverAngle) + constant.caseCLOffset, 0, constant.offset5));
                        AddNut(constant.casing, matrix, m_PhysicalAspect.Outputs, "Casing2");

                        //Lever
                        matrix = new Matrix4X4();
                        matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0));
                        matrix.Translate(new Vector(-constant.lever.width1 / 2, constant.leverSpacing / 2 + constant.lever.thickness1, -constant.offset4));
                        matrix.Rotate((Math.PI / 2 - leverAngle), new Vector(0, 1, 0), new Position(0, 0, 0));
                        matrix.Translate(new Vector(-CC / 2 - constant.length2 * Math.Cos(leverAngle), 0, constant.offset5 - constant.pivotOffset));
                        AddPlate(constant.lever, matrix, m_PhysicalAspect.Outputs, "Lever1A");

                        matrix = new Matrix4X4();
                        matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0));
                        matrix.Translate(new Vector(-constant.lever.width1 / 2, -constant.leverSpacing / 2, -constant.offset4));
                        matrix.Rotate((Math.PI / 2 - leverAngle), new Vector(0, 1, 0), new Position(0, 0, 0));
                        matrix.Translate(new Vector(-CC / 2 - constant.length2 * Math.Cos(leverAngle), 0, constant.offset5 - constant.pivotOffset));
                        AddPlate(constant.lever, matrix, m_PhysicalAspect.Outputs, "Lever1B");

                        matrix = new Matrix4X4();
                        matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0));
                        matrix.Translate(new Vector(-constant.lever.width1 / 2, constant.leverSpacing / 2 + constant.lever.thickness1, -constant.offset4));
                        matrix.Rotate((Math.PI / 2 - leverAngle), new Vector(0, 1, 0), new Position(0, 0, 0));
                        matrix.Rotate(Math.PI, new Vector(0, 0, 1), new Position(0, 0, 0));
                        matrix.Translate(new Vector(CC / 2 + constant.length2 * Math.Cos(leverAngle), 0, constant.offset5 - constant.pivotOffset));
                        AddPlate(constant.lever, matrix, m_PhysicalAspect.Outputs, "Lever2A");

                        matrix = new Matrix4X4();
                        matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0));
                        matrix.Translate(new Vector(-constant.lever.width1 / 2, -constant.leverSpacing / 2, -constant.offset4));
                        matrix.Rotate((Math.PI / 2 - leverAngle), new Vector(0, 1, 0), new Position(0, 0, 0));
                        matrix.Rotate(Math.PI, new Vector(0, 0, 1), new Position(0, 0, 0));
                        matrix.Translate(new Vector(CC / 2 + constant.length2 * Math.Cos(leverAngle), 0, constant.offset5 - constant.pivotOffset));
                        AddPlate(constant.lever, matrix, m_PhysicalAspect.Outputs, "Lever2B");

                        //Column
                        if (withTurnbuckle)
                            constant.column.length = constant.length1 - turnbuckleLength + constant.offset2 - constant.clevis.Opening1;
                        else
                            constant.column.length = constant.length1 - constant.columnEnd.ShapeLength - constant.clevis.Opening1;

                        if (HgrCompareDoubleService.cmpdbl(constant.clevis.Diameter2 , 0)==false || HgrCompareDoubleService.cmpdbl(constant.clevis.Width3 , 0)==false)
                        {
                            matrix = new Matrix4X4();
                            matrix.Translate(new Vector(CC / 2, 0, rodTakeout + constant.offset2 - constant.length1));
                            AddClevis(constant.clevis, matrix, m_PhysicalAspect.Outputs, "Clevis1");

                            matrix = new Matrix4X4();
                            matrix.Translate(new Vector(-CC / 2, 0, rodTakeout + constant.offset2 - constant.length1));
                            AddClevis(constant.clevis, matrix, m_PhysicalAspect.Outputs, "Clevis2");
                        }

                        matrix = new Matrix4X4();
                        matrix.Rotate(Math.PI, new Vector(1, 0, 0));
                        matrix.Translate(new Vector(CC / 2, 0, constant.offset5 - constant.pivotOffset + constant.length2 * Math.Sin(leverAngle) + constant.clevis.Opening1 + constant.column.length));
                        AddRod(constant.column, matrix, m_PhysicalAspect.Outputs, "Column1");

                        matrix = new Matrix4X4();
                        matrix.Rotate(Math.PI, new Vector(1, 0, 0));
                        matrix.Translate(new Vector(-CC / 2, 0, constant.offset5 - constant.pivotOffset + constant.length2 * Math.Sin(leverAngle) + constant.clevis.Opening1 + constant.column.length));
                        AddRod(constant.column, matrix, m_PhysicalAspect.Outputs, "Column2");

                        //Column End & Turnbuckle
                        if (withTurnbuckle)
                        {
                            matrix = new Matrix4X4();
                            matrix.Translate(new Vector(-CC / 2, 0, rodTakeout + constant.offset2 - turnbuckleLength / 2));
                            AddTurnbuckle(constant.turnbuckle, 0, matrix, m_PhysicalAspect.Outputs, "Turnbuckle1");

                            matrix = new Matrix4X4();
                            matrix.Translate(new Vector(CC / 2, 0, rodTakeout + constant.offset2 - turnbuckleLength / 2));
                            AddTurnbuckle(constant.turnbuckle, 0, matrix, m_PhysicalAspect.Outputs, "Turnbuckle2");
                        }
                        else
                        {
                            matrix = new Matrix4X4();
                            matrix.Translate(new Vector(-CC / 2, 0, rodTakeout + constant.offset2 - constant.columnEnd.ShapeLength));
                            AddNut(constant.columnEnd, matrix, m_PhysicalAspect.Outputs, "ColumnEnd1");

                            matrix = new Matrix4X4();
                            matrix.Translate(new Vector(CC / 2, 0, rodTakeout + constant.offset2 - constant.columnEnd.ShapeLength));
                            AddNut(constant.columnEnd, matrix, m_PhysicalAspect.Outputs, "ColumnEnd2");
                        }

                        //Pins
                        if (HgrCompareDoubleService.cmpdbl(constant.pivotPin.PinDiameter , 0)==false)
                        {
                            matrix = new Matrix4X4();
                            matrix.Translate(new Vector(-CC / 2 - constant.length2 * Math.Cos(leverAngle), 0, constant.offset5 - constant.pivotOffset));
                            AddPin(constant.pivotPin, matrix, m_PhysicalAspect.Outputs, "PivotPin1");

                            matrix = new Matrix4X4();
                            matrix.Translate(new Vector(CC / 2 + constant.length2 * Math.Cos(leverAngle), 0, constant.offset5 - constant.pivotOffset));
                            AddPin(constant.pivotPin, matrix, m_PhysicalAspect.Outputs, "PivotPin2");
                        }
                        if (HgrCompareDoubleService.cmpdbl(constant.columnPin.PinDiameter , 0)==false)
                        {
                            matrix = new Matrix4X4();
                            matrix.Translate(new Vector(-CC / 2, 0, constant.offset5 - constant.pivotOffset + constant.length2 * Math.Sin(leverAngle)));
                            AddPin(constant.columnPin, matrix, m_PhysicalAspect.Outputs, "ColumnPin1");

                            matrix = new Matrix4X4();
                            matrix.Translate(new Vector(CC / 2, 0, constant.offset5 - constant.pivotOffset + constant.length2 * Math.Sin(leverAngle)));
                            AddPin(constant.columnPin, matrix, m_PhysicalAspect.Outputs, "ColumnPin2");
                        }
                        break;
                    case 2:
                        //Frame
                        matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0));
                        matrix.Rotate(Math.PI, new Vector(0, 1, 0));
                        matrix.Translate(new Vector(CC / 2 + constant.length2 * Math.Cos(leverAngle) - constant.caseCLOffset + constant.casing.ShapeWidth1 / 2 + constant.frameOffset, constant.frameSpacing / 2 + constant.frame.thickness1, constant.offset5 + constant.frame.length1));
                        AddPlate(constant.frame, matrix, m_PhysicalAspect.Outputs, "Frame1A");

                        matrix = new Matrix4X4();
                        matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0));
                        matrix.Rotate(Math.PI, new Vector(0, 1, 0));
                        matrix.Translate(new Vector(CC / 2 + constant.length2 * Math.Cos(leverAngle) - constant.caseCLOffset + constant.casing.ShapeWidth1 / 2 + constant.frameOffset, -constant.frameSpacing / 2, constant.offset5 + constant.frame.length1));
                        AddPlate(constant.frame, matrix, m_PhysicalAspect.Outputs, "Frame1B");

                        matrix = new Matrix4X4();
                        matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0));
                        matrix.Rotate(Math.PI, new Vector(0, 1, 0));
                        matrix.Rotate(Math.PI, new Vector(0, 0, 1));
                        matrix.Translate(new Vector(-CC / 2 - constant.length2 * Math.Cos(leverAngle) + constant.caseCLOffset - constant.casing.ShapeWidth1 / 2 - constant.frameOffset, constant.frameSpacing / 2, constant.offset5 + constant.frame.length1));
                        AddPlate(constant.frame, matrix, m_PhysicalAspect.Outputs, "Frame2A");

                        matrix = new Matrix4X4();
                        matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0));
                        matrix.Rotate(Math.PI, new Vector(0, 1, 0));
                        matrix.Rotate(Math.PI, new Vector(0, 0, 1));
                        matrix.Translate(new Vector(-CC / 2 - constant.length2 * Math.Cos(leverAngle) + constant.caseCLOffset - constant.casing.ShapeWidth1 / 2 - constant.frameOffset, -constant.frameSpacing / 2 - constant.frame.thickness1, constant.offset5 + constant.frame.length1));
                        AddPlate(constant.frame, matrix, m_PhysicalAspect.Outputs, "Frame2B");

                        //Plates
                        matrix = new Matrix4X4();
                        matrix.Translate(new Vector(-CC / 2 - constant.length2 * Math.Cos(leverAngle) + constant.caseCLOffset - constant.casing.ShapeWidth1 / 2 + constant.plate1Offset, -constant.plate1.length1 / 2, constant.offset5 + constant.frame.length1 + constant.plate3.thickness1 - constant.length));
                        AddPlate(constant.plate1, matrix, m_PhysicalAspect.Outputs, "Plate1A");

                        matrix = new Matrix4X4();
                        matrix.Translate(new Vector(CC / 2 + constant.length2 * Math.Cos(leverAngle) - constant.caseCLOffset + constant.casing.ShapeWidth1 / 2 - constant.plate1Offset - constant.plate1.width1, -constant.plate1.length1 / 2, constant.offset5 + constant.frame.length1 + constant.plate3.thickness1 - constant.length));
                        AddPlate(constant.plate1, matrix, m_PhysicalAspect.Outputs, "Plate1B");

                        matrix = new Matrix4X4();
                        matrix.Translate(new Vector(-CC / 2 - constant.length2 * Math.Cos(leverAngle) + constant.caseCLOffset - constant.casing.ShapeWidth1 / 2 + constant.plate2Offset, -constant.plate2.length1 / 2, constant.offset5 - constant.plate2.thickness1));
                        AddPlate(constant.plate2, matrix, m_PhysicalAspect.Outputs, "Plate2A");

                        matrix = new Matrix4X4();
                        matrix.Translate(new Vector(CC / 2 + constant.length2 * Math.Cos(leverAngle) - constant.caseCLOffset + constant.casing.ShapeWidth1 / 2 - constant.plate2Offset - constant.plate2.width1, -constant.plate2.length1 / 2, constant.offset5 - constant.plate2.thickness1));
                        AddPlate(constant.plate2, matrix, m_PhysicalAspect.Outputs, "Plate2B");

                        matrix = new Matrix4X4();
                        matrix.Translate(new Vector(-CC / 2 - constant.length2 * Math.Cos(leverAngle) + constant.caseCLOffset - constant.casing.ShapeWidth1 / 2 - constant.frameOffset + constant.plate3Offset, -constant.plate3.length1 / 2, constant.offset5 + constant.frame.length1));
                        AddPlate(constant.plate3, matrix, m_PhysicalAspect.Outputs, "Plate3A");

                        matrix = new Matrix4X4();
                        matrix.Translate(new Vector(CC / 2 + constant.length2 * Math.Cos(leverAngle) - constant.caseCLOffset + constant.casing.ShapeWidth1 / 2 + constant.frameOffset - constant.plate3Offset - constant.plate3.width1, -constant.plate3.length1 / 2, constant.offset5 + constant.frame.length1));
                        AddPlate(constant.plate3, matrix, m_PhysicalAspect.Outputs, "Plate3B");

                        //Casing
                        constant.casing.ShapeLength = constant.length - constant.plate1.thickness1 - constant.plate2.thickness1 - constant.plate3.thickness1 - constant.frame.length1;
                        matrix = new Matrix4X4();
                        matrix.Translate(new Vector(-CC / 2 - constant.length2 * Math.Cos(leverAngle) + constant.caseCLOffset, 0, constant.offset5 + constant.frame.length1 + constant.plate3.thickness1 - constant.length + constant.plate1.thickness1));
                        AddNut(constant.casing, matrix, m_PhysicalAspect.Outputs, "Casing1");

                        matrix = new Matrix4X4();
                        matrix.Translate(new Vector(CC / 2 + constant.length2 * Math.Cos(leverAngle) - constant.caseCLOffset, 0, constant.offset5 + constant.frame.length1 + constant.plate3.thickness1 - constant.length + constant.plate1.thickness1));
                        AddNut(constant.casing, matrix, m_PhysicalAspect.Outputs, "Casing2");

                        //Lever
                        matrix = new Matrix4X4();
                        matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0));
                        matrix.Translate(new Vector(-constant.lever.width1 / 2, constant.leverSpacing / 2 + constant.lever.thickness1, -constant.offset4));
                        matrix.Rotate((Math.PI / 2 - leverAngle), new Vector(0, 1, 0), new Position(0, 0, 0));
                        matrix.Translate(new Vector(-CC / 2 - constant.length2 * Math.Cos(leverAngle), 0, constant.offset5 + constant.frame.length1 + constant.plate3.thickness1 - constant.pivotOffset));
                        AddPlate(constant.lever, matrix, m_PhysicalAspect.Outputs, "Lever1A");

                        matrix = new Matrix4X4();
                        matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0));
                        matrix.Translate(new Vector(-constant.lever.width1 / 2, -constant.leverSpacing / 2, -constant.offset4));
                        matrix.Rotate((Math.PI / 2 - leverAngle), new Vector(0, 1, 0), new Position(0, 0, 0));
                        matrix.Translate(new Vector(-CC / 2 - constant.length2 * Math.Cos(leverAngle), 0, constant.offset5 + constant.frame.length1 + constant.plate3.thickness1 - constant.pivotOffset));
                        AddPlate(constant.lever, matrix, m_PhysicalAspect.Outputs, "Lever1B");

                        matrix = new Matrix4X4();
                        matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0));
                        matrix.Translate(new Vector(-constant.lever.width1 / 2, constant.leverSpacing / 2 + constant.lever.thickness1, -constant.offset4));
                        matrix.Rotate((Math.PI / 2 - leverAngle), new Vector(0, 1, 0), new Position(0, 0, 0));
                        matrix.Rotate(Math.PI, new Vector(0, 0, 1), new Position(0, 0, 0));
                        matrix.Translate(new Vector(CC / 2 + constant.length2 * Math.Cos(leverAngle), 0, constant.offset5 + constant.frame.length1 + constant.plate3.thickness1 - constant.pivotOffset));
                        AddPlate(constant.lever, matrix, m_PhysicalAspect.Outputs, "Lever2A");

                        matrix = new Matrix4X4();
                        matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0));
                        matrix.Translate(new Vector(-constant.lever.width1 / 2, -constant.leverSpacing / 2, -constant.offset4));
                        matrix.Rotate((Math.PI / 2 - leverAngle), new Vector(0, 1, 0), new Position(0, 0, 0));
                        matrix.Rotate(Math.PI, new Vector(0, 0, 1), new Position(0, 0, 0));
                        matrix.Translate(new Vector(CC / 2 + constant.length2 * Math.Cos(leverAngle), 0, constant.offset5 + constant.frame.length1 + constant.plate3.thickness1 - constant.pivotOffset));
                        AddPlate(constant.lever, matrix, m_PhysicalAspect.Outputs, "Lever2B");

                        //Column
                        if (withTurnbuckle)
                            constant.column.length = constant.length1 - turnbuckleLength + constant.offset2 - constant.clevis.Opening1;
                        else
                            constant.column.length = constant.length1 - constant.columnEnd.ShapeLength - constant.clevis.Opening1;

                        if (HgrCompareDoubleService.cmpdbl(constant.clevis.Diameter2 , 0)==false || HgrCompareDoubleService.cmpdbl(constant.clevis.Width3 , 0)==false)
                        {
                            matrix = new Matrix4X4();
                            matrix.Translate(new Vector(CC / 2, 0, constant.offset5 + constant.frame.length1 + constant.plate3.thickness1 - constant.pivotOffset + constant.length2 * Math.Sin(leverAngle)));
                            AddClevis(constant.clevis, matrix, m_PhysicalAspect.Outputs, "Clevis1");

                            matrix = new Matrix4X4();
                            matrix.Translate(new Vector(-CC / 2, 0, constant.offset5 + constant.frame.length1 + constant.plate3.thickness1 - constant.pivotOffset + constant.length2 * Math.Sin(leverAngle)));
                            AddClevis(constant.clevis, matrix, m_PhysicalAspect.Outputs, "Clevis2");
                        }

                        matrix = new Matrix4X4();
                        matrix.Rotate(Math.PI, new Vector(1, 0, 0));
                        matrix.Translate(new Vector(CC / 2, 0, constant.offset5 + constant.frame.length1 + constant.plate3.thickness1 - constant.pivotOffset + constant.length2 * Math.Sin(leverAngle) + constant.column.length + constant.clevis.Opening1));
                        AddRod(constant.column, matrix, m_PhysicalAspect.Outputs, "Column1");

                        matrix = new Matrix4X4();
                        matrix.Rotate(Math.PI, new Vector(1, 0, 0));
                        matrix.Translate(new Vector(-CC / 2, 0, constant.offset5 + constant.frame.length1 + constant.plate3.thickness1 - constant.pivotOffset + constant.length2 * Math.Sin(leverAngle) + constant.column.length + constant.clevis.Opening1));
                        AddRod(constant.column, matrix, m_PhysicalAspect.Outputs, "Column2");

                        //Column End & Turnbuckle
                        if (withTurnbuckle)
                        {
                            matrix = new Matrix4X4();
                            matrix.Translate(new Vector(-CC / 2, 0, constant.offset5 + constant.frame.length1 + constant.plate3.thickness1 - constant.pivotOffset + constant.length2 * Math.Sin(leverAngle) + constant.length1 - turnbuckleLength / 2));
                            AddTurnbuckle(constant.turnbuckle, 0, matrix, m_PhysicalAspect.Outputs, "Turnbuckle1");

                            matrix = new Matrix4X4();
                            matrix.Translate(new Vector(CC / 2, 0, constant.offset5 + constant.frame.length1 + constant.plate3.thickness1 - constant.pivotOffset + constant.length2 * Math.Sin(leverAngle) + constant.length1 - turnbuckleLength / 2));
                            AddTurnbuckle(constant.turnbuckle, 0, matrix, m_PhysicalAspect.Outputs, "Turnbuckle2");
                        }
                        else
                        {
                            matrix = new Matrix4X4();
                            matrix.Translate(new Vector(-CC / 2, 0, constant.offset5 + constant.frame.length1 + constant.plate3.thickness1 - constant.pivotOffset + constant.length2 * Math.Sin(leverAngle) + constant.column.length + constant.clevis.Opening1));
                            AddNut(constant.columnEnd, matrix, m_PhysicalAspect.Outputs, "ColumnEnd1");

                            matrix = new Matrix4X4();
                            matrix.Translate(new Vector(CC / 2, 0, constant.offset5 + constant.frame.length1 + constant.plate3.thickness1 - constant.pivotOffset + constant.length2 * Math.Sin(leverAngle) + constant.column.length + constant.clevis.Opening1));
                            AddNut(constant.columnEnd, matrix, m_PhysicalAspect.Outputs, "ColumnEnd2");
                        }

                        //Pins
                        if (HgrCompareDoubleService.cmpdbl(constant.pivotPin.PinDiameter , 0)==false)
                        {
                            matrix = new Matrix4X4();
                            matrix.Translate(new Vector(-CC / 2 - constant.length2 * Math.Cos(leverAngle), 0, constant.offset5 + constant.frame.length1 + constant.plate3.thickness1 - constant.pivotOffset));
                            AddPin(constant.pivotPin, matrix, m_PhysicalAspect.Outputs, "PivotPin1");

                            matrix = new Matrix4X4();
                            matrix.Translate(new Vector(CC / 2 + constant.length2 * Math.Cos(leverAngle), 0, constant.offset5 + constant.frame.length1 + constant.plate3.thickness1 - constant.pivotOffset));
                            AddPin(constant.pivotPin, matrix, m_PhysicalAspect.Outputs, "PivotPin2");
                        }
                        if (HgrCompareDoubleService.cmpdbl(constant.columnPin.PinDiameter , 0)==false)
                        {
                            matrix = new Matrix4X4();
                            matrix.Translate(new Vector(-CC / 2, 0, constant.offset5 + constant.frame.length1 + constant.plate3.thickness1 - constant.pivotOffset + constant.length2 * Math.Sin(leverAngle)));
                            AddPin(constant.columnPin, matrix, m_PhysicalAspect.Outputs, "ColumnPin1");

                            matrix = new Matrix4X4();
                            matrix.Translate(new Vector(CC / 2, 0, constant.offset5 + constant.frame.length1 + constant.plate3.thickness1 - constant.pivotOffset + constant.length2 * Math.Sin(leverAngle)));
                            AddPin(constant.columnPin, matrix, m_PhysicalAspect.Outputs, "ColumnPin2");
                        }
                        break;
                }
                // Add the Steel Cross Section
                SweepOptions sweepOptions = (SweepOptions)7;
                CrossSectionServices crossSectionServices = new CrossSectionServices();

                Double lenAdj = 0;
                if (config == 1)
                    lenAdj = constant.length2 * Math.Cos(leverAngle) + constant.caseCLOffset + constant.casing.ShapeWidth1 / 2 + constant.frameOffset - constant.frame.width1 - constant.plate4.thickness1;
                else
                {
                    switch (constant.casing.ShapeType)
                    {
                        case 1:     //Round
                            if (bBGap >= 0)
                            {
                                if (width + bBGap / 2 > constant.casing.ShapeWidth1 / 2)
                                    lenAdj = constant.length2 * Math.Cos(leverAngle) - constant.caseCLOffset;
                                else
                                    lenAdj = constant.length2 * Math.Cos(leverAngle) - constant.caseCLOffset - Math.Sqrt((constant.casing.ShapeWidth1 / 2) * (constant.casing.ShapeWidth1 / 2) - (width + bBGap / 2) * (width + bBGap / 2));
                            }
                            else
                            {
                                if (width / 2 > constant.casing.ShapeWidth1 / 2)
                                    lenAdj = constant.length2 * Math.Cos(leverAngle) - constant.caseCLOffset;
                                else
                                    lenAdj = constant.length2 * Math.Cos(leverAngle) - constant.caseCLOffset - Math.Sqrt((constant.casing.ShapeWidth1 / 2) * (constant.casing.ShapeWidth1 / 2) - (width / 2) * (width / 2));
                            }
                            break;
                        case 2:     //Square
                            lenAdj = constant.length2 * Math.Cos(leverAngle) - constant.caseCLOffset - constant.casing.ShapeWidth1 / 2;
                            break;
                        case 3:     //Hex
                            if (bBGap >= 0)
                                lenAdj = constant.length2 * Math.Cos(leverAngle) - constant.caseCLOffset - ((constant.casing.ShapeWidth1 / 2) - (width + bBGap / 2) * Math.Tan(30 * Math.PI / 180));
                            else
                                lenAdj = constant.length2 * Math.Cos(leverAngle) - constant.caseCLOffset - ((constant.casing.ShapeWidth1 / 2) - (width / 2) * Math.Tan(30 * Math.PI / 180));
                            break;
                    }
                }

                Line3d projection;

                // Calculate the Cut Angles
                Double[] startPosNorm1 = new Double[6]; Double[] endPosNorm1 = new Double[6]; Double[] startPosNorm2 = new Double[6]; Double[] endPosNorm2 = new Double[6];
                if (config == 1)
                {
                    startPosNorm1[0] = -CC / 2 - lenAdj;
                    startPosNorm1[1] = 0;
                    startPosNorm1[2] = 0;
                    startPosNorm1[3] = 1;
                    startPosNorm1[4] = 0;
                    startPosNorm1[5] = 0;

                    endPosNorm1[0] = CC / 2 + lenAdj;
                    endPosNorm1[1] = 0;
                    endPosNorm1[2] = 0;
                    endPosNorm1[3] = 1;
                    endPosNorm1[4] = 0;
                    endPosNorm1[5] = 0;

                    startPosNorm2[0] = -CC / 2 - lenAdj;
                    startPosNorm2[1] = 0;
                    startPosNorm2[2] = 0;
                    startPosNorm2[3] = 1;
                    startPosNorm2[4] = 0;
                    startPosNorm2[5] = 0;

                    endPosNorm2[0] = CC / 2 + lenAdj;
                    endPosNorm2[1] = 0;
                    endPosNorm2[2] = 0;
                    endPosNorm2[3] = 1;
                    endPosNorm2[4] = 0;
                    endPosNorm2[5] = 0;
                }
                else
                {
                    if (bBGap >= 0)
                    {
                        switch (constant.casing.ShapeType)
                        {
                            case 1:
                                Double alpha;
                                if (width + bBGap / 2 > constant.casing.ShapeWidth1 / 2)
                                    alpha = 0;
                                else
                                    alpha = (Math.Atan((Math.Sqrt((constant.casing.ShapeWidth1 / 2) * (constant.casing.ShapeWidth1 / 2) - (bBGap / 2) * (bBGap / 2)) - Math.Sqrt((constant.casing.ShapeWidth1 / 2) * (constant.casing.ShapeWidth1 / 2) - (width + bBGap / 2) * (width + bBGap / 2))) / width)) * 180 / Math.PI;
                                startPosNorm1[0] = -CC / 2 - lenAdj;
                                startPosNorm1[1] = -width - bBGap / 2;
                                startPosNorm1[2] = 0;
                                startPosNorm1[3] = -Math.Cos(alpha * Math.PI / 180);
                                startPosNorm1[4] = Math.Sin(alpha * Math.PI / 180);
                                startPosNorm1[5] = 0;

                                endPosNorm1[0] = CC / 2 + lenAdj;
                                endPosNorm1[1] = -width - bBGap / 2;
                                endPosNorm1[2] = 0;
                                endPosNorm1[3] = Math.Cos(alpha * Math.PI / 180);
                                endPosNorm1[4] = Math.Sin(alpha * Math.PI / 180);
                                endPosNorm1[5] = 0;

                                startPosNorm2[0] = -CC / 2 - lenAdj;
                                startPosNorm2[1] = width + bBGap / 2;
                                startPosNorm2[2] = 0;
                                startPosNorm2[3] = -Math.Cos(alpha * Math.PI / 180);
                                startPosNorm2[4] = -Math.Sin(alpha * Math.PI / 180);
                                startPosNorm2[5] = 0;

                                endPosNorm2[0] = CC / 2 + lenAdj;
                                endPosNorm2[1] = width + bBGap / 2;
                                endPosNorm2[2] = 0;
                                endPosNorm2[3] = Math.Cos(alpha * Math.PI / 180);
                                endPosNorm2[4] = -Math.Sin(alpha * Math.PI / 180);
                                endPosNorm2[5] = 0;
                                break;
                            case 2:     //Square
                                startPosNorm1[0] = -CC / 2 - lenAdj;
                                startPosNorm1[1] = 0;
                                startPosNorm1[2] = 0;
                                startPosNorm1[3] = 1;
                                startPosNorm1[4] = 0;
                                startPosNorm1[5] = 0;

                                endPosNorm1[0] = CC / 2 + lenAdj;
                                endPosNorm1[1] = 0;
                                endPosNorm1[2] = 0;
                                endPosNorm1[3] = 1;
                                endPosNorm1[4] = 0;
                                endPosNorm1[5] = 0;

                                startPosNorm2[0] = -CC / 2 - lenAdj;
                                startPosNorm2[1] = 0;
                                startPosNorm2[2] = 0;
                                startPosNorm2[3] = 1;
                                startPosNorm2[4] = 0;
                                startPosNorm2[5] = 0;

                                endPosNorm2[0] = CC / 2 + lenAdj;
                                endPosNorm2[1] = 0;
                                endPosNorm2[2] = 0;
                                endPosNorm2[3] = 1;
                                endPosNorm2[4] = 0;
                                endPosNorm2[5] = 0;
                                break;
                            case 3:     //Hex
                                startPosNorm1[0] = -CC / 2 - lenAdj;
                                startPosNorm1[1] = -width - bBGap / 2;
                                startPosNorm1[2] = 0;
                                startPosNorm1[3] = -Math.Cos(30 * Math.PI / 180);
                                startPosNorm1[4] = Math.Sin(30 * Math.PI / 180);
                                startPosNorm1[5] = 0;

                                endPosNorm1[0] = CC / 2 + lenAdj;
                                endPosNorm1[1] = -width - bBGap / 2;
                                endPosNorm1[2] = 0;
                                endPosNorm1[3] = Math.Cos(30 * Math.PI / 180);
                                endPosNorm1[4] = Math.Sin(30 * Math.PI / 180);
                                endPosNorm1[5] = 0;

                                startPosNorm2[0] = -CC / 2 - lenAdj;
                                startPosNorm2[1] = width + bBGap / 2;
                                startPosNorm2[2] = 0;
                                startPosNorm2[3] = -Math.Cos(30 * Math.PI / 180);
                                startPosNorm2[4] = -Math.Sin(30 * Math.PI / 180);
                                startPosNorm2[5] = 0;

                                endPosNorm2[0] = CC / 2 + lenAdj;
                                endPosNorm2[1] = width + bBGap / 2;
                                endPosNorm2[2] = 0;
                                endPosNorm2[3] = Math.Cos(30 * Math.PI / 180);
                                endPosNorm2[4] = -Math.Sin(30 * Math.PI / 180);
                                endPosNorm2[5] = 0;
                                break;
                        }
                    }
                    else
                    {
                        startPosNorm1[0] = -CC / 2 - lenAdj;
                        startPosNorm1[1] = 0;
                        startPosNorm1[2] = 0;
                        startPosNorm1[3] = 1;
                        startPosNorm1[4] = 0;
                        startPosNorm1[5] = 0;

                        endPosNorm1[0] = CC / 2 + lenAdj;
                        endPosNorm1[1] = 0;
                        endPosNorm1[2] = 0;
                        endPosNorm1[3] = 1;
                        endPosNorm1[4] = 0;
                        endPosNorm1[5] = 0;

                        startPosNorm2[0] = -CC / 2 - lenAdj;
                        startPosNorm2[1] = 0;
                        startPosNorm2[2] = 0;
                        startPosNorm2[3] = 1;
                        startPosNorm2[4] = 0;
                        startPosNorm2[5] = 0;

                        endPosNorm2[0] = CC / 2 + lenAdj;
                        endPosNorm2[1] = 0;
                        endPosNorm2[2] = 0;
                        endPosNorm2[3] = 1;
                        endPosNorm2[4] = 0;
                        endPosNorm2[5] = 0;
                    }
                }
                string section1 = "Section1"; string section2 = "Section2";
                Collection<ISurface> crossSecSurfaces1, crossSecSurfaces2;
                if (CC / 2 > 0)
                {
                    if (bBGap >= 0)
                    {
                        projection = new Line3d(new Position(-CC / 2 - lenAdj, -bBGap / 2, 0), new Position(CC / 2 + lenAdj, -bBGap / 2, 0));
                        //First Section
                        crossSecSurfaces1 = crossSectionServices.GetProjectionSurfacesFromCrossSection(crossSection1, projection, 7, false, 0, new Position(startPosNorm1[0], startPosNorm1[1], startPosNorm1[2]), new Vector(startPosNorm1[3], startPosNorm1[4], startPosNorm1[5]), new Position(endPosNorm1[0], endPosNorm1[1], endPosNorm1[2]), new Vector(endPosNorm1[3], endPosNorm1[4], endPosNorm1[5]), sweepOptions);
                        for (int i = 1; i <= crossSecSurfaces1.Count; i++)
                        {
                            m_PhysicalAspect.Outputs[section1] = crossSecSurfaces1[i - 1];
                            section1 = "Section1" + i;
                        }
                        //Second Section
                        crossSection2 = crossSection1;
                        projection = new Line3d(new Position(-CC / 2 - lenAdj, bBGap / 2, 0), new Position(CC / 2 + lenAdj, bBGap / 2, 0));
                        crossSecSurfaces2 = crossSectionServices.GetProjectionSurfacesFromCrossSection(crossSection2, projection, 7, true, 0, new Position(startPosNorm2[0], startPosNorm2[1], startPosNorm2[2]), new Vector(startPosNorm2[3], startPosNorm2[4], startPosNorm2[5]), new Position(endPosNorm2[0], endPosNorm2[1], endPosNorm2[2]), new Vector(endPosNorm2[3], endPosNorm2[4], endPosNorm2[5]), sweepOptions);
                        for (int i = 1; i <= crossSecSurfaces2.Count; i++)
                        {
                            m_PhysicalAspect.Outputs[section2] = crossSecSurfaces2[i - 1];
                            section2 = "Section2" + i;
                        }
                    }
                    else
                    {
                        projection = new Line3d(new Position(-CC / 2 - lenAdj, 0, 0), new Position(CC / 2 + lenAdj, 0, 0));
                        //First Section
                        crossSecSurfaces1 = crossSectionServices.GetProjectionSurfacesFromCrossSection(crossSection1, projection, 8, false, 0, new Position(startPosNorm1[0], startPosNorm1[1], startPosNorm1[2]), new Vector(startPosNorm1[3], startPosNorm1[4], startPosNorm1[5]), new Position(endPosNorm1[0], endPosNorm1[1], endPosNorm1[2]), new Vector(endPosNorm1[3], endPosNorm1[4], endPosNorm1[5]), sweepOptions);
                        for (int i = 1; i <= crossSecSurfaces1.Count; i++)
                        {
                            m_PhysicalAspect.Outputs[section1] = crossSecSurfaces1[i - 1];
                            section1 = "Section1" + i;
                        }
                    }
                }
                else
                {
                    RaiseSmartPartTDLError(SmartPartSymbolResourceIDs.ErrInvalidCC, "CC Value should be greater than Zero.");
                    return;
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
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of TypeG."));
                    return;
                }
            }
        }
        #endregion

        #region ICustomWeightCG Members

        void ICustomWeightCG.EvaluateWeightCG(BusinessObject supportComponentBO)
        {
            try
            {
                ////System WCG Attributes

                Part CatalogPart = (Part)supportComponentBO.GetRelationship("madeFrom", "part").TargetObjects[0];

                Double weight, cogX, cogY, cogZ;
                try
                {
                    weight = (double)((PropertyValueDouble)CatalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryWeight")).PropValue;
                }
                catch
                {
                    weight = 0;
                }
                //Center of Gravity
                try
                {
                    cogX = (double)((PropertyValueDouble)CatalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryCogX")).PropValue;
                }
                catch
                {
                    cogX = 0;
                }
                try
                {
                    cogY = (double)((PropertyValueDouble)CatalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryCogY")).PropValue;
                }
                catch
                {
                    cogY = 0;
                }
                try
                {
                    cogZ = (double)((PropertyValueDouble)CatalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryCogZ")).PropValue;
                }
                catch
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
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrWeightCG, "Error in Weight CG of TypeG."));
                    return;
                }
            }
        }
        #endregion
    }
}

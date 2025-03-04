//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   ConstantTypeU.cs
//    HSSmartPart,Ingr.SP3D.Content.Support.Symbols.ConstantTypeU
//   Author       :  Vijay
//   Creation Date: 05.June.2013 
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   05.June.2013    Vijay   CR-CP-222480  Convert HS_S3DConstant Smartpart VB Project to C# .Net  
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
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
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
    public class ConstantTypeU : SmartPartComponentDefinition, ICustomWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "HSSmartPart,Ingr.SP3D.Content.Support.Symbols.ConstantTypeU"
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

                //Add Configuration Inputs
                additionalInputs.Add(new InputDouble(++endIndex, "ConstantConfig", "ConstantConfig", 1, false));

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

                //Add Travel Increment Rule Input
                additionalInputs.Add(new InputString(++endIndex, "TTIncrementByRule", "TTIncrementByRule", "No Value", true));

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
        [SymbolOutput("Structure", "Structure")]
        [SymbolOutput("Surface", "Surface")]
        public AspectDefinition m_PhysicalAspect;

        #endregion

        #region "Definition of Additional Outputs"

        public override IEnumerable<OutputDefinition> AdditionalOutputs(string aspectName)
        {
            List<OutputDefinition> additionalOutputs = new List<OutputDefinition>();
            if (aspectName == "Symbolic")
            {
                //Add Outputs for the Plates
                AddPlateOutputs(4, additionalOutputs);

                //Add Outputs for the Casing and Frame
                additionalOutputs.Add(new OutputDefinition("Casing", "Casing"));
                additionalOutputs.Add(new OutputDefinition("FrameA", "FrameA"));
                additionalOutputs.Add(new OutputDefinition("FrameB", "FrameB"));
                additionalOutputs.Add(new OutputDefinition("Lever1A", "Lever1A"));
                additionalOutputs.Add(new OutputDefinition("Lever1B", "Lever1B"));
                additionalOutputs.Add(new OutputDefinition("Lever2A", "Lever2A"));
                additionalOutputs.Add(new OutputDefinition("Lever2B", "Lever2B"));

                //Add Outputs for the Load Column, and the Column End Shape
                AddRod1Outputs("Column", additionalOutputs);
                additionalOutputs.Add(new OutputDefinition("ColumnPlate", "ColumnPlate"));
                additionalOutputs.Add(new OutputDefinition("ColumnEnd", "ColumnEnd"));
                additionalOutputs.Add(new OutputDefinition("ColumnNut", "ColumnNut"));

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
                int config = (int)GetDoubleInputValue(++startIndex);
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
                if (totalTravelUnits == "in")
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
                if (totalTravelUnits == "in")
                    totalTravel = MiddleServiceProvider.UOMMgr.ConvertUnitToDBU(UnitType.Distance, totalTravelInUnits, UnitName.DISTANCE_INCH);
                else
                    totalTravel = MiddleServiceProvider.UOMMgr.ConvertUnitToDBU(UnitType.Distance, totalTravelInUnits, UnitName.DISTANCE_MILLIMETER);

                //Load the Constant Reference Data
                ConstantInputs constant = LoadConstantData(sizeTable, travelTable, loadTable, size, Math.Round(totalTravelInUnits, 1), load);

                //Determine the Angle of the Lever Arm and the Rod Take Out
                Double leverAngle, rodTakeout;
                if(movementDirection == 1)
                {
                    leverAngle = -constant.angleLow;
                    rodTakeout = constant.rodTakeOutLow;
                }
                else
                {
                    leverAngle = constant.angleHigh;
                    rodTakeout = constant.rodTakeOutHigh;
                }

                //Set the Output Part Occurences
                try
                {
                    SupportComponent supportComponent = (SupportComponent)Occurrence;
                    supportComponent.SetPropertyValue(rodTakeout, "IJUAhsRodTakeOut", "RodTakeOut");
                    supportComponent.SetPropertyValue(totalTravel, "IJUAhsTotalTravel", "TotalTravel");
                }
                catch{ }

                //Add the Ports
                Port port;
                switch (config)
                {
                    case 1:     //Vertical Casing Above Frame
                        RaiseSmartPartTDLError(SmartPartSymbolResourceIDs.ErrUConfig, "Configuration 1 not valid for Type U Constants");
                        return;
                    case 2:     //Vertical Casing Below Frame
                        port = new Port(OccurrenceConnection, part, "Route", new Position(-constant.frameOffset, 0, rodTakeout + pipeOuterDiameter / 2 + shoeHeight), new Vector(1, 0, 0), new Vector(0, 0, 1));
                        m_PhysicalAspect.Outputs["Route"] = port;

                        port = new Port(OccurrenceConnection, part, "Structure", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                        m_PhysicalAspect.Outputs["Structure"] = port;

                        port = new Port(OccurrenceConnection, part, "Surface", new Position(-constant.frameOffset, 0, rodTakeout), new Vector(1, 0, 0), new Vector(0, 0, 1));
                        m_PhysicalAspect.Outputs["Surface"] = port;
                        break;
                    case 3:     //Horizontal Casing Beside Frame
                        port = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, rodTakeout + pipeOuterDiameter / 2 + shoeHeight), new Vector(1, 0, 0), new Vector(0, 0, 1));
                        m_PhysicalAspect.Outputs["Route"] = port;

                        port = new Port(OccurrenceConnection, part, "Structure", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                        m_PhysicalAspect.Outputs["Structure"] = port;

                        port = new Port(OccurrenceConnection, part, "Surface", new Position(0, 0, rodTakeout), new Vector(1, 0, 0), new Vector(0, 0, 1));
                        m_PhysicalAspect.Outputs["Surface"] = port;
                        break;
                }

                switch (config)
                {
                    case 1:     //Vertical Casing Above Frame
                        RaiseSmartPartTDLError(SmartPartSymbolResourceIDs.ErrUConfig, "Configuration 1 not valid for Type U Constants");
                        return;
                    case 2:     //Vertical Casing Below Frame
                        //Frame
                        matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0));
                        matrix.Translate(new Vector(-constant.casing.ShapeWidth1 / 2 - constant.frameOffset, constant.frameSpacing / 2 + constant.frame.thickness1, constant.length - constant.plate3.thickness1 - constant.frame.length1));
                        AddPlate(constant.frame, matrix, m_PhysicalAspect.Outputs, "FrameA");

                        matrix = new Matrix4X4();
                        matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0));
                        matrix.Translate(new Vector(-constant.casing.ShapeWidth1 / 2 - constant.frameOffset, -constant.frameSpacing / 2, constant.length - constant.plate3.thickness1 - constant.frame.length1));
                        AddPlate(constant.frame, matrix, m_PhysicalAspect.Outputs, "FrameB");

                        //Plates
                        matrix = new Matrix4X4();
                        matrix.Translate(new Vector(-constant.casing.ShapeWidth1 / 2 - constant.plate1Offset, -constant.plate1.length1 / 2, 0));
                        AddPlate(constant.plate1, matrix, m_PhysicalAspect.Outputs, "Plate1");

                        matrix = new Matrix4X4();
                        matrix.Translate(new Vector(-constant.casing.ShapeWidth1 / 2 - constant.plate2Offset, -constant.plate2.length1 / 2, constant.length - constant.plate3.thickness1 - constant.frame.length1 - constant.plate2.thickness1));
                        AddPlate(constant.plate2, matrix, m_PhysicalAspect.Outputs, "Plate2");

                        matrix = new Matrix4X4();
                        matrix.Translate(new Vector(-constant.casing.ShapeWidth1 / 2 - constant.frameOffset - constant.plate3Offset, -constant.plate3.length1 / 2, constant.length - constant.plate3.thickness1));
                        AddPlate(constant.plate3, matrix, m_PhysicalAspect.Outputs, "Plate3");

                        //Casing
                        constant.casing.ShapeLength = constant.length - constant.plate1.thickness1 - constant.plate2.thickness1 - constant.plate3.thickness1 - constant.frame.length1;
                        matrix = new Matrix4X4();
                        matrix.Translate(new Vector(0, 0, constant.plate1.thickness1));
                        AddNut(constant.casing, matrix, m_PhysicalAspect.Outputs, "Casing");

                        //Gusset
                        matrix = new Matrix4X4();
                        matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0));
                        matrix.Translate(new Vector(constant.casing.ShapeWidth1 / 2, -constant.gussetSpacing / 2, constant.plate1.thickness1));
                        AddPlate(constant.gussetPlate, matrix, m_PhysicalAspect.Outputs, "Gusset1");

                        matrix = new Matrix4X4();
                        matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0));
                        matrix.Translate(new Vector(constant.casing.ShapeWidth1 / 2, constant.gussetSpacing / 2 + constant.gussetPlate.thickness1, constant.plate1.thickness1));
                        AddPlate(constant.gussetPlate, matrix, m_PhysicalAspect.Outputs, "Gusset2");

                        //Up Thrust Plates
                        matrix = new Matrix4X4();
                        matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0));
                        matrix.Translate(new Vector(-constant.casing.ShapeWidth1 / 2 - constant.frameOffset + constant.frame.width1 / 2 - constant.upThrustPlate.width1 / 2, constant.upThrustSpacing / 2 + constant.upThrustPlate.thickness1, rodTakeout - constant.upThrustPlate.length1 - constant.plate3.thickness1));
                        AddPlate(constant.upThrustPlate, matrix, m_PhysicalAspect.Outputs, "UpThrustA");

                        matrix = new Matrix4X4();
                        matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0));
                        matrix.Translate(new Vector(-constant.casing.ShapeWidth1 / 2 - constant.frameOffset + constant.frame.width1 / 2 - constant.upThrustPlate.width1 / 2, -constant.upThrustSpacing / 2, rodTakeout - constant.upThrustPlate.length1 - constant.plate3.thickness1));
                        AddPlate(constant.upThrustPlate, matrix, m_PhysicalAspect.Outputs, "UpThrustB");

                        matrix = new Matrix4X4();
                        matrix.Translate(new Vector(-constant.casing.ShapeWidth1 / 2 - constant.frameOffset + constant.frame.width1 / 2 - constant.upThrustPlate.width1 / 2 + constant.plate4Offset, -constant.plate4.length1 / 2, rodTakeout - constant.plate4.thickness1));
                        AddPlate(constant.plate4, matrix, m_PhysicalAspect.Outputs, "Plate4");
                        break;
                    case 3:     //Horizontal Casing Beside Frame
                        //Frame
                        matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0));
                        matrix.Translate(new Vector(-constant.frame.width1 / 2, constant.frameSpacing / 2 + constant.frame.thickness1, constant.plate3.thickness1));
                        AddPlate(constant.frame, matrix, m_PhysicalAspect.Outputs, "FrameA");

                        matrix = new Matrix4X4();
                        matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0));
                        matrix.Translate(new Vector(-constant.frame.width1 / 2, -constant.frameSpacing / 2, constant.plate3.thickness1));
                        AddPlate(constant.frame, matrix, m_PhysicalAspect.Outputs, "FrameB");

                        //Plates
                        matrix = new Matrix4X4();
                        matrix.Rotate(Math.PI / 2, new Vector(0, 1, 0));
                        matrix.Translate(new Vector(constant.frame.width1 / 2 - constant.length, -constant.plate1.length1 / 2, constant.plate3.thickness1 + constant.frame.length1 + constant.plate4.thickness1 - constant.frameOffset + constant.plate1Offset));
                        AddPlate(constant.plate1, matrix, m_PhysicalAspect.Outputs, "Plate1");

                        matrix = new Matrix4X4();
                        matrix.Rotate(Math.PI / 2, new Vector(0, 1, 0));
                        matrix.Translate(new Vector(-constant.frame.width1 / 2 - constant.plate2.thickness1, -constant.plate2.length1 / 2, constant.plate3.thickness1 + constant.frame.length1 + constant.plate4.thickness1 - constant.frameOffset + constant.plate2Offset));
                        AddPlate(constant.plate2, matrix, m_PhysicalAspect.Outputs, "Plate2");

                        matrix = new Matrix4X4();
                        matrix.Translate(new Vector(-constant.frame.width1 / 2 + constant.plate3Offset, -constant.plate3.length1 / 2, 0));
                        AddPlate(constant.plate3, matrix, m_PhysicalAspect.Outputs, "Plate3");

                        matrix = new Matrix4X4();
                        matrix.Translate(new Vector(-constant.upThrustPlate.width1 / 2 + constant.plate4Offset, -constant.plate4.length1 / 2, rodTakeout - constant.plate4.thickness1));
                        AddPlate(constant.plate4, matrix, m_PhysicalAspect.Outputs, "Plate4");

                        //Casing
                        constant.casing.ShapeLength = constant.length - constant.plate1.thickness1 - constant.plate2.thickness1 - constant.frame.width1;
                        matrix = new Matrix4X4();
                        matrix.Rotate(Math.PI / 2, new Vector(0, 1, 0));
                        matrix.Translate(new Vector(constant.frame.width1 / 2 - constant.length + constant.plate1.thickness1, 0, constant.plate3.thickness1 + constant.frame.length1 + constant.plate4.thickness1 - constant.frameOffset - constant.casing.ShapeWidth1 / 2));
                        AddNut(constant.casing, matrix, m_PhysicalAspect.Outputs, "Casing");

                        //Up Thrust Plates
                        matrix = new Matrix4X4();
                        matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0));
                        matrix.Translate(new Vector(-constant.upThrustPlate.width1 / 2, constant.upThrustSpacing / 2 + constant.upThrustPlate.thickness1, rodTakeout - constant.upThrustPlate.length1 - constant.plate4.thickness1));
                        AddPlate(constant.upThrustPlate, matrix, m_PhysicalAspect.Outputs, "UpThrustA");

                        matrix = new Matrix4X4();
                        matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0));
                        matrix.Translate(new Vector(-constant.upThrustPlate.width1 / 2, -constant.upThrustSpacing / 2, rodTakeout - constant.upThrustPlate.length1 - constant.plate4.thickness1));
                        AddPlate(constant.upThrustPlate, matrix, m_PhysicalAspect.Outputs, "UpThrustB");
                        break;
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
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of TypeU."));
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
                //System WCG Attributes
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
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrWeightCG, "Error in Weight CG of TypeU."));
                    return;
                }
            }
        }
        #endregion
    }
}

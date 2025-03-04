//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   VariableTypeF.cs
//    HSSmartPart,Ingr.SP3D.Content.Support.Symbols.VariableTypeF
//   Author       :  Rajeswari
//   Creation Date:  10/06/2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   10/06/2013   Rajeswari   CR-CP-222469-Initial Creation
//   11-02-2013   Chethan     DI-CP-263820  Fix priority 3 items to .net SmartParts as a result of new testing  
//   10-06-2015   PVK	      TR-CP-274155	SmartPart TDL Errors should be corrected.
//   25-11-2015   PVK	      CR-CP-284435	VariableOutput must be added to Variable Springs in HSSmartPart project
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
using System;
using System.Collections.Generic;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;

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
    [SymbolVersion("1.0.0.0")]
    [VariableOutputs]
    [CacheOption(CacheOptionType.NonCached)]
    public class VariableTypeF : SmartPartComponentDefinition, ICustomWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "VariableSpring,Ingr.SP3D.Content.Support.Symbols.TypeF"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "Height1", "Height1", 0)]
        public InputDouble m_dHeight1;
        [InputDouble(3, "Offset1", "Offset1", 0)]
        public InputDouble m_dOffset1;
        [InputDouble(4, "Length1", "Length1", 0)]
        public InputDouble m_dLength1;
        [InputDouble(5, "ColumnHeight", "ColumnHeight", 0)]
        public InputDouble m_dColumnHeight;
        [InputDouble(6, "ColumnHeightMin", "ColumnHeightMin", 0)]
        public InputDouble m_dColumnHeightMin;
        [InputDouble(7, "ColumnHeightMax", "ColumnHeightMax", 0)]
        public InputDouble m_dColumnHeightMax;
        [InputDouble(8, "PipeOD", "PipeOD", 0)]
        public InputDouble m_dPipeOD;
        [InputDouble(9, "ShoeHeight", "ShoeHeight", 0)]
        public InputDouble m_dShoeHeight;
        [InputDouble(10, "Hole1HInset", "Hole1HInset", 0)]
        public InputDouble m_dHole1HInset;
        [InputDouble(11, "Hole1VInset", "Hole1VInset", 0)]
        public InputDouble m_dHole1VInset;
        [InputDouble(12, "OperatingLoad", "OperatingLoad", 0)]
        public InputDouble m_dOperatingLoad;
        [InputDouble(13, "Movement", "Movement", 0)]
        public InputDouble m_dMovement;
        [InputDouble(14, "MovementDirection", "MovementDirection",1)]
        public InputDouble m_oMovementDirection;
        [InputDouble(15, "SpringRate", "SpringRate", 0)]
        public InputDouble m_dSpringRate;
        [InputDouble(16, "MaxVariability", "MaxVariability", 0)]
        public InputDouble m_dMaxVariability;
        [InputDouble(17, "MinWorkingLoad", "MinWorkingLoad", 0)]
        public InputDouble m_dMinWorkingLoad;
        [InputDouble(18, "MaxWorkingLoad", "MaxWorkingLoad", 0)]
        public InputDouble m_dMaxWorkingLoad;
        [InputDouble(19, "MinOvertravelLoad", "MinOvertravelLoad", 0)]
        public InputDouble m_dMinOvertravelLoad;
        [InputDouble(20, "MaxOvertravelLoad", "MaxOvertravelLoad", 0)]
        public InputDouble m_dMaxOvertravelLoad;
        [InputDouble(21, "MinLimitStopTravel", "MinLimitStopTravel", 0)]
        public InputDouble m_dMinLimitStopTravel;
        [InputDouble(22, "MaxLimitStopTravel", "MaxLimitStopTravel", 0)]
        public InputDouble m_dMaxLimitStopTravel;
        [InputString(23, "Plate1Shape", "Plate1Shape","No Value")]
        public InputString m_oPlate1Shape;
        [InputString(24, "Plate2Shape", "Plate2Shape","No Value")]
        public InputString m_oPlate2Shape;
        [InputString(25, "Plate3Shape", "Plate3Shape","No Value")]
        public InputString m_oPlate3Shape;
        [InputString(26, "CasingShape", "CasingShape","No Value")]
        public InputString m_oCasingShape;
        [InputString(27, "ColumnShape", "ColumnShape","No Value")]
        public InputString m_oColumnShape;
        [InputString(28, "ColumnEndShape", "ColumnEndShape","No Value")]
        public InputString m_oColumnEndShape;
        [InputString(29, "TopAttachment", "TopAttachment","No Value")]
        public InputString m_oTopAttachment;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Route", "Route")]
        [SymbolOutput("Structure", "Structure")]
        [SymbolOutput("Surface", "Surface")]
        [SymbolOutput("Hole1", "Hole1")]
        [SymbolOutput("Hole2", "Hole2")]
        [SymbolOutput("Hole3", "Hole3")]
        [SymbolOutput("Hole4", "Hole4")]
        public AspectDefinition m_Symbolic;

        #endregion
        #region "Definition of Additional Outputs"
        public override IEnumerable<OutputDefinition> AdditionalOutputs(string aspectName)
        {
            List<OutputDefinition> additionalOutputs = new List<OutputDefinition>();
            // Add Outputs for the Plate Shapes
            AddPlateOutputs(3, additionalOutputs);
            // Add Outputs for the Load Column, and the Column End Shape
            AddRod1Outputs("Column", additionalOutputs);
            additionalOutputs.Add(new OutputDefinition("ColumnEnd", "ColumnEnd"));
            // Add Outputs for the Spring Casing
            additionalOutputs.Add(new OutputDefinition("Casing", "Casing"));
            // Add Possible Outputs for the Top Attachments
            additionalOutputs.Add(new OutputDefinition("TopPlate", "TopPlate"));
            AddShieldOutputs(additionalOutputs, "Repad");
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
                Part part = m_PartInput.Value as Part;

                Double height1 = m_dHeight1.Value;
                Double offset1 = m_dOffset1.Value;
                Double length1 = m_dLength1.Value;
                Double columnHeight = m_dColumnHeight.Value;
                Double columnHeightMin = m_dColumnHeightMin.Value;
                Double columnHeightMax = m_dColumnHeightMax.Value;
                Double pipeOD = m_dPipeOD.Value;
                Double shoeHeight = m_dShoeHeight.Value;
                Double hole1HInset = m_dHole1HInset.Value;
                Double hole1VInset = m_dHole1VInset.Value;
                Double operatingLoad = m_dOperatingLoad.Value;
                Double movement = m_dMovement.Value;
                int movementDirection = (int)m_oMovementDirection.Value;
                Double springRate = m_dSpringRate.Value;
                Double maxVariability = m_dMaxVariability.Value;
                Double minWorkingLoad = m_dMinWorkingLoad.Value;
                Double maxWorkingLoad = m_dMaxWorkingLoad.Value;
                Double minOvertravelLoad = m_dMinOvertravelLoad.Value;
                Double maxOvertravelLoad = m_dMaxOvertravelLoad.Value;
                Double minLimitStopTravel = m_dMinLimitStopTravel.Value;
                Double maxLimitStopTravel = m_dMaxLimitStopTravel.Value;

                //=================================================
                //Construction of Physical Aspect 
                //=================================================

                if(columnHeight < columnHeightMin || columnHeight > columnHeightMax)
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrInvalidnSpringrateValueGTZero, "Adjustable Column Height is outside the allowable range."));

                if (maxVariability <= 0)
                    maxVariability = 0.25;

                // Get the Shape Name Attribute Values
                string plate1 = m_oPlate1Shape.Value;
                string plate2 = m_oPlate2Shape.Value;
                string plate3 = m_oPlate3Shape.Value;
                string casing = m_oCasingShape.Value;
                string column = m_oColumnShape.Value;
                string columnEnd = m_oColumnEndShape.Value;
                string topAttachment = m_oTopAttachment.Value;

                // Load The Shape Data Structures
                PlateInputs plate1Inputs = LoadPlateDataByQuery(plate1);
                PlateInputs plate2Inputs = LoadPlateDataByQuery(plate2);
                PlateInputs plate3Inputs = LoadPlateDataByQuery(plate3);
                Rod1Inputs columnInputs = LoadRodDataByQuery(column);
                NutInputs columnEndInputs = LoadNutDataByQuery(columnEnd, 1);
                NutInputs casingInputs = LoadNutDataByQuery(casing, 1);

                // Override Shape Inputs
                casingInputs.ShapeLength = height1 - plate1Inputs.thickness1 - plate2Inputs.thickness1 - plate3Inputs.thickness1;
                columnInputs.length = length1 - columnEndInputs.ShapeLength + columnHeight - columnHeightMin;

                // Calculate the Output Engineering Attributes
                Double installedLoad = 0, variability = 0, operatingTravel = 0, installedTravel = 0, installedColumnHeight = 0;

                if (springRate <= 0)
                {
                    RaiseSmartPartTDLError(SmartPartSymbolResourceIDs.ErrInvalidnSpringrateValueGTZero, "The Spring Rate must be greater than zero.");
                    return;
                }

                // Installed Load
                switch (movementDirection)
                {
                    case 1:  // UP
                        installedLoad = operatingLoad + springRate * movement;
                        break;
                    case 2:  // DOWN
                        installedLoad = operatingLoad - springRate * movement;
                        break;
                }
                // Check Operating Load
                // added the condition 'And dOperatingLoad > 0' so that if the load has not been set the warnings will not be displayed
                // default load in the xls workbook can be -1 or a valid load
                if (operatingLoad < minWorkingLoad && operatingLoad >= 0)
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrInvalidnMinOperatingLoad, "The Operating Load is below the minimum working range."));
                else if (operatingLoad > maxWorkingLoad)
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrInvalidnMaxOperatingLoad, "The Operating Load is above the maximum working range."));

                // Check Installed Load
                // added the condition 'And dOperatingLoad > 0' so that if the load has not been set the warnings will not be displayed
                // default load in the xls workbook can be -1 or a valid load
                if (installedLoad < minWorkingLoad && operatingLoad >= 0)
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrInvalidnMinInstalledLoad, "The Installed Load is below the minimum working range."));
                else if (installedLoad > maxWorkingLoad)
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrInvalidnMaxInstalledLoad, "The Installed Load is above the maximum working range."));

                // Variability
                if (operatingLoad > 0)
                    variability = (movement * springRate) / operatingLoad;
                else
                    variability = 0;

                // Check Variability
                if (variability > maxVariability)
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrInvalidnVariability, "Variability has exceeded the maximum allowable value."));

                // Installed Travel
                if (installedLoad < minOvertravelLoad)
                    installedTravel = (minOvertravelLoad - minWorkingLoad) / springRate;
                else if (installedLoad > maxOvertravelLoad)
                    installedTravel = (maxOvertravelLoad - minWorkingLoad) / springRate;
                else
                    installedTravel = (installedLoad - minWorkingLoad) / springRate;

                // Adjust the Installed Travel to Account for Limit Stops
                if (minLimitStopTravel > (minOvertravelLoad - minWorkingLoad) / springRate)
                {
                    if (installedTravel < minLimitStopTravel)
                        installedTravel = minLimitStopTravel;
                }
                if (maxLimitStopTravel > (minOvertravelLoad - minWorkingLoad) / springRate)
                {
                    if (installedTravel > maxLimitStopTravel)
                        installedTravel = maxLimitStopTravel;
                }

                // Operating Travel
                if (operatingLoad < minOvertravelLoad)
                    operatingTravel = (minOvertravelLoad - minWorkingLoad) / springRate;
                else if (operatingLoad > maxOvertravelLoad)
                    operatingTravel = (maxOvertravelLoad - minWorkingLoad) / springRate;
                else
                    operatingTravel = (operatingLoad - minWorkingLoad) / springRate;

                // Adjust Operating Travel to Account for Limit Stops
                if (minLimitStopTravel > (minOvertravelLoad - minWorkingLoad) / springRate)
                {
                    if (operatingTravel < minLimitStopTravel)
                        operatingTravel = minLimitStopTravel;
                }
                if (maxLimitStopTravel > (minOvertravelLoad - minWorkingLoad) / springRate)
                {
                    if (operatingTravel > maxLimitStopTravel)
                        operatingTravel = maxLimitStopTravel;
                }

                // Installed Column Height
                installedColumnHeight = columnHeight - installedTravel;

                // Set the Output Engineering Attributes on the Part Occurence
                try
                {
                    SupportComponent supportComponent = (SupportComponent)Occurrence;
                    supportComponent.SetPropertyValue(installedLoad, "IJUAhsInstalledLoad", "InstalledLoad");
                    supportComponent.SetPropertyValue(variability, "IJUAhsVariability", "Variability");
                    supportComponent.SetPropertyValue(operatingTravel, "IJUAhsOperatingTravel", "OperatingTravel");
                    supportComponent.SetPropertyValue(installedTravel, "IJUAhsInstalledTravel", "InstalledTravel");
                    supportComponent.SetPropertyValue(installedColumnHeight, "IJUAhsInstalledHeight", "InstalledHeight");
                }
                catch
                {
                }
                
                Double holeXOffset = 0, holeYOffset = 0, holeZOffset = 0;

                if (plate2 == "" || plate2 == "No Value")
                {
                    holeXOffset = plate3Inputs.width1 / 2 - hole1HInset;
                    holeYOffset = plate3Inputs.length1 / 2 - hole1VInset;
                    holeZOffset = plate3Inputs.thickness1;
                }
                else
                {
                    if (plate3 != "" && plate3 != "No Value")
                    {
                        holeXOffset = plate3Inputs.width1 / 2 - hole1HInset;
                        holeYOffset = plate3Inputs.length1 / 2 - hole1VInset;
                        holeZOffset = plate3Inputs.thickness1;
                    }
                    else
                    {
                        holeXOffset = plate2Inputs.width1 / 2 - hole1HInset;
                        holeYOffset = plate2Inputs.length1 / 2 - hole1VInset;
                        holeZOffset = plate2Inputs.thickness1;
                    }
                }

                // Add the Ports
                Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, - height1 / 2 + installedColumnHeight + offset1 + shoeHeight + pipeOD / 2), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Route"] = port1;
                Port port2 = new Port(OccurrenceConnection, part, "Structure", new Position(0, 0, -height1 / 2), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Structure"] = port2;
                Port port3 = new Port(OccurrenceConnection, part, "Surface", new Position(0, 0, -height1 / 2 + installedColumnHeight + offset1), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Surface"] = port3;
                try
                {
                    Port port4 = new Port(OccurrenceConnection, part, "Hole1", new Position(holeXOffset, holeYOffset, -height1 / 2 + holeZOffset), new Vector(1, 0, 0), new Vector(0, 0, 1));
                    m_Symbolic.Outputs["Hole1"] = port4;
                }
                catch
                {
                    // not available for this refData
                }
                try
                {
                    Port port5 = new Port(OccurrenceConnection, part, "Hole2", new Position(- holeXOffset, holeYOffset, -height1 / 2 + holeZOffset), new Vector(1, 0, 0), new Vector(0, 0, 1));
                    m_Symbolic.Outputs["Hole2"] = port5;
                }
                catch
                {
                    // not available for this refData
                }
                try
                {
                    Port port6 = new Port(OccurrenceConnection, part, "Hole3", new Position(-holeXOffset, -holeYOffset, -height1 / 2 + holeZOffset), new Vector(1, 0, 0), new Vector(0, 0, 1));
                    m_Symbolic.Outputs["Hole3"] = port6;
                }
                catch
                {
                    // not available for this refData
                }
                try
                {
                    Port port7 = new Port(OccurrenceConnection, part, "Hole4", new Position(holeXOffset, -holeYOffset, -height1 / 2 + holeZOffset), new Vector(1, 0, 0), new Vector(0, 0, 1));
                    m_Symbolic.Outputs["Hole4"] = port7;
                }
                catch
                {
                    // not available for this refData
                }

                // Add the Graphics
                Matrix4X4 matrix = new Matrix4X4();
                if (plate1 != "" && plate1 != "No Value")
                {
                    matrix.Translate(new Vector(-plate1Inputs.width1 / 2, -plate1Inputs.length1 / 2, height1 / 2 - plate1Inputs.thickness1));
                    AddPlate(plate1Inputs, matrix, m_Symbolic.Outputs, "Plate1");
                }

                if (plate2 != "" && plate2 != "No Value")
                {
                    matrix = new Matrix4X4();
                    matrix.Translate(new Vector(-plate2Inputs.width1 / 2, -plate2Inputs.length1 / 2, -height1 / 2 + plate3Inputs.thickness1));
                    AddPlate(plate2Inputs, matrix, m_Symbolic.Outputs, "Plate2");
                }

                if (plate3 != "" && plate3 != "No Value")
                {
                    matrix = new Matrix4X4();
                    matrix.Translate(new Vector(-plate3Inputs.width1 / 2, -plate3Inputs.length1 / 2, -height1 / 2));
                    AddPlate(plate3Inputs, matrix, m_Symbolic.Outputs, "Plate3");
                }

                if (casing != "" && casing != "No Value")
                {
                    if (casingInputs.ShapeType == 3)
                    {
                        matrix = new Matrix4X4();
                        matrix.Rotate(Math.PI / 6, new Vector(0, 0, 1));
                        matrix.Translate(new Vector(0, 0, -height1 / 2 + plate2Inputs.thickness1 + plate3Inputs.thickness1));
                        AddNut(casingInputs, matrix, m_Symbolic.Outputs, "Casing");
                    }
                    else
                    {
                        matrix = new Matrix4X4();
                        matrix.Translate(new Vector(0, 0, -height1 / 2 + plate2Inputs.thickness1 + plate3Inputs.thickness1));
                        AddNut(casingInputs, matrix, m_Symbolic.Outputs, "Casing");
                    }
                }

                if (columnEnd != "" && columnEnd != "No Value")
                {
                    matrix = new Matrix4X4();
                    matrix.Translate(new Vector(0, 0, - height1 / 2 + installedColumnHeight - columnEndInputs.ShapeLength));
                    AddNut(columnEndInputs, matrix, m_Symbolic.Outputs, "ColumnEnd");
                }

                if (column != "" && column != "No Value")
                {
                    matrix = new Matrix4X4();
                    matrix.Rotate(Math.PI, new Vector(1, 0, 0));
                    matrix.Translate(new Vector(0, 0, -height1 / 2 + installedColumnHeight - columnEndInputs.ShapeLength));
                    AddRod(columnInputs, matrix, m_Symbolic.Outputs, "Column");
                }

                // Top Attachment - Currently handles Plate and Shield.  Roller to be added later.
                if (topAttachment != "" && topAttachment != "No Value")
                {
                    int topAttachmentValue=0;
                    Catalog m_oPlantCatalog = MiddleServiceProvider.SiteMgr.ActiveSite.ActivePlant.PlantCatalog;
                    BusinessObject businessObject = m_oPlantCatalog.GetNamedObject(topAttachment);

                    topAttachmentValue = (int)((PropertyValueCodelist)businessObject.GetPropertyValue("IJUAhsSmartShapeType", "SmartShapeType")).PropValue;
                    
                    switch(topAttachmentValue)
                    {
                        case 0:  // No Top Shape, or Undefined Top Shape
                            break;
                        case 5:  // Standard Plate Top Shape
                            {
                                PlateInputs topPlateInputs = LoadPlateDataByQuery(topAttachment);
                                matrix = new Matrix4X4();
                                matrix.Translate(new Vector(-topPlateInputs.width1 / 2, -topPlateInputs.length1 / 2, -height1 / 2 + installedColumnHeight));
                                AddPlate(topPlateInputs, matrix, m_Symbolic.Outputs, "TopPlate");
                            }
                            break;
                        case 75:  // Shield/Repad Top Shape
                            {
                                ShieldInputs repadInputs = LoadSheildDataByQuery(topAttachment);
                                matrix = new Matrix4X4();
                                matrix.Translate(new Vector(0, 0, -height1 / 2 + installedColumnHeight + repadInputs.Thickness1 + pipeOD / 2));
                                AddShield(repadInputs, matrix, m_Symbolic.Outputs, "Repad");
                            }
                            break;
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
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of VariableTypeF.cs."));
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

                Part catalogPart = (Part)supportComponentBO.GetRelationship("madeFrom", "part").TargetObjects[0];

                Double weight, cogX, cogY, cogZ;
                try
                {
                    weight = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryWeight")).PropValue;
                }
                catch
                {
                    weight = 0;
                }
                //Center of Gravity
                try
                {
                    cogX = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryCogX")).PropValue;
                }
                catch
                {
                    cogX = 0;
                }
                try
                {
                    cogY = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryCogY")).PropValue;
                }
                catch
                {
                    cogY = 0;
                }
                try
                {
                    cogZ = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryCogZ")).PropValue;
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
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrWeightCG, "Error in Weight CG of VariableTypeF.cs."));
                }
            }
        }

        #endregion

    }

}

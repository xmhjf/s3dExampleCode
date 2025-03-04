//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   VariableTypeE.cs
//    HSSmartPart,Ingr.SP3D.Content.Support.Symbols.VariableTypeE
//   Author       :  Rajeswari
//   Creation Date:  10/06/2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   10/06/2013   Rajeswari   CR-CP-222469-Initial Creation
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
    public class VariableTypeE : SmartPartComponentDefinition, ICustomWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "VariableSpring,Ingr.SP3D.Content.Support.Symbols.TypeE"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "RodDiameter", "RodDiameter", 0)]
        public InputDouble m_dRodDiameter;
        [InputDouble(3, "Height1", "Height1", 0)]
        public InputDouble m_dHeight1;
        [InputDouble(4, "Offset1", "Offset1", 0)]
        public InputDouble m_dOffset1;
        [InputDouble(5, "Offset2", "Offset2", 0)]
        public InputDouble m_dOffset2;
        [InputDouble(6, "RodLength", "RodLength", 0)]
        public InputDouble m_dRodLength;
        [InputDouble(7, "RodAddIn", "RodAddIn", 0)]
        public InputDouble m_dRodAddIn;
        [InputDouble(8, "OperatingLoad", "OperatingLoad", 0)]
        public InputDouble m_dOperatingLoad;
        [InputDouble(9, "Movement", "Movement", 0)]
        public InputDouble m_dMovement;
        [InputDouble(10, "MovementDirection", "MovementDirection",1)]
        public InputDouble m_oMovementDirection;
        [InputDouble(11, "SpringRate", "SpringRate", 0)]
        public InputDouble m_dSpringRate;
        [InputDouble(12, "MaxVariability", "MaxVariability", 0)]
        public InputDouble m_dMaxVariability;
        [InputDouble(13, "MinWorkingLoad", "MinWorkingLoad", 0)]
        public InputDouble m_dMinWorkingLoad;
        [InputDouble(14, "MaxWorkingLoad", "MaxWorkingLoad", 0)]
        public InputDouble m_dMaxWorkingLoad;
        [InputDouble(15, "MinOvertravelLoad", "MinOvertravelLoad", 0)]
        public InputDouble m_dMinOvertravelLoad;
        [InputDouble(16, "MaxOvertravelLoad", "MaxOvertravelLoad", 0)]
        public InputDouble m_dMaxOvertravelLoad;
        [InputDouble(17, "MinLimitStopTravel", "MinLimitStopTravel", 0)]
        public InputDouble m_dMinLimitStopTravel;
        [InputDouble(18, "MaxLimitStopTravel", "MaxLimitStopTravel", 0)]
        public InputDouble m_dMaxLimitStopTravel;
        [InputString(19, "Plate1Shape", "Plate1Shape", "No Value")]
        public InputString m_oPlate1Shape;
        [InputString(20, "Plate2Shape", "Plate2Shape", "No Value")]
        public InputString m_oPlate2Shape;
        [InputString(21, "CasingShape", "CasingShape", "No Value")]
        public InputString m_oCasingShape;
        [InputString(22, "TurnbuckleShape", "TurnbuckleShape", "No Value")]
        public InputString m_oTurnbuckleShape;
        [InputString(23, "NutShape", "NutShape", "No Value")]
        public InputString m_oNutShape;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("RodEnd1", "RodEnd1")]
        [SymbolOutput("RodEnd2", "RodEnd2")]
        [SymbolOutput("RodEnd3", "RodEnd3")]
        [SymbolOutput("Surface1", "Surface1")]
        [SymbolOutput("Structure", "Structure")]
        public AspectDefinition m_Symbolic;

        #endregion
        #region "Definition of Additional Outputs"
        public override IEnumerable<OutputDefinition> AdditionalOutputs(string aspectName)
        {
            List<OutputDefinition> additionalOutputs = new List<OutputDefinition>();
            // Add Outputs for the Nut Shapes
            AddNutOutputs(additionalOutputs);
            // Add Outputs for the Plate Shapes
            AddPlateOutputs(2, additionalOutputs);
            // Add Outputs for the Spring Casing
            additionalOutputs.Add(new OutputDefinition("Casing", "Casing"));
            // Add Outputs for the Turnbuckle / Coupling
            AddTurnbuckleOutputs(additionalOutputs, "Turnbuckle");

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

                Double rodDiameter = m_dRodDiameter.Value;
                Double height1 = m_dHeight1.Value;
                Double offset1 = m_dOffset1.Value;
                Double offset2 = m_dOffset2.Value;
                Double rodLength = m_dRodLength.Value;
                Double rodAddIn = m_dRodAddIn.Value;
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

                if (maxVariability <= 0)
                    maxVariability = 0.25;

                // Get the Shape Name Attribute Values
                string nut = m_oNutShape.Value;
                string plate1 = m_oPlate1Shape.Value;
                string plate2 = m_oPlate2Shape.Value;
                string casing = m_oCasingShape.Value;
                string turnbuckle = m_oTurnbuckleShape.Value;

                // Load The Shape Data Structures
                NutInputs nutInputs = LoadNutDataByQuery(nut, 1);
                PlateInputs plate1Inputs = LoadPlateDataByQuery(plate1);
                PlateInputs plate2Inputs = LoadPlateDataByQuery(plate2);
                NutInputs casingInputs = LoadNutDataByQuery(casing, 1);
                TurnbuckleInputs turnbuckleInputs = LoadTurnbuckleDataByQuery(turnbuckle);

                Double turnbuckleLength = turnbuckleInputs.Opening1 + turnbuckleInputs.Nut.ShapeLength * 2;
                Double turnbuckleTakeOut = turnbuckleLength - offset1 - offset2;

                // Override Shape Inputs
                casingInputs.ShapeLength = height1 - plate1Inputs.thickness1 - plate2Inputs.thickness1;

                // Calculate the Output Engineering Attributes
                Double installedLoad = 0, variability = 0, operatingTravel = 0, installedTravel = 0, installedRodAddIn = 0;

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

                // Installed Rod Add In
                installedRodAddIn = rodAddIn - installedTravel;

                // Set the Output Engineering Attributes on the Part Occurence
                try
                {
                    SupportComponent supportComponent = (SupportComponent)Occurrence;
                    supportComponent.SetPropertyValue(installedLoad, "IJUAhsInstalledLoad", "InstalledLoad");
                    supportComponent.SetPropertyValue(variability, "IJUAhsVariability", "Variability");
                    supportComponent.SetPropertyValue(operatingTravel, "IJUAhsOperatingTravel", "OperatingTravel");
                    supportComponent.SetPropertyValue(installedTravel, "IJUAhsInstalledTravel", "InstalledTravel");
                    supportComponent.SetPropertyValue(installedRodAddIn, "IJUAhsInstalledRodAddIn", "InstalledRodAddIn");
                }
                catch
                {
                }

                // Add the Ports
                Port port1 = new Port(OccurrenceConnection, part, "RodEnd1", new Position(0, 0, - height1 / 2 + installedRodAddIn - rodLength - turnbuckleTakeOut), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["RodEnd1"] = port1;
                Port port2 = new Port(OccurrenceConnection, part, "RodEnd2", new Position(0, 0, -height1 / 2 + installedRodAddIn - rodLength), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["RodEnd2"] = port2;
                Port port3 = new Port(OccurrenceConnection, part, "RodEnd3", new Position(0, 0, - height1 / 2 + installedRodAddIn), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["RodEnd3"] = port3;
                Port port4 = new Port(OccurrenceConnection, part, "Surface1", new Position(0, 0, -height1 / 2 + installedRodAddIn - rodLength - turnbuckleTakeOut - offset2), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Surface1"] = port4;
                Port port5 = new Port(OccurrenceConnection, part, "Structure", new Position(0, 0, - height1 / 2), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Structure"] = port5;
                // Add the Graphics
                Matrix4X4 matrix = new Matrix4X4();
                if (nut != "" && nut != "No Value")
                {
                    matrix.Translate(new Vector(0, 0, - height1 / 2 + installedRodAddIn - rodLength + offset1 - turnbuckleLength - nutInputs.ShapeLength));
                    AddNut(nutInputs, matrix, m_Symbolic.Outputs, "Nut");
                }

                if (plate1 != "" && plate1 != "No Value")
                {
                    matrix = new Matrix4X4();
                    matrix.Translate(new Vector(-plate1Inputs.width1 / 2, -plate1Inputs.length1 / 2, height1 / 2 - plate1Inputs.thickness1));
                    AddPlate(plate1Inputs, matrix, m_Symbolic.Outputs, "Plate1");
                }
                if (plate2 != "" && plate2 != "No Value")
                {
                    matrix = new Matrix4X4();
                    matrix.Translate(new Vector(-plate2Inputs.width1 / 2, -plate2Inputs.length1 / 2, -height1 / 2));
                    AddPlate(plate2Inputs, matrix, m_Symbolic.Outputs, "Plate2");
                }
                if (casing != "" && casing != "No Value")
                {
                    if (casingInputs.ShapeType == 3)
                    {
                        matrix = new Matrix4X4();
                        matrix.Rotate(Math.PI / 6, new Vector(0, 0, 1));
                        matrix.Translate(new Vector(0, 0, -height1 / 2 + plate2Inputs.thickness1));
                        AddNut(casingInputs, matrix, m_Symbolic.Outputs, "Casing");
                    }
                    else
                    {
                        matrix = new Matrix4X4();
                        matrix.Translate(new Vector(0, 0, -height1 / 2 + plate2Inputs.thickness1));
                        AddNut(casingInputs, matrix, m_Symbolic.Outputs, "Casing");
                    }
                }

                if (turnbuckle != "" && turnbuckle != "No Value")
                {
                    matrix = new Matrix4X4();
                    matrix.Translate(new Vector(0, 0, - height1 / 2 + installedRodAddIn - rodLength - turnbuckleLength / 2 +offset1));
                    AddTurnbuckle(turnbuckleInputs, 0, matrix, m_Symbolic.Outputs, "Turnbuckle");
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
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of VariableTypeE.cs."));
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
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrWeightCG, "Error in Weight CG of VariableTypeE.cs."));
                }
            }
        }

        #endregion

    }

}

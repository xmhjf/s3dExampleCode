//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   VariableTypeGA.cs
//    HSSmartPart,Ingr.SP3D.Content.Support.Symbols.VariableTypeGA
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
using System.Linq;
using System.Collections.ObjectModel;
using Ingr.SP3D.Structure.Middle.Services;
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
    public class VariableTypeGA : SmartPartComponentDefinition, ICustomWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "VariableSpring,Ingr.SP3D.Content.Support.Symbols.TypeGA"
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
        [InputDouble(6, "Offset3", "Offset3", 0)]
        public InputDouble m_dOffset3;
        [InputDouble(7, "Length1", "Length1", 0)]
        public InputDouble m_dLength1;
        [InputDouble(8, "RodTakeOut", "RodTakeOut", 0)]
        public InputDouble m_dRodTakeOut;
        [InputDouble(9, "RodTakeOut2", "RodTakeOut2", 0)]
        public InputDouble m_dRodTakeOut2;
        [InputDouble(10, "PipeOD", "PipeOD", 0)]
        public InputDouble m_dPipeOD;
        [InputDouble(11, "ShoeHeight", "ShoeHeight", 0)]
        public InputDouble m_dShoeHeight;
        [InputDouble(12, "BBGap", "BBGap", 0)]
        public InputDouble m_dBBGap;
        [InputDouble(13, "CC", "CC", 0)]
        public InputDouble m_dCC;
        [InputDouble(14, "CCMin", "CCMin", 0)]
        public InputDouble m_dCCMin;
        [InputDouble(15, "CCMax", "CCMax", 0)]
        public InputDouble m_dCCMax;
        [InputDouble(16, "OperatingLoad", "OperatingLoad", 0)]
        public InputDouble m_dOperatingLoad;
        [InputDouble(17, "Movement", "Movement", 0)]
        public InputDouble m_dMovement;
        [InputDouble(18, "MovementDirection", "MovementDirection",1)]
        public InputDouble m_oMovementDirection;
        [InputDouble(19, "SpringRate", "SpringRate", 0)]
        public InputDouble m_dSpringRate;
        [InputDouble(20, "MaxVariability", "MaxVariability", 0)]
        public InputDouble m_dMaxVariability;
        [InputDouble(21, "MinWorkingLoad", "MinWorkingLoad", 0)]
        public InputDouble m_dMinWorkingLoad;
        [InputDouble(22, "MaxWorkingLoad", "MaxWorkingLoad", 0)]
        public InputDouble m_dMaxWorkingLoad;
        [InputDouble(23, "MinOvertravelLoad", "MinOvertravelLoad", 0)]
        public InputDouble m_dMinOvertravelLoad;
        [InputDouble(24, "MaxOvertravelLoad", "MaxOvertravelLoad", 0)]
        public InputDouble m_dMaxOvertravelLoad;
        [InputDouble(25, "MinLimitStopTravel", "MinLimitStopTravel", 0)]
        public InputDouble m_dMinLimitStopTravel;
        [InputDouble(26, "MaxLimitStopTravel", "MaxLimitStopTravel", 0)]
        public InputDouble m_dMaxLimitStopTravel;
        [InputString(27, "Nut1Shape", "Nut1Shape", "No Value")]
        public InputString m_oNut1Shape;
        [InputString(28, "Nut2Shape", "Nut2Shape", "No Value")]
        public InputString m_oNut2Shape;
        [InputString(29, "Plate1Shape", "Plate1Shape", "No Value")]
        public InputString m_oPlate1Shape;
        [InputString(30, "Plate2Shape", "Plate2Shape", "No Value")]
        public InputString m_oPlate2Shape;
        [InputString(31, "CasingShape", "CasingShape", "No Value")]
        public InputString m_oCasingShape;
        [InputString(32, "ColumnShape", "ColumnShape", "No Value")]
        public InputString m_oColumnShape;
        [InputString(33, "ColumnEndShape", "ColumnEndShape", "No Value")]
        public InputString m_oColumnEndShape;
        [InputString(34, "TurnbuckleShape", "TurnbuckleShape", "No Value")]
        public InputString m_oTurnbuckleShape;
        [InputString(35, "Nut3Shape", "Nut3Shape", "No Value")]
        public InputString m_oNut3Shape;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Route", "Route")]
        [SymbolOutput("Surface", "Surface")]
        [SymbolOutput("RodEnd1", "RodEnd1")]
        [SymbolOutput("RodEnd2", "RodEnd2")]
        [SymbolOutput("Surface1", "Surface1")]
        [SymbolOutput("Surface2", "Surface2")]
        [SymbolOutput("RodA", "RodA")]
        [SymbolOutput("RodB", "RodB")]
        [SymbolOutput("Section1", "Section1")]
        [SymbolOutput("Section2", "Section2")]
        public AspectDefinition m_Symbolic;

        #endregion
        #region "Definition of Additional Outputs"
        public override IEnumerable<OutputDefinition> AdditionalOutputs(string aspectName)
        {
            List<OutputDefinition> additionalOutputs = new List<OutputDefinition>();
            // Add Outputs for the Nut Shapes
            additionalOutputs.Add(new OutputDefinition("Nut1A", "Nut1A"));
            additionalOutputs.Add(new OutputDefinition("Nut2A", "Nut2A"));
            additionalOutputs.Add(new OutputDefinition("Nut3A", "Nut3A"));
            additionalOutputs.Add(new OutputDefinition("Nut1B", "Nut1B"));
            additionalOutputs.Add(new OutputDefinition("Nut2B", "Nut2B"));
            additionalOutputs.Add(new OutputDefinition("Nut3B", "Nut3B"));
            // Add Outputs for the Plate Shapes
            additionalOutputs.Add(new OutputDefinition("Plate1A", "Plate1A"));
            additionalOutputs.Add(new OutputDefinition("Plate2A", "Plate2A"));
            additionalOutputs.Add(new OutputDefinition("Plate1B", "Plate1B"));
            additionalOutputs.Add(new OutputDefinition("Plate2B", "Plate2B"));
            // Add Outputs for the Load Column, and the Column End Shape
            AddRod1Outputs("ColumnA", additionalOutputs);
            additionalOutputs.Add(new OutputDefinition("ColumnEndA", "ColumnEndA"));
            AddRod1Outputs("ColumnB", additionalOutputs);
            additionalOutputs.Add(new OutputDefinition("ColumnEndB", "ColumnEndB"));
            // Add Outputs for the Turnbuckle / Coupling
            AddTurnbuckleOutputs(additionalOutputs, "TurnbuckleA");
            AddTurnbuckleOutputs(additionalOutputs, "TurnbuckleB");
            // Add Outputs for the Spring Casing
            additionalOutputs.Add(new OutputDefinition("CasingA", "CasingA"));
            additionalOutputs.Add(new OutputDefinition("CasingB", "CasingB"));

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
                Double offset3 = m_dOffset3.Value;
                Double length1 = m_dLength1.Value;
                Double rodTakeOut = m_dRodTakeOut.Value;
                Double rodTakeOut2 = m_dRodTakeOut2.Value;
                Double pipeOD = m_dPipeOD.Value;
                Double shoeHeight = m_dShoeHeight.Value;
                Double bBGap = m_dBBGap.Value;
                Double CC = m_dCC.Value;
                Double cCMin = m_dCCMin.Value;
                Double cCMax = m_dCCMax.Value;
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
                string nut1 = m_oNut1Shape.Value;
                string nut2 = m_oNut2Shape.Value;
                string nut3 = m_oNut3Shape.Value;
                string plate1 = m_oPlate1Shape.Value;
                string plate2 = m_oPlate2Shape.Value;
                string casing = m_oCasingShape.Value;
                string column = m_oColumnShape.Value;
                string columnEnd = m_oColumnEndShape.Value;
                string turnbuckle = m_oTurnbuckleShape.Value;

                // Load The Shape Data Structures
                NutInputs nut1Inputs = LoadNutDataByQuery(nut1, 1);
                NutInputs nut2Inputs = LoadNutDataByQuery(nut2, 1);
                NutInputs nut3Inputs = LoadNutDataByQuery(nut3, 1);
                PlateInputs plate1Inputs = LoadPlateDataByQuery(plate1);
                PlateInputs plate2Inputs = LoadPlateDataByQuery(plate2);
                Rod1Inputs columnInputs = LoadRodDataByQuery(column);
                NutInputs columnEndInputs = LoadNutDataByQuery(columnEnd, 1);
                TurnbuckleInputs turnbuckleInputs = LoadTurnbuckleDataByQuery(turnbuckle);
                NutInputs casingInputs = LoadNutDataByQuery(casing, 1);

                // Override Shape Inputs
                casingInputs.ShapeLength = height1 - plate1Inputs.thickness1 - plate2Inputs.thickness1;
                columnInputs.length = length1 - columnEndInputs.ShapeLength - turnbuckleInputs.Length2 + offset2;

                // Calculate the Output Engineering Attributes
                Double installedLoad = 0, variability = 0, operatingTravel = 0, installedTravel = 0, installedRodTakeOut = 0, installedRodTakeOut2 = 0;


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

                installedRodTakeOut = rodTakeOut + installedTravel;
                installedRodTakeOut2 = rodTakeOut2 + installedTravel;

                // Set the Output Engineering Attributes on the Part Occurence
                try
                {
                    SupportComponent supportComponent = (SupportComponent)Occurrence;
                    supportComponent.SetPropertyValue(installedLoad, "IJUAhsInstalledLoadG", "InstalledLoad");
                    supportComponent.SetPropertyValue(variability, "IJUAhsVariability", "Variability");
                    supportComponent.SetPropertyValue(operatingTravel, "IJUAhsOperatingTravel", "OperatingTravel");
                    supportComponent.SetPropertyValue(installedTravel, "IJUAhsInstalledTravel", "InstalledTravel");
                    supportComponent.SetPropertyValue(installedRodTakeOut, "IJUAhsInstalledRodTakeOut", "InstalledRodTakeOut");
                }
                catch
                {
                }

                // Add the Ports
                Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, shoeHeight + pipeOD / 2), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Route"] = port1;
                Port port2 = new Port(OccurrenceConnection, part, "Surface", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Surface"] = port2;
                Port port3 = new Port(OccurrenceConnection, part, "RodEnd1", new Position(0, CC/2, installedRodTakeOut), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["RodEnd1"] = port3;
                Port port4 = new Port(OccurrenceConnection, part, "RodEnd2", new Position(0, -CC / 2, installedRodTakeOut), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["RodEnd2"] = port4;
                Port port5 = new Port(OccurrenceConnection, part, "Surface1", new Position(0, CC / 2, installedRodTakeOut + offset1), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Surface1"] = port5;
                Port port6 = new Port(OccurrenceConnection, part, "Surface2", new Position(0, -CC / 2, installedRodTakeOut + offset1), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Surface2"] = port6;

                // Add the Graphics
                Matrix4X4 matrix = new Matrix4X4();
                if (nut1 != "" && nut1 != "No Value")
                {
                    matrix.Translate(new Vector(0, CC / 2, installedRodTakeOut + offset1));
                    AddNut(nut1Inputs, matrix, m_Symbolic.Outputs, "Nut1A");
                    matrix = new Matrix4X4();
                    matrix.Translate(new Vector(0, -CC / 2, installedRodTakeOut + offset1));
                    AddNut(nut1Inputs, matrix, m_Symbolic.Outputs, "Nut1B");
                }

                if (nut2 != "" && nut2 != "No Value")
                {
                    matrix.Translate(new Vector(0, CC / 2, installedRodTakeOut + offset1 + nut1Inputs.ShapeLength));
                    AddNut(nut2Inputs, matrix, m_Symbolic.Outputs, "Nut2A");
                    matrix = new Matrix4X4();
                    matrix.Translate(new Vector(0, -CC / 2, installedRodTakeOut + offset1 + nut1Inputs.ShapeLength));
                    AddNut(nut2Inputs, matrix, m_Symbolic.Outputs, "Nut2B");
                }

                if (plate1 != "" && plate1 != "No Value")
                {
                    matrix = new Matrix4X4();
                    matrix.Translate(new Vector(-plate1Inputs.width1 / 2, -plate1Inputs.length1 / 2 + CC / 2, installedRodTakeOut + offset1 - plate1Inputs.thickness1));
                    AddPlate(plate1Inputs, matrix, m_Symbolic.Outputs, "Plate1A");
                    matrix = new Matrix4X4();
                    matrix.Translate(new Vector(-plate1Inputs.width1 / 2, -plate1Inputs.length1 / 2 - CC / 2, installedRodTakeOut + offset1 - plate1Inputs.thickness1));
                    AddPlate(plate1Inputs, matrix, m_Symbolic.Outputs, "Plate1B");
                }
                if (plate2 != "" && plate2 != "No Value")
                {
                    matrix = new Matrix4X4();
                    matrix.Translate(new Vector(-plate2Inputs.width1 / 2, -plate2Inputs.length1 / 2 + CC / 2, installedRodTakeOut + offset1 - height1));
                    AddPlate(plate2Inputs, matrix, m_Symbolic.Outputs, "Plate2A");
                    matrix = new Matrix4X4();
                    matrix.Translate(new Vector(-plate2Inputs.width1 / 2, -plate2Inputs.length1 / 2 - CC / 2, installedRodTakeOut + offset1 - height1));
                    AddPlate(plate2Inputs, matrix, m_Symbolic.Outputs, "Plate2B");
                }

                if (casing != "" && casing != "No Value")
                {
                    matrix = new Matrix4X4();
                    matrix.Translate(new Vector(0, CC / 2, installedRodTakeOut + offset1 - height1 + plate2Inputs.thickness1));
                    AddNut(casingInputs, matrix, m_Symbolic.Outputs, "CasingA");
                    matrix = new Matrix4X4();
                    matrix.Translate(new Vector(0, -CC / 2, installedRodTakeOut + offset1 - height1 + plate2Inputs.thickness1));
                    AddNut(casingInputs, matrix, m_Symbolic.Outputs, "CasingB");
                }

                if (columnEnd != "" && columnEnd != "No Value")
                {
                    matrix = new Matrix4X4();
                    matrix.Translate(new Vector(0, CC / 2, installedRodTakeOut - installedRodTakeOut2 - offset2));
                    AddNut(columnEndInputs, matrix, m_Symbolic.Outputs, "ColumnEndA");
                    matrix = new Matrix4X4();
                    matrix.Translate(new Vector(0, -CC / 2, installedRodTakeOut - installedRodTakeOut2 - offset2));
                    AddNut(columnEndInputs, matrix, m_Symbolic.Outputs, "ColumnEndB");
                }
                if (column != "" && column != "No Value")
                {
                    if (turnbuckle != "" && turnbuckle != "No Value")
                    {
                        matrix = new Matrix4X4();
                        matrix.Translate(new Vector(0, CC / 2, installedRodTakeOut - installedRodTakeOut2 -offset2 + turnbuckleInputs.Length2 -offset2));
                        AddRod(columnInputs, matrix, m_Symbolic.Outputs, "ColumnA");
                        matrix = new Matrix4X4();
                        matrix.Translate(new Vector(0, -CC / 2, installedRodTakeOut - installedRodTakeOut2 - offset2 + turnbuckleInputs.Length2 - offset2));
                        AddRod(columnInputs, matrix, m_Symbolic.Outputs, "ColumnB");
                    }
                    else
                    {
                        matrix = new Matrix4X4();
                        matrix.Translate(new Vector(0, CC / 2, installedRodTakeOut - installedRodTakeOut2 - offset2 + columnEndInputs.ShapeLength));
                        AddRod(columnInputs, matrix, m_Symbolic.Outputs, "ColumnA");
                        matrix = new Matrix4X4();
                        matrix.Translate(new Vector(0, -CC / 2, installedRodTakeOut - installedRodTakeOut2 - offset2 + columnEndInputs.ShapeLength));
                        AddRod(columnInputs, matrix, m_Symbolic.Outputs, "ColumnB");
                    }
                }
                if (turnbuckle != "" && turnbuckle != "No Value")
                {
                    matrix = new Matrix4X4();
                    matrix.Translate(new Vector(0, CC / 2, installedRodTakeOut - installedRodTakeOut2));
                    AddTurnbuckle(turnbuckleInputs, turnbuckleInputs.Length2 - 2 * offset2, matrix, m_Symbolic.Outputs, "TurnbuckleA");
                    matrix = new Matrix4X4();
                    matrix.Translate(new Vector(0, -CC / 2, installedRodTakeOut - installedRodTakeOut2));
                    AddTurnbuckle(turnbuckleInputs, turnbuckleInputs.Length2 - 2 * offset2, matrix, m_Symbolic.Outputs, "TurnbuckleB");
                }
                if (nut3 != "" && nut3 != "No Value")
                {
                    matrix = new Matrix4X4();
                    matrix.Translate(new Vector(0, CC / 2, installedRodTakeOut - installedRodTakeOut2 - offset2 - nut3Inputs.ShapeLength));
                    AddNut(nut3Inputs, matrix, m_Symbolic.Outputs, "Nut3A");
                    matrix = new Matrix4X4();
                    matrix.Translate(new Vector(0, -CC / 2, installedRodTakeOut - installedRodTakeOut2 - offset2 - nut3Inputs.ShapeLength));
                    AddNut(nut3Inputs, matrix, m_Symbolic.Outputs, "Nut3B");
                }

                // Add the Extra Rods.
                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, CC / 2, 0);
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 0, 1).GetOrthogonalVector());
                Projection3d rodA = symbolGeometryHelper.CreateCylinder(null, rodDiameter / 2, rodTakeOut - rodTakeOut2);
                m_Symbolic.Outputs["RodA"] = rodA;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, -CC / 2, 0);
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 0, 1).GetOrthogonalVector());
                Projection3d rodB = symbolGeometryHelper.CreateCylinder(null, rodDiameter / 2, rodTakeOut - rodTakeOut2);
                m_Symbolic.Outputs["RodB"] = rodB;

                // Add the Steel Cross Section
                SweepOptions sweepOptions = (SweepOptions)1;
                CrossSectionServices crossSectionServices = new CrossSectionServices();
                CrossSection crossSection1, crossSection2;
                try
                {
                    crossSection1 = (CrossSection)part.GetRelationship("HgrCrossSection", "CrossSection").TargetObjects.First();                    
                }
                catch
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrCrossSectionNotFound, "Could not get Cross-section object."));
                    return;
                }

                // Calculate the Cut Angles
                Line3d projection;
                Double[] startPosition = new Double[3]; Double[] endPosition = new Double[3];
                startPosition[0] = 0;
                startPosition[1] = 0;
                startPosition[2] = 1;

                endPosition[0] = 0;
                endPosition[1] = 0;
                endPosition[2] = 1;

                Collection<ISurface> crossSecSurfaces1, crossSecSurfaces2;
                // here validating CC / 2 to only positive because it's creating wrong graphics, if it is negative or zero.
                if (CC / 2 > 0)
                {
                    if (bBGap >= 0)
                    {
                        crossSection2 = crossSection1;
                        // here bBGap / 2 is taken as X-coordinate to create the CrossSections with some gap similar to VB Graphic.
                        projection = new Line3d(new Position(bBGap / 2, -CC / 2 - offset3, 0), new Position(bBGap / 2, CC / 2 + offset3, 0));
                        // First Section
                        crossSecSurfaces1 = crossSectionServices.GetProjectionSurfacesFromCrossSection(crossSection1, projection, 7, false, 0, new Position(startPosition[0], startPosition[1], startPosition[2]), null, new Position(endPosition[0], endPosition[1], endPosition[2]), null, sweepOptions);
                        m_Symbolic.Outputs["Section1"] = crossSecSurfaces1[0];
                        // Second Section
                        projection = new Line3d(new Position(-bBGap / 2, -CC / 2 - offset3, 0), new Position(-bBGap / 2, CC / 2 + offset3, 0));
                        crossSecSurfaces2 = crossSectionServices.GetProjectionSurfacesFromCrossSection(crossSection2, projection, 7, true, 0, new Position(startPosition[0], startPosition[1], startPosition[2]), null, new Position(endPosition[0], endPosition[1], endPosition[2]), null, sweepOptions);
                        m_Symbolic.Outputs["Section2"] = crossSecSurfaces2[0];
                    }
                    else
                    {
                        projection = new Line3d(new Position(0, -CC / 2 - offset3, 0), new Position(0, CC / 2 + offset3, 0));
                        // First Section
                        crossSecSurfaces1 = crossSectionServices.GetProjectionSurfacesFromCrossSection(crossSection1, projection, 8, false, 0, new Position(startPosition[0], startPosition[1], startPosition[2]), null, new Position(endPosition[0], endPosition[1], endPosition[2]), null, sweepOptions);
                        m_Symbolic.Outputs["Section1"] = crossSecSurfaces1[0];
                    }
                }
                else
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrInvalidCC, "CC value should be greater than zero."));
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
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of VariableTypeGA.cs."));
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
                    Double weightValue, baseLength, weightPerLength, CC;
                    weightValue = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJUAhsWeight", "Weight")).PropValue;
                    baseLength = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJUAhsBaseLength", "BaseLength")).PropValue;
                    weightPerLength = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJUAhsWtPerLen", "WtPerLen")).PropValue;
                    CC = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJUAhsCC", "CC")).PropValue;
                    weight = weightValue + (CC - baseLength) * weightPerLength;
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
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrWeightCG, "Error in Weight CG of VariableTypeGA.cs."));
                }
            }
        }

        #endregion

    }

}

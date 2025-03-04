//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   VariableTypeG.cs
//    HSSmartPart,Ingr.SP3D.Content.Support.Symbols.VariableTypeG
//   Author       :  Rajeswari
//   Creation Date:  10/06/2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   10/06/2013   Rajeswari   CR-CP-222469-Initial Creation
//   10-06-2015   PVK	      TR-CP-274155	SmartPart TDL Errors should be corrected.
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
    public class VariableTypeG : SmartPartComponentDefinition, ICustomWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "VariableSpring,Ingr.SP3D.Content.Support.Symbols.TypeG"
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
        [InputDouble(6, "Length1", "Length1", 0)]
        public InputDouble m_dLength1;
        [InputDouble(7, "RodTakeOut", "RodTakeOut", 0)]
        public InputDouble m_dRodTakeOut;
        [InputDouble(8, "PipeOD", "PipeOD", 0)]
        public InputDouble m_dPipeOD;
        [InputDouble(9, "ShoeHeight", "ShoeHeight", 0)]
        public InputDouble m_dShoeHeight;
        [InputDouble(10, "BBGap", "BBGap", 0)]
        public InputDouble m_dBBGap;
        [InputDouble(11, "CC", "CC", 0)]
        public InputDouble m_dCC;
        [InputDouble(12, "CCMin", "CCMin", 0)]
        public InputDouble m_dCCMin;
        [InputDouble(13, "CCMax", "CCMax", 0)]
        public InputDouble m_dCCMax;
        [InputDouble(14, "OperatingLoad", "OperatingLoad", 0)]
        public InputDouble m_dOperatingLoad;
        [InputDouble(15, "Movement", "Movement", 0)]
        public InputDouble m_dMovement;
        [InputDouble(16, "MovementDirection", "MovementDirection", 1)]
        public InputDouble m_oMovementDirection;
        [InputDouble(17, "SpringRate", "SpringRate", 0)]
        public InputDouble m_dSpringRate;
        [InputDouble(18, "MaxVariability", "MaxVariability", 0)]
        public InputDouble m_dMaxVariability;
        [InputDouble(19, "MinWorkingLoad", "MinWorkingLoad", 0)]
        public InputDouble m_dMinWorkingLoad;
        [InputDouble(20, "MaxWorkingLoad", "MaxWorkingLoad", 0)]
        public InputDouble m_dMaxWorkingLoad;
        [InputDouble(21, "MinOvertravelLoad", "MinOvertravelLoad", 0)]
        public InputDouble m_dMinOvertravelLoad;
        [InputDouble(22, "MaxOvertravelLoad", "MaxOvertravelLoad", 0)]
        public InputDouble m_dMaxOvertravelLoad;
        [InputDouble(23, "MinLimitStopTravel", "MinLimitStopTravel", 0)]
        public InputDouble m_dMinLimitStopTravel;
        [InputDouble(24, "MaxLimitStopTravel", "MaxLimitStopTravel", 0)]
        public InputDouble m_dMaxLimitStopTravel;
        [InputString(25, "NutShape", "NutShape", "No Value")]
        public InputString m_oNutShape;
        [InputString(26, "Plate1Shape", "Plate1Shape", "No Value")]
        public InputString m_oPlate1Shape;
        [InputString(27, "Plate2Shape", "Plate2Shape", "No Value")]
        public InputString m_oPlate2Shape;
        [InputString(28, "CasingShape", "CasingShape", "No Value")]
        public InputString m_oCasingShape;
        [InputString(29, "ColumnShape", "ColumnShape", "No Value")]
        public InputString m_oColumnShape;
        [InputString(30, "ColumnEndShape", "ColumnEndShape", "No Value")]
        public InputString m_oColumnEndShape;
        [InputString(31, "TurnbuckleShape", "TurnbuckleShape", "No Value")]
        public InputString m_oTurnbuckleShape;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Route", "Route")]
        [SymbolOutput("Surface", "Surface")]
        [SymbolOutput("RodEnd1", "RodEnd1")]
        [SymbolOutput("RodEnd2", "RodEnd2")]
        [SymbolOutput("Surface1", "Surface1")]
        [SymbolOutput("Surface2", "Surface2")]
        [SymbolOutput("Section1", "Section1")]
        [SymbolOutput("Section2", "Section2")]
        public AspectDefinition m_Symbolic;

        #endregion
        #region "Definition of Additional Outputs"
        public override IEnumerable<OutputDefinition> AdditionalOutputs(string aspectName)
        {
            List<OutputDefinition> additionalOutputs = new List<OutputDefinition>();
            // Add Outputs for the Nut Shapes
            AddNutOutputs(2, additionalOutputs);
            // Add Outputs for the Plate Shapes
            AddPlateOutputs(4, additionalOutputs);
            // Add Outputs for the Load Column, and the Column End Shape
            AddRod1Outputs("Column1", additionalOutputs);
            additionalOutputs.Add(new OutputDefinition("ColumnEnd1", "ColumnEnd1"));
            AddRod1Outputs("Column2", additionalOutputs);
            additionalOutputs.Add(new OutputDefinition("ColumnEnd2", "ColumnEnd2"));
            // Add Outputs for the Turnbuckle / Coupling
            AddTurnbuckleOutputs(additionalOutputs, "Turnbuckle1");
            AddTurnbuckleOutputs(additionalOutputs, "Turnbuckle2");
            // Add Outputs for the Spring Casing
            additionalOutputs.Add(new OutputDefinition("Casing1", "Casing1"));
            additionalOutputs.Add(new OutputDefinition("Casing2", "Casing2"));
 
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
                Double length1 = m_dLength1.Value;
                Double rodTakeOut = m_dRodTakeOut.Value;
                Double pipeOD = m_dPipeOD.Value;
                Double shoeHeight = m_dShoeHeight.Value;
                Double bBGap = m_dBBGap.Value;
                Double CC = m_dCC.Value;
                Double ccMin = m_dCCMin.Value;
                Double ccMax = m_dCCMax.Value;
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
                string column = m_oColumnShape.Value;
                string columnEnd = m_oColumnEndShape.Value;
                string turnbuckle = m_oTurnbuckleShape.Value;

                // Load The Shape Data Structures
                NutInputs nutInputs = LoadNutDataByQuery(nut, 1);
                PlateInputs plate1Inputs = LoadPlateDataByQuery(plate1);
                PlateInputs plate2Inputs = LoadPlateDataByQuery(plate2);
                Rod1Inputs columnInputs = LoadRodDataByQuery(column);
                NutInputs columnEndInputs = LoadNutDataByQuery(columnEnd, 1);
                TurnbuckleInputs turnbuckleInputs = LoadTurnbuckleDataByQuery(turnbuckle);
                NutInputs casingInputs = LoadNutDataByQuery(casing, 1);

                // Override Shape Inputs
                Double turnbuckleLength = turnbuckleInputs.Opening1 + turnbuckleInputs.Nut.ShapeLength * 2;
                Double turnbuckleTakeOut = turnbuckleLength - 2 * offset1;

                casingInputs.ShapeLength = height1 - plate1Inputs.thickness1 - plate2Inputs.thickness1;
                if (turnbuckle != "" && turnbuckle != "No Value")
                    columnInputs.length = length1 - turnbuckleLength + offset1;
                else
                    columnInputs.length = length1 - columnEndInputs.ShapeLength;

                // Calculate the Output Engineering Attributes
                Double installedLoad = 0, variability = 0, operatingTravel = 0, installedTravel = 0, installedRodTakeOut = 0;

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
                Port port3 = new Port(OccurrenceConnection, part, "RodEnd1", new Position(0, CC / 2, installedRodTakeOut), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["RodEnd1"] = port3;
                Port port4 = new Port(OccurrenceConnection, part, "RodEnd2", new Position(0, -CC / 2, installedRodTakeOut), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["RodEnd2"] = port4;
                Port port5 = new Port(OccurrenceConnection, part, "Surface1", new Position(0, CC / 2, installedRodTakeOut + offset1), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Surface1"] = port5;
                Port port6 = new Port(OccurrenceConnection, part, "Surface2", new Position(0, -CC / 2, installedRodTakeOut + offset1), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Surface2"] = port6;

                // Add the Graphics
                Matrix4X4 matrix = new Matrix4X4();
                if (nut != "" && nut != "No Value")
                {
                    matrix.Translate(new Vector(0, CC / 2, installedRodTakeOut + offset1));
                    AddNut(nutInputs, matrix, m_Symbolic.Outputs, "Nut1");
                    matrix = new Matrix4X4();
                    matrix.Translate(new Vector(0, -CC / 2, installedRodTakeOut + offset1));
                    AddNut(nutInputs, matrix, m_Symbolic.Outputs, "Nut2");
                }

                if (plate1 != "" && plate1 != "No Value")
                {
                    matrix = new Matrix4X4();
                    matrix.Translate(new Vector(-plate1Inputs.width1 / 2, -plate1Inputs.length1 / 2 + CC / 2, offset2 - plate1Inputs.thickness1));
                    AddPlate(plate1Inputs, matrix, m_Symbolic.Outputs, "Plate1");
                    matrix = new Matrix4X4();
                    matrix.Translate(new Vector(-plate1Inputs.width1 / 2, -plate1Inputs.length1 / 2 - CC / 2, offset2 - plate1Inputs.thickness1));
                    AddPlate(plate1Inputs, matrix, m_Symbolic.Outputs, "Plate3");
                }
                if (plate2 != "" && plate2 != "No Value")
                {
                    matrix = new Matrix4X4();
                    matrix.Translate(new Vector(-plate2Inputs.width1 / 2, -plate2Inputs.length1 / 2 + CC / 2, offset2 - height1));
                    AddPlate(plate2Inputs, matrix, m_Symbolic.Outputs, "Plate2");
                    matrix = new Matrix4X4();
                    matrix.Translate(new Vector(-plate2Inputs.width1 / 2, -plate2Inputs.length1 / 2 - CC / 2, offset2 - height1));
                    AddPlate(plate2Inputs, matrix, m_Symbolic.Outputs, "Plate4");
                }

                if (casing != "" && casing != "No Value")
                {
                    if (casingInputs.ShapeType == 3)
                    {
                        matrix = new Matrix4X4();
                        matrix.Rotate(Math.PI / 6, new Vector(0, 0, 1));
                        matrix.Translate(new Vector(0, CC / 2, offset2 - height1 + plate2Inputs.thickness1));
                        AddNut(casingInputs, matrix, m_Symbolic.Outputs, "Casing1");
                        matrix = new Matrix4X4();
                        matrix.Rotate(Math.PI / 6, new Vector(0, 0, 1));
                        matrix.Translate(new Vector(0, -CC / 2, offset2 - height1 + plate2Inputs.thickness1));
                        AddNut(casingInputs, matrix, m_Symbolic.Outputs, "Casing2");
                    }
                    else
                    {
                        matrix = new Matrix4X4();
                        matrix.Translate(new Vector(0, CC / 2, offset2 - height1 + plate2Inputs.thickness1));
                        AddNut(casingInputs, matrix, m_Symbolic.Outputs, "Casing1");
                        matrix = new Matrix4X4();
                        matrix.Translate(new Vector(0, -CC / 2, offset2 - height1 + plate2Inputs.thickness1));
                        AddNut(casingInputs, matrix, m_Symbolic.Outputs, "Casing2");
                    }
                }

                if (columnEnd != "" && columnEnd != "No Value")
                {
                    matrix = new Matrix4X4();
                    matrix.Translate(new Vector(0, CC / 2, installedRodTakeOut + offset1 - columnEndInputs.ShapeLength));
                    AddNut(columnEndInputs, matrix, m_Symbolic.Outputs, "ColumnEnd1");
                    matrix = new Matrix4X4();
                    matrix.Translate(new Vector(0, -CC / 2, installedRodTakeOut + offset1 - columnEndInputs.ShapeLength));
                    AddNut(columnEndInputs, matrix, m_Symbolic.Outputs, "ColumnEnd2");
                }
                if (column != "" && column != "No Value")
                {
                    if (turnbuckle != "" && turnbuckle != "No Value")
                    {
                        matrix = new Matrix4X4();
                        matrix.Rotate(Math.PI, new Vector(1, 0, 0));
                        matrix.Translate(new Vector(0,CC / 2, installedRodTakeOut -turnbuckleTakeOut));
                        AddRod(columnInputs, matrix, m_Symbolic.Outputs, "Column1");
                        matrix = new Matrix4X4();
                        matrix.Rotate(Math.PI, new Vector(1, 0, 0));
                        matrix.Translate(new Vector(0, -CC / 2, installedRodTakeOut - turnbuckleTakeOut));
                        AddRod(columnInputs, matrix, m_Symbolic.Outputs, "Column2");
                    }
                    else
                    {
                        matrix = new Matrix4X4();
                        matrix.Rotate(Math.PI, new Vector(1, 0, 0));
                        matrix.Translate(new Vector(0, CC / 2, installedRodTakeOut + offset1 - columnEndInputs.ShapeLength));
                        AddRod(columnInputs, matrix, m_Symbolic.Outputs, "Column1");
                        matrix = new Matrix4X4();
                        matrix.Rotate(Math.PI, new Vector(1, 0, 0));
                        matrix.Translate(new Vector(0, -CC / 2, installedRodTakeOut + offset1 - columnEndInputs.ShapeLength));
                        AddRod(columnInputs, matrix, m_Symbolic.Outputs, "Column2");
                    }
                }
                if (turnbuckle != "" && turnbuckle != "No Value")
                {
                    matrix = new Matrix4X4();
                    matrix.Translate(new Vector(0, CC / 2,installedRodTakeOut));
                    AddTurnbuckle(turnbuckleInputs, turnbuckleTakeOut, matrix, m_Symbolic.Outputs, "Turnbuckle1");
                    matrix = new Matrix4X4();
                    matrix.Translate(new Vector(0, -CC / 2, installedRodTakeOut));
                    AddTurnbuckle(turnbuckleInputs, turnbuckleTakeOut, matrix, m_Symbolic.Outputs, "Turnbuckle2");
                }

                // Add the Steel Cross Section
                // SweepOptions as 7 is given to create 10 surfaces.
                SweepOptions sweepOptions = (SweepOptions)7;
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

                Double width = crossSection1.Width;

                Double lenAdj = 0;
                if (offset2 < 0)
                    lenAdj = casingInputs.ShapeWidth1 / 2;
                else
                {
                    switch(casingInputs.ShapeType)
                    {
                        case 1:  // Round
                            {
                                if (bBGap >= 0)
                                {
                                    if ((width + bBGap / 2) > casingInputs.ShapeWidth1 / 2)
                                        lenAdj = 0;
                                    else
                                        lenAdj = -Math.Sqrt(((casingInputs.ShapeWidth1 / 2) * (casingInputs.ShapeWidth1 / 2)) - ((width + bBGap / 2) * (width + bBGap / 2)));
                                }
                                else
                                {
                                    if (width / 2 > casingInputs.ShapeWidth1 / 2)
                                        lenAdj = 0;
                                    else
                                        lenAdj = -Math.Sqrt(((casingInputs.ShapeWidth1 / 2) * (casingInputs.ShapeWidth1 / 2)) - ((width / 2) * (width / 2)));
                                }
                            }
                            break;
                        case 2:  // Square
                            lenAdj = - casingInputs.ShapeWidth1 / 2;
                            break;
                        case 3:  // Hex
                            {
                                if (bBGap >= 0)
                                    lenAdj = -((casingInputs.ShapeWidth1 / 2) - (width + bBGap / 2) * Math.Tan(30 * Math.PI / 180));
                                else
                                    lenAdj = -((casingInputs.ShapeWidth1 / 2) - (width / 2) * Math.Tan(30 * Math.PI / 180));
                            }
                            break;
                    }
                }
                Line3d projection;

                // Calculate the Cut Angles
                Double[] startPosNorm1 = new Double[6]; Double[] endPosNorm1 = new Double[6]; Double[] startPosNorm2 = new Double[6]; Double[] endPosNorm2 = new Double[6];


                if (offset2 >= 0)
                {
                    if (bBGap >= 0)
                    {
                        switch (casingInputs.ShapeType)
                        {
                            case 1: // Round
                                {
                                    Double alpha = 0;
                                    if (width + bBGap / 2 > casingInputs.ShapeWidth1 / 2)
                                        alpha = 0;
                                    else
                                        alpha = ((Math.Atan((Math.Sqrt(((casingInputs.ShapeWidth1 / 2) * (casingInputs.ShapeWidth1 / 2)) - ((bBGap / 2) * (bBGap / 2))) - Math.Sqrt(((casingInputs.ShapeWidth1 / 2) * (casingInputs.ShapeWidth1 / 2)) - ((width + bBGap / 2) * (width + bBGap / 2)))) / width)) * 180 / Math.PI);

                                    startPosNorm1[0] = width + bBGap / 2;
                                    startPosNorm1[1] = -CC / 2 - lenAdj;
                                    startPosNorm1[2] = 0;
                                    startPosNorm1[3] = -Math.Sin(alpha * Math.PI / 180);
                                    startPosNorm1[4] = -Math.Cos(alpha * Math.PI / 180);
                                    startPosNorm1[5] = 0;

                                    endPosNorm1[0] = width + bBGap / 2;
                                    endPosNorm1[1] = CC / 2 + lenAdj;
                                    endPosNorm1[2] = 0;
                                    endPosNorm1[3] = -Math.Sin(alpha * Math.PI / 180);
                                    endPosNorm1[4] = Math.Cos(alpha * Math.PI / 180);
                                    endPosNorm1[5] = 0;

                                    startPosNorm2[0] = -width - bBGap / 2;
                                    startPosNorm2[1] = -CC / 2 - lenAdj;
                                    startPosNorm2[2] = 0;
                                    startPosNorm2[3] = Math.Sin(alpha * Math.PI / 180);
                                    startPosNorm2[4] = -Math.Cos(alpha * Math.PI / 180);
                                    startPosNorm2[5] = 0;

                                    endPosNorm2[0] = -width - bBGap / 2;
                                    endPosNorm2[1] = CC / 2 + lenAdj;
                                    endPosNorm2[2] = 0;
                                    endPosNorm2[3] = Math.Sin(alpha * Math.PI / 180);
                                    endPosNorm2[4] = Math.Cos(alpha * Math.PI / 180);
                                    endPosNorm2[5] = 0;
                                }
                                break;
                            case 2: // Square
                                {
                                    startPosNorm1[0] = 0;
                                    startPosNorm1[1] = -CC / 2 - lenAdj;
                                    startPosNorm1[2] = 0;
                                    startPosNorm1[3] = 0;
                                    startPosNorm1[4] = 1;
                                    startPosNorm1[5] = 0;

                                    endPosNorm1[0] = 0;
                                    endPosNorm1[1] = CC / 2 + lenAdj;
                                    endPosNorm1[2] = 0;
                                    endPosNorm1[3] = 0;
                                    endPosNorm1[4] = 1;
                                    endPosNorm1[5] = 0;

                                    startPosNorm2[0] = 0;
                                    startPosNorm2[1] = -CC / 2 - lenAdj;
                                    startPosNorm2[2] = 0;
                                    startPosNorm2[3] = 0;
                                    startPosNorm2[4] = 1;
                                    startPosNorm2[5] = 0;

                                    endPosNorm2[0] = 0;
                                    endPosNorm2[1] = CC / 2 + lenAdj;
                                    endPosNorm2[2] = 0;
                                    endPosNorm2[3] = 0;
                                    endPosNorm2[4] = 1;
                                    endPosNorm2[5] = 0;
                                }
                                break;
                            case 3: // Hex
                                {
                                    startPosNorm1[0] = width + bBGap / 2;
                                    startPosNorm1[1] = -CC / 2 - lenAdj;
                                    startPosNorm1[2] = 0;
                                    startPosNorm1[3] = -Math.Sin(30 * Math.PI / 180);
                                    startPosNorm1[4] = -Math.Cos(30 * Math.PI / 180);
                                    startPosNorm1[5] = 0;

                                    endPosNorm1[0] = width + bBGap / 2;
                                    endPosNorm1[1] = CC / 2 + lenAdj;
                                    endPosNorm1[2] = 0;
                                    endPosNorm1[3] = -Math.Sin(30 * Math.PI / 180);
                                    endPosNorm1[4] = Math.Cos(30 * Math.PI / 180);
                                    endPosNorm1[5] = 0;

                                    startPosNorm2[0] = -width - bBGap / 2;
                                    startPosNorm2[1] = -CC / 2 - lenAdj;
                                    startPosNorm2[2] = 0;
                                    startPosNorm2[3] = Math.Sin(30 * Math.PI / 180);
                                    startPosNorm2[4] = -Math.Cos(30 * Math.PI / 180);
                                    startPosNorm2[5] = 0;

                                    endPosNorm2[0] = -width - bBGap / 2;
                                    endPosNorm2[1] = CC / 2 + lenAdj;
                                    endPosNorm2[2] = 0;
                                    endPosNorm2[3] = Math.Sin(30 * Math.PI / 180);
                                    endPosNorm2[4] = Math.Cos(30 * Math.PI / 180);
                                    endPosNorm2[5] = 0;
                                }
                                break;
                        }
                    }
                    else
                    {
                        startPosNorm1[0] = 0;
                        startPosNorm1[1] = -CC / 2 - lenAdj;
                        startPosNorm1[2] = 0;
                        startPosNorm1[3] = 0;
                        startPosNorm1[4] = 1;
                        startPosNorm1[5] = 0;

                        endPosNorm1[0] = 0;
                        endPosNorm1[1] = CC / 2 + lenAdj;
                        endPosNorm1[2] = 0;
                        endPosNorm1[3] = 0;
                        endPosNorm1[4] = 1;
                        endPosNorm1[5] = 0;

                        startPosNorm2[0] = 0;
                        startPosNorm2[1] = -CC / 2 - lenAdj;
                        startPosNorm2[2] = 0;
                        startPosNorm2[3] = 0;
                        startPosNorm2[4] = 1;
                        startPosNorm2[5] = 0;

                        endPosNorm2[0] = 0;
                        endPosNorm2[1] = CC / 2 + lenAdj;
                        endPosNorm2[2] = 0;
                        endPosNorm2[3] = 0;
                        endPosNorm2[4] = 1;
                        endPosNorm2[5] = 0;
                    }
                }
                else
                {
                    startPosNorm1[0] = 0;
                    startPosNorm1[1] = -CC / 2 - lenAdj;
                    startPosNorm1[2] = 0;
                    startPosNorm1[3] = 0;
                    startPosNorm1[4] = 1;
                    startPosNorm1[5] = 0;

                    endPosNorm1[0] = 0;
                    endPosNorm1[1] = CC / 2 + lenAdj;
                    endPosNorm1[2] = 0;
                    endPosNorm1[3] = 0;
                    endPosNorm1[4] = 1;
                    endPosNorm1[5] = 0;

                    startPosNorm2[0] = 0;
                    startPosNorm2[1] = -CC / 2 - lenAdj;
                    startPosNorm2[2] = 0;
                    startPosNorm2[3] = 0;
                    startPosNorm2[4] = 1;
                    startPosNorm2[5] = 0;

                    endPosNorm2[0] = 0;
                    endPosNorm2[1] = CC / 2 + lenAdj;
                    endPosNorm2[2] = 0;
                    endPosNorm2[3] = 0;
                    endPosNorm2[4] = 1;
                    endPosNorm2[5] = 0;
                }

                string section1 = "Section1"; string section2 = "Section2";
                Collection<ISurface> crossSecSurfaces1, crossSecSurfaces2;
                // here validating CC / 2 to only positive because it's creating wrong graphics, if it is negative or zero.
                if (CC / 2 > 0)
                {
                    if (bBGap >= 0)
                    {
                        // here bBGap / 2 is taken as X-coordinate to create the CrossSections with some gap similar to VB Graphic.
                        projection = new Line3d(new Position(bBGap / 2, -CC / 2 - lenAdj, 0), new Position(bBGap / 2, CC / 2 + lenAdj, 0));
                        // First Section
                        crossSecSurfaces1 = crossSectionServices.GetProjectionSurfacesFromCrossSection(crossSection1, projection, 7, false, 0, new Position(startPosNorm1[0], startPosNorm1[1], startPosNorm1[2]), new Vector(startPosNorm1[3], startPosNorm1[4], startPosNorm1[5]), new Position(endPosNorm1[0], endPosNorm1[1], endPosNorm1[2]), new Vector(endPosNorm1[3], endPosNorm1[4], endPosNorm1[5]), sweepOptions);
                        for (int i = 1; i <= crossSecSurfaces1.Count; i++)
                        {
                            m_Symbolic.Outputs[section1] = crossSecSurfaces1[i - 1];
                            section1 = "Section1" + i;
                        }
                        // Second Section
                        crossSection2 = crossSection1;
                        projection = new Line3d(new Position(-bBGap / 2, -CC / 2 - lenAdj, 0), new Position(-bBGap / 2, CC / 2 + lenAdj, 0));
                        crossSecSurfaces2 = crossSectionServices.GetProjectionSurfacesFromCrossSection(crossSection2, projection, 7, true, 0, new Position(startPosNorm2[0], startPosNorm2[1], startPosNorm2[2]), new Vector(startPosNorm2[3], startPosNorm2[4], startPosNorm2[5]), new Position(endPosNorm2[0], endPosNorm2[1], endPosNorm2[2]), new Vector(endPosNorm2[3], endPosNorm2[4], endPosNorm2[5]), sweepOptions);
                        for (int i = 1; i <= crossSecSurfaces2.Count; i++)
                        {
                            m_Symbolic.Outputs[section2] = crossSecSurfaces2[i - 1];
                            section2 = "Section2" + i;
                        }
                    }
                    else
                    {
                        projection = new Line3d(new Position(0, -CC / 2 - lenAdj, 0), new Position(0, CC / 2 + lenAdj, 0));
                        // First Section
                        crossSecSurfaces1 = crossSectionServices.GetProjectionSurfacesFromCrossSection(crossSection1, projection, 8, false, 0, new Position(startPosNorm1[0], startPosNorm1[1], startPosNorm1[2]), new Vector(startPosNorm1[3], startPosNorm1[4], startPosNorm1[5]), new Position(endPosNorm1[0], endPosNorm1[1], endPosNorm1[2]), new Vector(endPosNorm1[3], endPosNorm1[4], endPosNorm1[5]), sweepOptions);
                        for (int i = 1; i <= crossSecSurfaces1.Count; i++)
                        {
                            m_Symbolic.Outputs[section1] = crossSecSurfaces1[i - 1];
                            section1 = "Section1" + i;
                        }
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
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of VariableTypeG.cs."));
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
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrWeightCG, "Error in Weight CG of VariableTypeG.cs."));
                }
            }
        }

        #endregion

    }

}

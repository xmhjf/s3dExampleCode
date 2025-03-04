//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   PipeClamp.cs
//    HSSmartPart,Ingr.SP3D.Content.Support.Symbols.PipeClamp
//   Author       :  Rajeswari
//   Creation Date:  25-Jan-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   25-Jan-2013   Rajeswari  CR-CP-222485 Initial Creation 
//   25/Mar/2013   Rajeswari  DI-CP-228142  Modify the error handling for delivered H&S symbols
//   30/10/2013     Hema      CR-CP-242533  Provide the ability to store GType outputs as a single blob.
//   11/08/2014    Ramya      TR-CP-256377  Additional input values are retrieved from catalog part in smart parts 
//   17-Oct-2014    Chethan   CR-CP-253371  Add Maintenance Aspect to PipeClamp smartpart and Strap smartpart
//   12-12-2014     PVK       TR-CP-264951	Resolve P3 coverity issues found in November 2014 report  
//   10-06-2015     PVK	      TR-CP-274155	SmartPart TDL Errors should be corrected. 
//   26-10-2015     PVK       Resolve coverity issues found in Octpber 2015 report
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using System.Text;
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
    [CacheOption(CacheOptionType.Cached)]
    public class PipeClamp : SmartPartComponentDefinition, ICustomWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "HSSmartPart,Ingr.SP3D.Content.Support.Symbols.PipeClamp"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"
        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        #endregion

        #region "Definitions of Aspects and their Outputs"
        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Route", "Route")]
        [SymbolOutput("Side", "Side")]
        [SymbolOutput("Wing", "Wing")]
        public AspectDefinition m_Symbolic;

        [Aspect("Maintenance", "Maintenance Aspect", AspectID.Maintenance)]
        public AspectDefinition m_Maintenance;

        #endregion


        #region "Definition of Additional Inputs"
        public override IEnumerable<Input> AdditionalInputs
        {
            get
            {
                int endIndex;
                List<Input> additionalInputs = new List<Input>();
                AddPipeClampInputs(2, out endIndex, additionalInputs);
                additionalInputs.Add((Input)new InputDouble(endIndex, "InsulationTh", "InsulationTh", 0, true));
				additionalInputs.Add(new InputDouble(++endIndex, "MaintenanceThickness", "MaintenanceThickness", 0, true));
                additionalInputs.Add(new InputDouble(++endIndex, "CreateMaintenanceAspect", "CreateMaintenanceAspect", 0, true));
				return additionalInputs;
            }
        }
        #endregion

        #region "Definition of Additional Outputs"
        public override IEnumerable<OutputDefinition> AdditionalOutputs(string aspectName)
        {
            List<OutputDefinition> additionalOutputs = new List<OutputDefinition>();
            if (aspectName == "Symbolic")
            {
                AddPipeClampOutputs(aspectName, additionalOutputs, "clamp1");
            }
            if (aspectName == "Maintenance")
            {
                AddPipeClampOutputs(aspectName, additionalOutputs, "clamp2");
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

                int endIndex;
                PipeClampInputs pipeClamp = LoadPipeClampData(2, out endIndex);
                
                if (base.ToDoListMessage != null)
                    if (base.ToDoListMessage.Type == ToDoMessageTypes.ToDoMessageError)
                        return;

                Double insulationThickness = 0;
                insulationThickness = GetDoubleInputValue(endIndex++);

                ///Maintenance inputs
                Double maintenanceThickness = 0;
                bool createMaintenanceAspect = false; ;
                maintenanceThickness = GetDoubleInputValue(endIndex++);
                createMaintenanceAspect= Convert.ToBoolean(GetDoubleInputValue(endIndex++));

                if (insulationThickness > 0)
                {
                    pipeClamp.Diameter1 = insulationThickness * 2 + pipeClamp.Diameter1;
                    if (pipeClamp.Diameter4 > 0)
                        pipeClamp.Diameter4 = insulationThickness * 2 + pipeClamp.Diameter1;
                    pipeClamp.Height1 = pipeClamp.Height1 + insulationThickness;
                    pipeClamp.Height2 = pipeClamp.Height2 + insulationThickness;

                    pipeClamp.BoltRow1.Offset = pipeClamp.BoltRow1.Offset + insulationThickness;
                    pipeClamp.BoltRow2.Offset = pipeClamp.BoltRow2.Offset + insulationThickness;
                    pipeClamp.BoltRow3.Offset = pipeClamp.BoltRow3.Offset + insulationThickness;
                    pipeClamp.BoltRow4.Offset = pipeClamp.BoltRow4.Offset + insulationThickness;
                    pipeClamp.BoltRow5.Offset = pipeClamp.BoltRow5.Offset + insulationThickness;
                    pipeClamp.BoltRow6.Offset = pipeClamp.BoltRow6.Offset + insulationThickness;
                }

                Double angle3OffsetCalcZ = 0, angle3OffsetCalcY = 0, offsetCalcy = 0, offsetCalcZ = 0, angle2OffsetCalcX = 0, angle2OffsetCalcY = 0;

                // Add the Graphic
                AddPipeClamp(pipeClamp, new Matrix4X4(), m_Symbolic.Outputs, "clamp1");

                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                //ports
                int clampconfig = pipeClamp.ClampCfg;
                switch (clampconfig)
                {
                    case 1:
                        // Rod Hanger
                        Port route = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, Math.Sin(pipeClamp.Angle3), Math.Cos(pipeClamp.Angle3)));
                        m_Symbolic.Outputs["Route"] = route;
                        Port side = new Port(OccurrenceConnection, part, "Side", new Position(0, pipeClamp.Diameter1 / 2 + pipeClamp.Thickness1 + pipeClamp.Height3, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                        m_Symbolic.Outputs["Side"] = side;

                        if ((int)pipeClamp.BoltRow1.MultiLocateBy == 1)
                        {
                            if (pipeClamp.BoltRow1.MultiQty < 2)
                            {
                                Port wing = new Port(OccurrenceConnection, part, "Wing", new Position(-pipeClamp.BoltRow1.MultiLocation, 0, pipeClamp.BoltRow1.Offset), new Vector(1, 0, 0), new Vector(0, 0, 1));
                                m_Symbolic.Outputs["Wing"] = wing;
                            }
                            else if ((int)pipeClamp.BoltRow1.MultiQty == 2)
                            {
                                Port wing = new Port(OccurrenceConnection, part, "Wing", new Position(-pipeClamp.BoltRow1.MultiLocation / 2, 0, pipeClamp.BoltRow1.Offset), new Vector(1, 0, 0), new Vector(0, 0, 1));
                                m_Symbolic.Outputs["Wing"] = wing;
                            }
                            else
                            {
                                Port wing = new Port(OccurrenceConnection, part, "Wing", new Position(-pipeClamp.BoltRow1.MultiLocation, 0, pipeClamp.BoltRow1.Offset), new Vector(1, 0, 0), new Vector(0, 0, 1));
                                m_Symbolic.Outputs["Wing"] = wing;
                            }
                        }
                        else
                        {
                            if ((int)pipeClamp.BoltRow1.MultiLocation != 0)
                            {
                                Port wing = new Port(OccurrenceConnection, part, "Wing", new Position(-pipeClamp.Width1 / 2 + pipeClamp.BoltRow1.MultiLocation + pipeClamp.Pin1Diameter / 2, 0, pipeClamp.BoltRow1.Offset), new Vector(1, 0, 0), new Vector(0, 0, 1));
                                m_Symbolic.Outputs["Wing"] = wing;
                            }
                            else
                            {
                                Port wing = new Port(OccurrenceConnection, part, "Wing", new Position(-pipeClamp.Width1 / 2 + pipeClamp.BoltRow1.MultiLocation, 0, pipeClamp.BoltRow1.Offset), new Vector(1, 0, 0), new Vector(0, 0, 1));
                                m_Symbolic.Outputs["Wing"] = wing;
                            }
                        }
                        break;
                    case 2:
                        // Side Mounted
                        if (HgrCompareDoubleService.cmpdbl(pipeClamp.Angle3, 0)==false)
                        {
                            angle3OffsetCalcY = -Math.Cos(pipeClamp.Angle3);
                            angle3OffsetCalcZ = -Math.Sin(pipeClamp.Angle3);
                            if (HgrCompareDoubleService.cmpdbl(pipeClamp.BoltRow1.Offset, 0)==false)
                            {
                                offsetCalcy = Math.Cos(pipeClamp.Angle3) * (pipeClamp.BoltRow1.Offset);
                                offsetCalcZ = Math.Sin(pipeClamp.Angle3) * (pipeClamp.BoltRow1.Offset);
                            }
                            else
                            {
                                offsetCalcy = Math.Cos(pipeClamp.Angle3) * (pipeClamp.Diameter1 / 2 + pipeClamp.Thickness1);
                                offsetCalcZ = Math.Sin(pipeClamp.Angle3) * (pipeClamp.Diameter1 / 2 + pipeClamp.Thickness1);
                            }
                        }
                        else
                        {
                            angle3OffsetCalcY = 0;
                            angle3OffsetCalcZ = 1;
                            if (HgrCompareDoubleService.cmpdbl(pipeClamp.BoltRow1.Offset, 0) == false)
                                offsetCalcy = pipeClamp.BoltRow1.Offset;
                            else
                                offsetCalcy = pipeClamp.Diameter1 / 2 + pipeClamp.Thickness1;
                            offsetCalcZ = 0;
                        }

                        Port route1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                        m_Symbolic.Outputs["Route"] = route1;
                        if (HgrCompareDoubleService.cmpdbl(pipeClamp.Angle2, 0) == false)
                        {
                            angle2OffsetCalcX = Math.Cos(pipeClamp.Angle2) * (Math.PI / 180);
                            angle2OffsetCalcY = Math.Sin(pipeClamp.Angle2) * (Math.PI / 180);

                            Port side1 = new Port(OccurrenceConnection, part, "Side", new Position(0, 0, pipeClamp.Diameter1 / 2 + pipeClamp.Thickness1 + pipeClamp.Height3), new Vector(angle2OffsetCalcX, angle2OffsetCalcY, 0), new Vector(0, 0, 1));
                            m_Symbolic.Outputs["Side"] = side1;
                        }
                        else
                        {
                            if (HgrCompareDoubleService.cmpdbl(pipeClamp.Width3, 0) == false && (pipeClamp.Width3 / 2) < (Math.Sin(45) * (Math.PI / 180)) * ((pipeClamp.Diameter1 / 2) + (pipeClamp.Thickness1)))
                            {
                                Double portPositionZAngle = Math.Asin(((pipeClamp.Width3 / 2) / ((pipeClamp.Diameter1 / 2) + pipeClamp.Thickness1))) * (180 / Math.PI);
                                Double portPositionZ = (Math.Cos(portPositionZAngle) * (Math.PI / 180)) * ((pipeClamp.Diameter1 / 2) + pipeClamp.Thickness1);

                                Port side2 = new Port(OccurrenceConnection, part, "Side", new Position(0, 0, portPositionZ), new Vector(1, 0, 0), new Vector(0, 0, 1));
                                m_Symbolic.Outputs["Side"] = side2;
                            }
                            else
                            {
                                Port side2 = new Port(OccurrenceConnection, part, "Side", new Position(0, 0, pipeClamp.Diameter1 / 2 + pipeClamp.Thickness1 + pipeClamp.Height3), new Vector(1, 0, 0), new Vector(0, 0, 1));
                                m_Symbolic.Outputs["Side"] = side2;
                            }
                        }
                        Port wing1 = new Port(OccurrenceConnection, part, "Wing", new Position(0, -offsetCalcy, -offsetCalcZ), new Vector(1, 0, 0), new Vector(0, angle3OffsetCalcY, angle3OffsetCalcZ));
                        m_Symbolic.Outputs["Wing"] = wing1;
                        break;
                    case 3:
                        // Side Mounted Two Port
                        Double positionY, positionZ;
                        if (HgrCompareDoubleService.cmpdbl(pipeClamp.Angle2, 0) == false && pipeClamp.Width3 > -0.9999 && pipeClamp.Width3 < 0.0001)
                        {
                            positionY = Math.Sin((pipeClamp.Angle2 / (180 * Math.PI)) / 2) * ((pipeClamp.Diameter1 / 2) + pipeClamp.Thickness1);
                            positionZ = Math.Cos((pipeClamp.Angle2 / (180 * Math.PI)) / 2) * ((pipeClamp.Diameter1 / 2) + pipeClamp.Thickness1);
                        }
                        else
                        {
                            if (pipeClamp.Width3 > (pipeClamp.Diameter1 + (pipeClamp.Thickness1 * 2)))
                            {
                                //error msg
                                if (pipeClamp.Width3 > (pipeClamp.Diameter1 + (pipeClamp.Thickness1 * 2)))
                                    RaiseSmartPartTDLError(SmartPartSymbolResourceIDs.ErrInvalidWidth3NGTDiameter, "Width3 can not be greater than the outside diameter of the pipe");

                                positionY = -pipeClamp.Diameter1 / 2 - pipeClamp.Thickness1;
                                positionZ = 0;
                                pipeClamp.Angle2 = 0;
                            }
                            else
                            {
                                positionY = -pipeClamp.Width3 / 2;
                                positionZ = Math.Sqrt((((pipeClamp.Diameter1 / 2) + pipeClamp.Thickness1) * ((pipeClamp.Diameter1 / 2) + pipeClamp.Thickness1)) - ((pipeClamp.Width3 / 2) * (pipeClamp.Width3 / 2)));
                            }
                        }

                        Double angle3OffsetCalcYPort1 = 0, angle3OffsetCalcZPort1 = 0, angle3OffsetCalcYPort2 = 0, angle3OffsetCalcZPort2 = 0;
                        if (HgrCompareDoubleService.cmpdbl(pipeClamp.Angle2, 0) == false)
                        {
                            angle3OffsetCalcYPort1 = -Math.Cos(pipeClamp.Angle2);
                            angle3OffsetCalcZPort1 = -Math.Sin(pipeClamp.Angle2);
                            angle3OffsetCalcYPort2 = -Math.Cos(pipeClamp.Angle2);
                            angle3OffsetCalcZPort2 = -Math.Sin(pipeClamp.Angle2);
                        }

                        Port route2 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(-1, 0, 0), new Vector(0, 0, -1));
                        m_Symbolic.Outputs["Route"] = route2;

                        if (HgrCompareDoubleService.cmpdbl(pipeClamp.Angle2, 0) == false)
                        {
                            Port side3 = new Port(OccurrenceConnection, part, "Side", new Position(0, positionY, -positionZ), new Vector(1, 0, 0), new Vector(0, angle3OffsetCalcZPort1, angle3OffsetCalcYPort1));
                            m_Symbolic.Outputs["Side"] = side3;
                            Port wing2 = new Port(OccurrenceConnection, part, "Wing", new Position(0, -positionY, -positionZ), new Vector(1, 0, 0), new Vector(0, -angle3OffsetCalcZPort1, angle3OffsetCalcYPort1));
                            m_Symbolic.Outputs["Wing"] = wing2;
                        }
                        else
                        {
                            Port side3 = new Port(OccurrenceConnection, part, "Side", new Position(0, positionY, -positionZ), new Vector(1, 0, 0), new Vector(0, 0, -1));
                            m_Symbolic.Outputs["Side"] = side3;
                            Port wing2 = new Port(OccurrenceConnection, part, "Wing", new Position(0, -positionY, -positionZ), new Vector(-1, 0, 0), new Vector(0, 0, -1));
                            m_Symbolic.Outputs["Wing"] = wing2;
                        }
                        break;
                    case 4:
                        // Direct Bolted
                        Port route3 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                        m_Symbolic.Outputs["Route"] = route3;
                        Port side4 = new Port(OccurrenceConnection, part, "Side", new Position(0, pipeClamp.Width3 / 2, --pipeClamp.Height3), new Vector(1, 0, 0), new Vector(0, 0, 1));
                        m_Symbolic.Outputs["Side"] = side4;
                        Port wing3 = new Port(OccurrenceConnection, part, "Wing", new Position(0, -pipeClamp.Width3 / 2, -pipeClamp.Height3), new Vector(-1, 0, 0), new Vector(0, 0, 1));
                        m_Symbolic.Outputs["Wing"] = wing3;
                        break;
                    case 5:
                        // Riser Rod Hanger
                        Port route4 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(0, 0, 1), new Vector(-Math.Sin(pipeClamp.Angle3), -Math.Cos(pipeClamp.Angle3), 0));
                        m_Symbolic.Outputs["Route"] = route4;

                        // Determine the position of the load pin port on Bolt Row 2 Based upon the Mutli Fucntioniality
                        double portZOffset = 0;

                        if ((int)pipeClamp.BoltRow2.MultiLocateBy == 1 || pipeClamp.BoltRow2.MultiQty > 3)
                        {
                            if ((pipeClamp.BoltRow2.MultiQty % 2) > 0)
                            {
                                if ((int)pipeClamp.BoltRow2.MultiQty == 1)
                                {
                                    portZOffset = pipeClamp.BoltRow2.MultiLocation;
                                }
                                else
                                {
                                    portZOffset = pipeClamp.BoltRow2.MultiLocation * ((pipeClamp.BoltRow2.MultiQty - 1) / 2);
                                }
                            }
                            else
                            {
                                portZOffset = pipeClamp.BoltRow2.MultiLocation / 2 + (pipeClamp.BoltRow2.MultiLocation * (pipeClamp.BoltRow2.MultiQty / 2 - 1));
                            }
                        }
                        else
                        {
                            portZOffset = pipeClamp.Width1 / 2 - pipeClamp.BoltRow2.MultiLocation - pipeClamp.Pin1Diameter / 2;
                        }

                        Port side5 = new Port(OccurrenceConnection, part, "Side", new Position(0, pipeClamp.BoltRow2.Offset, portZOffset), new Vector(0, 1, 0), new Vector(0, 0, 1));
                        m_Symbolic.Outputs["Side"] = side5;

                        // Determine the position of the load pin port on Bolt Row 1 Based upon the Mutli Fucntioniality
                        if ((int)pipeClamp.BoltRow1.MultiLocateBy == 1 || pipeClamp.BoltRow1.MultiQty > 3)
                        {
                            if ((pipeClamp.BoltRow1.MultiQty % 2) > 0)
                            {
                                if ((int)pipeClamp.BoltRow1.MultiQty == 1)
                                {
                                    portZOffset = pipeClamp.BoltRow1.MultiLocation;
                                }
                                else
                                {
                                    portZOffset = pipeClamp.BoltRow1.MultiLocation * ((pipeClamp.BoltRow1.MultiQty - 1) / 2);
                                }
                            }
                            else
                            {
                                portZOffset = pipeClamp.BoltRow1.MultiLocation / 2 + (pipeClamp.BoltRow1.MultiLocation * (pipeClamp.BoltRow1.MultiQty / 2 - 1));
                            }
                        }
                        else
                        {
                            portZOffset = pipeClamp.Width1 / 2 - pipeClamp.BoltRow1.MultiLocation - pipeClamp.Pin1Diameter / 2;
                        }

                        Port wing4 = new Port(OccurrenceConnection, part, "Wing", new Position(0, -pipeClamp.BoltRow1.Offset, portZOffset), new Vector(0, -1, 0), new Vector(0, 0, 1));
                        m_Symbolic.Outputs["Wing"] = wing4;
                        break;
                    case 6:
                        // Riser on Structure
                        Port route5 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(0, 0, 1), new Vector(0, -1, 0));
                        m_Symbolic.Outputs["Route"] = route5;
                        if (pipeClamp.Width1 > pipeClamp.Width2)
                        {
                            Port side6 = new Port(OccurrenceConnection, part, "Side", new Position(0, pipeClamp.Width3 / 2, -pipeClamp.Width1 / 2), new Vector(-1, 0, 0), new Vector(0, 0, 1));
                            m_Symbolic.Outputs["Side"] = side6;
                            Port wing5 = new Port(OccurrenceConnection, part, "Wing", new Position(0, -pipeClamp.Width3 / 2, -pipeClamp.Width1 / 2), new Vector(1, 0, 0), new Vector(0, 0, 1));
                            m_Symbolic.Outputs["Wing"] = wing5;
                        }
                        else
                        {
                            Port side6 = new Port(OccurrenceConnection, part, "Side", new Position(0, pipeClamp.Width3 / 2, -pipeClamp.Width2 / 2), new Vector(-1, 0, 0), new Vector(0, 0, 1));
                            m_Symbolic.Outputs["Side"] = side6;
                            Port wing5 = new Port(OccurrenceConnection, part, "Wing", new Position(0, -pipeClamp.Width3 / 2, -pipeClamp.Width2 / 2), new Vector(1, 0, 0), new Vector(0, 0, 1));
                            m_Symbolic.Outputs["Wing"] = wing5;
                        }
                        break;
                    case 7:
                        //  Offset Pipe Clamp
                        Port route6 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(-1, 0, 0), new Vector(0, 0, -1));
                        m_Symbolic.Outputs["Route"] = route6;
                        Port side7 = new Port(OccurrenceConnection, part, "Side", new Position(0, 0, -pipeClamp.Length), new Vector(0, 1, 0), new Vector(0, 0, 1));
                        m_Symbolic.Outputs["Side"] = side7;
                        break;
                }

                // Construction of maintenance Aspect

                if (createMaintenanceAspect)
                {
                    if (pipeClamp.Dim2 > 0)
                    {
                        pipeClamp.Dim2 = pipeClamp.Dim2 + maintenanceThickness;
                    }
                    else
                    {
                        pipeClamp.Thickness1 = pipeClamp.Thickness1 + maintenanceThickness;
                    }
                    if (pipeClamp.Height1 > 0)
                    {
                        pipeClamp.Height1 = pipeClamp.Height1 + maintenanceThickness;
                    }
                    if (pipeClamp.Height2 > 0)
                    {
                        pipeClamp.Height2 = pipeClamp.Height2 + maintenanceThickness;
                    }
                    if (pipeClamp.Dim1 > 0)
                    {
                        pipeClamp.Dim1 = pipeClamp.Dim1 + maintenanceThickness;
                    }
                    if (pipeClamp.Thickness2 > 0)
                    {
                        pipeClamp.Thickness2 = pipeClamp.Thickness2 + maintenanceThickness;
                    }

                    // Add the Graphic
                    AddPipeClamp(pipeClamp, new Matrix4X4(), m_Maintenance.Outputs, "clamp2");


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
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of PipeClamp"));
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
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrWeightCG, "Error in Weight CG of PipeClamp"));
                }
            }
        }

        #endregion
    }
}

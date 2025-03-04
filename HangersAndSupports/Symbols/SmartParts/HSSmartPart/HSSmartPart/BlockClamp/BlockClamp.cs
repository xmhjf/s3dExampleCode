//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   BlockClamp.cs
//    HSSmartPart,Ingr.SP3D.Content.Support.Symbols.BlockClamp
//   Author       : Rajeswari
//   Creation Date:  15/Feb/2013
//   Description: CR-CP-222473 Initial Creation

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   15/Feb/2013  Rajeswari  CR-CP-222473 Initial Creation 
//   25/Mar/2013  Rajeswari  DI-CP-228142  Modify the error handling for delivered H&S symbols
//   12-12-2014      PVK     TR-CP-264951	Resolve P3 coverity issues found in November 2014 report
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
    [CacheOption(CacheOptionType.Cached)]
    [VariableOutputs]
    public class BlockClamp : SmartPartComponentDefinition, ICustomWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "HSSmartPart,Ingr.SP3D.Content.Support.Symbols.BlockClamp"
        //----------------------------------------------------------------------------------


        #region "Definition of Inputs"
        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Route", "Route")]
        [SymbolOutput("Structure", "Structure")]
        [SymbolOutput("Rod", "Rod")]
        [SymbolOutput("Weld", "Weld")]
        public AspectDefinition m_Symbolic;

        #endregion
        #region "Definition of Additional Inputs"
        public override IEnumerable<Input> AdditionalInputs
        {
            get
            {
                int endIndex;
                List<Input> additionalInputs = new List<Input>();
                AddBlockClampInputs(2, out endIndex, additionalInputs);
                return additionalInputs;
            }
        }
        #endregion
        const int iNumofShapes = 2;
        #region "Definition of Additional Outputs"
        public override IEnumerable<OutputDefinition> AdditionalOutputs(string aspectName)
        {
            List<OutputDefinition> additionalOutputs = new List<OutputDefinition>();
            if (aspectName == "Symbolic")
            {
                for (int i = 1; i <= iNumofShapes; i++)
                    AddBlockClampOutputs(aspectName, additionalOutputs);
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

                BlockClampInputs blockClamp = LoadBlockClampData(2);
                if (base.ToDoListMessage != null)
                    if (base.ToDoListMessage.Type == ToDoMessageTypes.ToDoMessageError)
                        return;

                // Middle Plates
                PlateInputs midTopPlate = LoadPlateDataByQuery(blockClamp.MidPlateTopShape);
                PlateInputs midBotPlate = LoadPlateDataByQuery(blockClamp.MidPlateBotShape);

                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                //ports

                // Add Route Port
                if (blockClamp.Diameter1 > 0 && blockClamp.Diameter2 > 0 && blockClamp.Offset1 > 0)
                {
                    Port route = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, -1));
                    m_Symbolic.Outputs["Route"] = route;
                }
                else if (HgrCompareDoubleService.cmpdbl(blockClamp.Offset1, 0) == false && HgrCompareDoubleService.cmpdbl(blockClamp.Diameter2, 0) == true)
                {
                    Port route = new Port(OccurrenceConnection, part, "Route", new Position(0, blockClamp.Offset1, 0), new Vector(1, 0, 0), new Vector(0, 0, -1));
                    m_Symbolic.Outputs["Route"] = route;
                }
                else if (HgrCompareDoubleService.cmpdbl(blockClamp.Offset1 , 0)==true && HgrCompareDoubleService.cmpdbl(blockClamp.Diameter2 , 0)==true && blockClamp.Offset2 > 0)
                {
                    if (HgrCompareDoubleService.cmpdbl(blockClamp.Width1, 0)==true)
                    {
                        Port route = new Port(OccurrenceConnection, part, "Route", new Position(0, blockClamp.Offset2 - midTopPlate.width1 / 2, 0), new Vector(1, 0, 0), new Vector(0, 0, -1));
                        m_Symbolic.Outputs["Route"] = route;
                    }
                    else
                    {
                        Port route = new Port(OccurrenceConnection, part, "Route", new Position(0, blockClamp.Offset2 - blockClamp.Width1 / 2, 0), new Vector(1, 0, 0), new Vector(0, 0, -1));
                        m_Symbolic.Outputs["Route"] = route;
                    }
                }
                else
                {
                    Port route = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, -1));
                    m_Symbolic.Outputs["Route"] = route;
                }

                // Add Structure Port
                if (HgrCompareDoubleService.cmpdbl(blockClamp.Width2, 0)==true && HgrCompareDoubleService.cmpdbl(blockClamp.Length2, 0)==true && HgrCompareDoubleService.cmpdbl(blockClamp.Thickness2, 0)==true && midBotPlate.curvedEndY > 0)
                {
                    Port structure = new Port(OccurrenceConnection, part, "Structure", new Position(0, 0, -midBotPlate.length1 - midBotPlate.curvedEndY - blockClamp.Thickness4 - blockClamp.Thickness3), new Vector(1, 0, 0), new Vector(0, 0, 1));
                    m_Symbolic.Outputs["Structure"] = structure;
                }
                else if (HgrCompareDoubleService.cmpdbl(blockClamp.Width2, 0)==true && HgrCompareDoubleService.cmpdbl(blockClamp.Length2, 0)==true && HgrCompareDoubleService.cmpdbl(blockClamp.Thickness2, 0)==true && midBotPlate.curvedEndY < 0)
                {
                    Port structure = new Port(OccurrenceConnection, part, "Structure", new Position(0, 0, -midBotPlate.length1 - midBotPlate.curvedEndY - blockClamp.Thickness4 - blockClamp.Thickness3), new Vector(1, 0, 0), new Vector(0, 0, 1));
                    m_Symbolic.Outputs["Structure"] = structure;
                }
                else if (HgrCompareDoubleService.cmpdbl(blockClamp.Width2, 0)==true && HgrCompareDoubleService.cmpdbl(blockClamp.Length2, 0)==true && HgrCompareDoubleService.cmpdbl(blockClamp.Thickness2, 0)==true && HgrCompareDoubleService.cmpdbl(midBotPlate.curvedEndY, 0)==true && HgrCompareDoubleService.cmpdbl(midBotPlate.curvedEndRad, 0)==true)
                {
                    Port structure = new Port(OccurrenceConnection, part, "Structure", new Position(0, 0, -midBotPlate.length1 - blockClamp.Diameter1 / 2 - blockClamp.Thickness4 - blockClamp.Thickness3), new Vector(1, 0, 0), new Vector(0, 0, 1));
                    m_Symbolic.Outputs["Structure"] = structure;
                }
                else
                {
                    Port structure = new Port(OccurrenceConnection, part, "Structure", new Position(0, 0, -blockClamp.Height2 - blockClamp.Thickness4 - blockClamp.Thickness3), new Vector(1, 0, 0), new Vector(0, 0, 1));
                    m_Symbolic.Outputs["Structure"] = structure;
                }

                // Add Rod Port
                if (blockClamp.Thickness6 > 0)
                {
                    if (HgrCompareDoubleService.cmpdbl(blockClamp.Width1, 0)==true && HgrCompareDoubleService.cmpdbl(blockClamp.Length1, 0)==true && HgrCompareDoubleService.cmpdbl(blockClamp.Thickness1, 0)==true && midTopPlate.curvedEndY > 0)
                    {
                        Port rod = new Port(OccurrenceConnection, part, "Rod", new Position(0, 0, midTopPlate.length1 + midTopPlate.curvedEndY + blockClamp.Thickness6), new Vector(1, 0, 0), new Vector(0, 0, 1));
                        m_Symbolic.Outputs["Rod"] = rod;
                    }
                    else if (HgrCompareDoubleService.cmpdbl(blockClamp.Width1, 0)==true && HgrCompareDoubleService.cmpdbl(blockClamp.Length1, 0)==true && HgrCompareDoubleService.cmpdbl(blockClamp.Thickness1, 0)==true && midTopPlate.curvedEndY < 0)
                    {
                        Port rod = new Port(OccurrenceConnection, part, "Rod", new Position(0, 0, midTopPlate.length1 + midTopPlate.curvedEndY + blockClamp.Thickness6), new Vector(1, 0, 0), new Vector(0, 0, 1));
                        m_Symbolic.Outputs["Rod"] = rod;
                    }
                    else if (HgrCompareDoubleService.cmpdbl(blockClamp.Width1, 0)==true && HgrCompareDoubleService.cmpdbl(blockClamp.Length1, 0)==true && HgrCompareDoubleService.cmpdbl(blockClamp.Thickness1, 0)==true && HgrCompareDoubleService.cmpdbl(midTopPlate.curvedEndY, 0)==true && HgrCompareDoubleService.cmpdbl(midTopPlate.curvedEndRad, 0)==true)
                    {
                        Port rod = new Port(OccurrenceConnection, part, "Rod", new Position(0, 0, midTopPlate.length1 + Math.Abs(midTopPlate.trCornerY) + blockClamp.Thickness6), new Vector(1, 0, 0), new Vector(0, 0, 1));
                        m_Symbolic.Outputs["Rod"] = rod;
                    }
                    else
                    {
                        Port rod = new Port(OccurrenceConnection, part, "Rod", new Position(0, 0, blockClamp.Height1 + blockClamp.Thickness6), new Vector(1, 0, 0), new Vector(0, 0, 1));
                        m_Symbolic.Outputs["Rod"] = rod;
                    }
                }
                // Add Weld Port
                if (blockClamp.Diameter1 > 0 && blockClamp.Diameter2 > 0 && blockClamp.Offset1 > 0)
                {
                    if (blockClamp.Width3 > 0)
                    {
                        if (HgrCompareDoubleService.cmpdbl(blockClamp.Width2, 0)==true && HgrCompareDoubleService.cmpdbl(blockClamp.Length2, 0)==true && HgrCompareDoubleService.cmpdbl(blockClamp.Thickness2, 0)==true && midBotPlate.curvedEndY > 0)
                        {
                            Port weld = new Port(OccurrenceConnection, part, "Weld", new Position(blockClamp.Length3 / 2, blockClamp.Width3 / 2, -midBotPlate.length1 - midBotPlate.curvedEndY - blockClamp.Thickness4 - blockClamp.Thickness3), new Vector(1, 0, 0), new Vector(0, 0, -1));
                            m_Symbolic.Outputs["Weld"] = weld;
                        }
                        else if (HgrCompareDoubleService.cmpdbl(blockClamp.Width2, 0)==true && HgrCompareDoubleService.cmpdbl(blockClamp.Length2, 0)==true && HgrCompareDoubleService.cmpdbl(blockClamp.Thickness2, 0)==true && midBotPlate.curvedEndY < 0)
                        {
                            Port weld = new Port(OccurrenceConnection, part, "Weld", new Position(blockClamp.Length3 / 2, blockClamp.Width3 / 2, -midBotPlate.length1 - midBotPlate.curvedEndY - blockClamp.Thickness4 - blockClamp.Thickness3), new Vector(1, 0, 0), new Vector(0, 0, -1));
                            m_Symbolic.Outputs["Weld"] = weld;
                        }
                        else if (HgrCompareDoubleService.cmpdbl(blockClamp.Width2, 0)==true && HgrCompareDoubleService.cmpdbl(blockClamp.Length2, 0)==true && HgrCompareDoubleService.cmpdbl(blockClamp.Thickness2, 0)==true && HgrCompareDoubleService.cmpdbl(midBotPlate.curvedEndY, 0)==true && HgrCompareDoubleService.cmpdbl(midBotPlate.curvedEndRad, 0)==true)
                        {
                            Port weld = new Port(OccurrenceConnection, part, "Weld", new Position(blockClamp.Length3 / 2, blockClamp.Width3 / 2, -midBotPlate.length1 - blockClamp.Diameter1 / 2 - blockClamp.Thickness4 - blockClamp.Thickness3), new Vector(1, 0, 0), new Vector(0, 0, -1));
                            m_Symbolic.Outputs["Weld"] = weld;
                        }
                        else
                        {
                            Port weld = new Port(OccurrenceConnection, part, "Weld", new Position(blockClamp.Length3 / 2, blockClamp.Width3 / 2, -blockClamp.Height2 - blockClamp.Thickness4 - blockClamp.Thickness3), new Vector(1, 0, 0), new Vector(0, 0, -1));
                            m_Symbolic.Outputs["Weld"] = weld;
                        }
                    }
                    else if (blockClamp.Width4 > 0)
                    {
                        if (HgrCompareDoubleService.cmpdbl(blockClamp.Width2, 0)==true && HgrCompareDoubleService.cmpdbl(blockClamp.Length2, 0)==true && HgrCompareDoubleService.cmpdbl(blockClamp.Thickness2, 0)==true && midBotPlate.curvedEndY > 0)
                        {
                            Port weld = new Port(OccurrenceConnection, part, "Weld", new Position(blockClamp.Length4 / 2, blockClamp.Width4 / 2, -midBotPlate.length1 - midBotPlate.curvedEndY - blockClamp.Thickness4), new Vector(1, 0, 0), new Vector(0, 0, -1));
                            m_Symbolic.Outputs["Weld"] = weld;
                        }
                        else if (HgrCompareDoubleService.cmpdbl(blockClamp.Width2, 0)==true && HgrCompareDoubleService.cmpdbl(blockClamp.Length2, 0)==true && HgrCompareDoubleService.cmpdbl(blockClamp.Thickness2, 0)==true && midBotPlate.curvedEndY < 0)
                        {
                            Port weld = new Port(OccurrenceConnection, part, "Weld", new Position(blockClamp.Length4 / 2, blockClamp.Width4 / 2, -midBotPlate.length1 - midBotPlate.curvedEndY - blockClamp.Thickness4), new Vector(1, 0, 0), new Vector(0, 0, -1));
                            m_Symbolic.Outputs["Weld"] = weld;
                        }
                        else if (HgrCompareDoubleService.cmpdbl(blockClamp.Width2, 0)==true && HgrCompareDoubleService.cmpdbl(blockClamp.Length2, 0)==true && HgrCompareDoubleService.cmpdbl(blockClamp.Thickness2, 0)==true && HgrCompareDoubleService.cmpdbl(midBotPlate.curvedEndY, 0)==true && HgrCompareDoubleService.cmpdbl(midBotPlate.curvedEndRad, 0)==true)
                        {
                            Port weld = new Port(OccurrenceConnection, part, "Weld", new Position(blockClamp.Length4 / 2, blockClamp.Width4 / 2, -midBotPlate.length1 - blockClamp.Diameter1 / 2 - blockClamp.Thickness4), new Vector(1, 0, 0), new Vector(0, 0, -1));
                            m_Symbolic.Outputs["Weld"] = weld;
                        }
                        else
                        {
                            Port weld = new Port(OccurrenceConnection, part, "Weld", new Position(blockClamp.Length4 / 2, blockClamp.Width4 / 2, -blockClamp.Height2 - blockClamp.Thickness4), new Vector(1, 0, 0), new Vector(0, 0, -1));
                            m_Symbolic.Outputs["Weld"] = weld;
                        }
                    }
                    else if (HgrCompareDoubleService.cmpdbl(blockClamp.Width3, 0) == true && HgrCompareDoubleService.cmpdbl(blockClamp.Width4, 0) == true)
                    {
                        if (HgrCompareDoubleService.cmpdbl(blockClamp.Width2, 0)==true && HgrCompareDoubleService.cmpdbl(blockClamp.Length2, 0)==true && HgrCompareDoubleService.cmpdbl(blockClamp.Thickness2, 0)==true && midBotPlate.curvedEndY > 0)
                        {
                            Port weld = new Port(OccurrenceConnection, part, "Weld", new Position(midBotPlate.thickness1 / 2, midBotPlate.width1 / 2 + blockClamp.Offset1 / 2, -midBotPlate.length1 - midBotPlate.curvedEndY), new Vector(1, 0, 0), new Vector(0, 0, -1));
                            m_Symbolic.Outputs["Weld"] = weld;
                        }
                        else if (HgrCompareDoubleService.cmpdbl(blockClamp.Width2, 0)==true && HgrCompareDoubleService.cmpdbl(blockClamp.Length2, 0)==true && HgrCompareDoubleService.cmpdbl(blockClamp.Thickness2, 0)==true && midBotPlate.curvedEndY < 0)
                        {
                            Port weld = new Port(OccurrenceConnection, part, "Weld", new Position(midBotPlate.thickness1 / 2, midBotPlate.width1 / 2 + blockClamp.Offset1 / 2, -midBotPlate.length1 - midBotPlate.curvedEndY), new Vector(1, 0, 0), new Vector(0, 0, -1));
                            m_Symbolic.Outputs["Weld"] = weld;
                        }
                        else
                        {
                            Port weld = new Port(OccurrenceConnection, part, "Weld", new Position(midBotPlate.thickness1 / 2, midBotPlate.width1 / 2 + blockClamp.Offset1 / 2, -midBotPlate.length1 - blockClamp.Diameter1 / 2), new Vector(1, 0, 0), new Vector(0, 0, -1));
                            m_Symbolic.Outputs["Weld"] = weld;
                        }
                    }
                    else
                    {
                        Port weld = new Port(OccurrenceConnection, part, "Weld", new Position(blockClamp.Length2 / 2, blockClamp.Width2 / 2 + blockClamp.Offset1 / 2, -blockClamp.Height2), new Vector(1, 0, 0), new Vector(0, 0, -1));
                        m_Symbolic.Outputs["Weld"] = weld;
                    }
                }
                else
                {
                    if (blockClamp.Width3 > 0)
                    {
                        if (HgrCompareDoubleService.cmpdbl(blockClamp.Width2, 0)==true && HgrCompareDoubleService.cmpdbl(blockClamp.Length2, 0)==true && HgrCompareDoubleService.cmpdbl(blockClamp.Thickness2, 0)==true && midBotPlate.curvedEndY > 0)
                        {
                            Port weld = new Port(OccurrenceConnection, part, "Weld", new Position(blockClamp.Length3 / 2, blockClamp.Width3 / 2, -midBotPlate.length1 - midBotPlate.curvedEndY - blockClamp.Thickness4 - blockClamp.Thickness3), new Vector(1, 0, 0), new Vector(0, 0, -1));
                            m_Symbolic.Outputs["Weld"] = weld;
                        }
                        else if (HgrCompareDoubleService.cmpdbl(blockClamp.Width2, 0)==true && HgrCompareDoubleService.cmpdbl(blockClamp.Length2, 0)==true && HgrCompareDoubleService.cmpdbl(blockClamp.Thickness2, 0)==true && midBotPlate.curvedEndY < 0)
                        {
                            Port weld = new Port(OccurrenceConnection, part, "Weld", new Position(blockClamp.Length3 / 2, blockClamp.Width3 / 2, -midBotPlate.length1 - midBotPlate.curvedEndY - blockClamp.Thickness4 - blockClamp.Thickness3), new Vector(1, 0, 0), new Vector(0, 0, -1));
                            m_Symbolic.Outputs["Weld"] = weld;
                        }
                        else if (HgrCompareDoubleService.cmpdbl(blockClamp.Width2, 0)==true && HgrCompareDoubleService.cmpdbl(blockClamp.Length2, 0)==true && HgrCompareDoubleService.cmpdbl(blockClamp.Thickness2, 0)==true && HgrCompareDoubleService.cmpdbl(midBotPlate.curvedEndY, 0)==true && HgrCompareDoubleService.cmpdbl(midBotPlate.curvedEndRad, 0)==true)
                        {
                            Port weld = new Port(OccurrenceConnection, part, "Weld", new Position(blockClamp.Length3 / 2, blockClamp.Width3 / 2, -midBotPlate.length1 - blockClamp.Diameter1 / 2 - blockClamp.Thickness4 - blockClamp.Thickness3), new Vector(1, 0, 0), new Vector(0, 0, -1));
                            m_Symbolic.Outputs["Weld"] = weld;
                        }
                        else
                        {
                            Port weld = new Port(OccurrenceConnection, part, "Weld", new Position(blockClamp.Length3 / 2, blockClamp.Width3 / 2, -blockClamp.Height2 - blockClamp.Thickness4 - blockClamp.Thickness3), new Vector(1, 0, 0), new Vector(0, 0, -1));
                            m_Symbolic.Outputs["Weld"] = weld;
                        }
                    }
                    else if (blockClamp.Width4 > 0)
                    {
                        if (HgrCompareDoubleService.cmpdbl(blockClamp.Width2, 0)==true && HgrCompareDoubleService.cmpdbl(blockClamp.Length2, 0)==true && HgrCompareDoubleService.cmpdbl(blockClamp.Thickness2, 0)==true && midBotPlate.curvedEndY > 0)
                        {
                            Port weld = new Port(OccurrenceConnection, part, "Weld", new Position(blockClamp.Length4 / 2, blockClamp.Width4 / 2, -midBotPlate.length1 - midBotPlate.curvedEndY - blockClamp.Thickness4), new Vector(1, 0, 0), new Vector(0, 0, -1));
                            m_Symbolic.Outputs["Weld"] = weld;
                        }
                        else if (HgrCompareDoubleService.cmpdbl(blockClamp.Width2, 0)==true && HgrCompareDoubleService.cmpdbl(blockClamp.Length2, 0)==true && HgrCompareDoubleService.cmpdbl(blockClamp.Thickness2, 0)==true && midBotPlate.curvedEndY < 0)
                        {
                            Port weld = new Port(OccurrenceConnection, part, "Weld", new Position(blockClamp.Length4 / 2, blockClamp.Width4 / 2, -midBotPlate.length1 - midBotPlate.curvedEndY - blockClamp.Thickness4), new Vector(1, 0, 0), new Vector(0, 0, -1));
                            m_Symbolic.Outputs["Weld"] = weld;
                        }
                        else if (HgrCompareDoubleService.cmpdbl(blockClamp.Width2, 0)==true && HgrCompareDoubleService.cmpdbl(blockClamp.Length2, 0)==true && HgrCompareDoubleService.cmpdbl(blockClamp.Thickness2, 0)==true && HgrCompareDoubleService.cmpdbl(midBotPlate.curvedEndY, 0)==true && HgrCompareDoubleService.cmpdbl(midBotPlate.curvedEndRad, 0)==true)
                        {
                            Port weld = new Port(OccurrenceConnection, part, "Weld", new Position(blockClamp.Length4 / 2, blockClamp.Width4 / 2, -midBotPlate.length1 - blockClamp.Diameter1 / 2 - blockClamp.Thickness4), new Vector(1, 0, 0), new Vector(0, 0, -1));
                            m_Symbolic.Outputs["Weld"] = weld;
                        }
                        else
                        {
                            Port weld = new Port(OccurrenceConnection, part, "Weld", new Position(blockClamp.Length4 / 2, blockClamp.Width4 / 2, -blockClamp.Height2 - blockClamp.Thickness4), new Vector(1, 0, 0), new Vector(0, 0, -1));
                            m_Symbolic.Outputs["Weld"] = weld;
                        }
                    }
                    else if (HgrCompareDoubleService.cmpdbl(blockClamp.Width3, 0) == true && HgrCompareDoubleService.cmpdbl(blockClamp.Width4, 0) == true)
                    {
                        if (HgrCompareDoubleService.cmpdbl(blockClamp.Width2, 0)==true && HgrCompareDoubleService.cmpdbl(blockClamp.Length2, 0)==true && HgrCompareDoubleService.cmpdbl(blockClamp.Thickness2, 0)==true && midBotPlate.curvedEndY > 0)
                        {
                            Port weld = new Port(OccurrenceConnection, part, "Weld", new Position(midBotPlate.thickness1 / 2 + blockClamp.Offset3 / 2, midBotPlate.width1 / 2, -midBotPlate.length1 - midBotPlate.curvedEndY), new Vector(1, 0, 0), new Vector(0, 0, -1));
                            m_Symbolic.Outputs["Weld"] = weld;
                        }
                        else if (HgrCompareDoubleService.cmpdbl(blockClamp.Width2, 0)==true && HgrCompareDoubleService.cmpdbl(blockClamp.Length2, 0)==true && HgrCompareDoubleService.cmpdbl(blockClamp.Thickness2, 0)==true && midBotPlate.curvedEndY < 0)
                        {
                            Port weld = new Port(OccurrenceConnection, part, "Weld", new Position(midBotPlate.thickness1 / 2 + blockClamp.Offset3 / 2, midBotPlate.width1 / 2, -midBotPlate.length1 - midBotPlate.curvedEndY), new Vector(1, 0, 0), new Vector(0, 0, -1));
                            m_Symbolic.Outputs["Weld"] = weld;
                        }
                        else
                        {
                            Port weld = new Port(OccurrenceConnection, part, "Weld", new Position(midBotPlate.thickness1 / 2 + blockClamp.Offset3 / 2, midBotPlate.width1 / 2, -midBotPlate.length1 - blockClamp.Diameter1 / 2), new Vector(1, 0, 0), new Vector(0, 0, -1));
                            m_Symbolic.Outputs["Weld"] = weld;
                        }
                    }
                    else
                    {
                        Port weld = new Port(OccurrenceConnection, part, "Weld", new Position(blockClamp.Length2 / 2, blockClamp.Width2 / 2, -blockClamp.Height2), new Vector(1, 0, 0), new Vector(0, 0, -1));
                        m_Symbolic.Outputs["Weld"] = weld;
                    }
                }

                // Make Graphic

                AddBlockClamp(blockClamp, new Matrix4X4(), m_Symbolic.Outputs, "BlockClampShape1");

            }
            catch  //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of BlockClampp"));
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
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrWeightCG, "Error in Weight CG of BlockClamp"));
                }
            }
        }

        #endregion

    }

}

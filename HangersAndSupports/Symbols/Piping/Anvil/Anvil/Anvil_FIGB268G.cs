//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   Anvil_FIGB268G.cs
//    Anvil,Ingr.SP3D.Content.Support.Symbols.Anvil_FIGB268G
//   Author       :  BS
//   Creation Date:  02-05-2013
//   Description:
//   
//   Anvil_FIG268G.cs is same for Anvil_FIG98G.cs

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   02-05-2013      BS      CR-CP-233113 Convert HS_Anvil VB Project to C# .Net
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.ReferenceData.Middle.Services;
using System.Collections.ObjectModel;

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
    [SymbolVersion("1.0.0.0")]
    public class Anvil_FIGB268G : HangerComponentSymbolDefinition
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Anvil,Ingr.SP3D.Content.Support.Symbols.Anvil_FIGB268G"
        //----------------------------------------------------------------------------------


        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "SPAN", "SPAN", 0.999999)]
        public InputDouble m_dSPAN;
        [InputDouble(3, "P_USER", "P_USER", 0.999999)]
        public InputDouble m_dP_USER;
        [InputDouble(4, "P", "P", 0.999999)]
        public InputDouble m_dP;
        [InputDouble(5, "A", "A", 0.999999)]
        public InputDouble m_dA;
        [InputDouble(6, "B", "B", 0.999999)]
        public InputDouble m_dB;
        [InputDouble(7, "C", "C", 0.999999)]
        public InputDouble m_dC;
        [InputDouble(8, "F", "F", 0.999999)]
        public InputDouble m_dF;
        [InputDouble(9, "D", "D", 0.999999)]
        public InputDouble m_dD;
        [InputDouble(10, "G", "G", 0.999999)]
        public InputDouble m_dG;
        [InputDouble(11, "Z", "Z", 0.999999)]
        public InputDouble m_dZ;
        [InputDouble(12, "PIPE_DIA", "PIPE_DIA", 0.999999)]
        public InputDouble m_dPIPE_DIA;
        [InputString(13, "SECTION_SIZE", "SECTION_SIZE", "No Value")]
        public InputString m_oSECTION_SIZE;
        [InputDouble(14, "W", "W", 0.999999)]
        public InputDouble m_dW;
        [InputDouble(15, "HOT_LOAD", "HOT_LOAD", 0.999999)]
        public InputDouble m_dHOT_LOAD;
        [InputDouble(16, "DIR", "DIR", 1)]
        public InputDouble m_oDIR;
        [InputDouble(17, "WORKING_TRAV", "WORKING_TRAV", 0.999999)]
        public InputDouble m_dWORKING_TRAV;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("Port3", "Port3")]
        [SymbolOutput("Port4", "Port4")]
        [SymbolOutput("SPRING1", "SPRING1")]
        [SymbolOutput("SPRING2", "SPRING2")]
        [SymbolOutput("FLANGE1", "FLANGE1")]
        [SymbolOutput("FLANGE2", "FLANGE2")]
        [SymbolOutput("TOP1", "TOP1")]
        [SymbolOutput("TOP2", "TOP2")]
        [SymbolOutput("SECTIONS", "SECTIONS")]
        public AspectDefinition m_Symbolic;
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

                Double span = m_dSPAN.Value;
                Double P = m_dP_USER.Value;
                Double rodDiameter = m_dA.Value;
                Double B = m_dB.Value;
                Double C = m_dC.Value;
                Double F = m_dF.Value;
                Double D = m_dD.Value;
                Double G = m_dG.Value;
                Double Z = m_dZ.Value;
                Double pipeDiameter = m_dPIPE_DIA.Value;
                Double W = m_dW.Value;

                double sectionWidth = 0, sectionDepth = 0;
                if (P == 0)
                {
                    P = m_dP.Value;
                }

                double takeOut = P + Z - F;
                CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();
                PartClass anvilfig268GParts = (PartClass)catalogBaseHelper.GetPartClass("Anvil_C_SECTION");
                ReadOnlyCollection<BusinessObject> fig268GParts = anvilfig268GParts.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;

                foreach (BusinessObject part1 in fig268GParts)
                {
                    if (((string)((PropertyValueString)part1.GetPropertyValue("IJUAHgrAnvil_C_SECTION", "SIZE")).PropValue == m_oSECTION_SIZE.Value))
                    {
                        sectionDepth = (double)((PropertyValueDouble)part1.GetPropertyValue("IJUAHgrAnvil_C_SECTION", "DEPTH")).PropValue;
                        sectionWidth = (double)((PropertyValueDouble)part1.GetPropertyValue("IJUAHgrAnvil_C_SECTION", "WIDTH")).PropValue;
                        break;
                    }
                }
                //ports
                Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;
                port1 = new Port(OccurrenceConnection, part, "Other", new Position(0, 0, -pipeDiameter / 2), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port1;
                port1 = new Port(OccurrenceConnection, part, "LeftInThdRH", new Position(0, -span / 2, takeOut - pipeDiameter / 2), new Vector(1, 0, 0), new Vector(0, 0, -1));
                m_Symbolic.Outputs["Port3"] = port1;
                port1 = new Port(OccurrenceConnection, part, "RightInThdRH", new Position(0, span / 2, takeOut - pipeDiameter / 2), new Vector(1, 0, 0), new Vector(0, 0, -1));
                m_Symbolic.Outputs["Port4"] = port1;

                //Validating Inputs
                if (C <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidCGTZero, "C value should be greater than zero"));
                    return;
                }
                if (D <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidDGTZero, "D value should be greater than zero"));
                    return;
                }
                if (rodDiameter <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidRodDiaGTZero, "Rod diameter should be greater than zero"));
                    return;
                }
                if (G == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidGNZero, "G value cannot be zero"));
                    return;
                }
                if (W <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidWGTZero, "W value should be greater than zero"));
                    return;
                }
                if (span <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidSpanGTZero, "SPAN value should be greater than zero"));
                    return;
                }
                if (sectionWidth <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidSectionWidthGTZero, "Section width should be greater than zero"));
                    return;
                }
                if (sectionDepth <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidSectionDepthGTZero, "Section depth  should be greater than zero"));
                    return;
                }
                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                Matrix4X4 rotateMatrix = new Matrix4X4();

                symbolGeometryHelper.ActivePosition = new Position(-span / 2, 0, P - pipeDiameter / 2 - B + G);
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 0, 1).GetOrthogonalVector());
                Projection3d spring1 = symbolGeometryHelper.CreateCylinder(null, C / 2, B - G);
                rotateMatrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                spring1.Transform(rotateMatrix);
                m_Symbolic.Outputs["SPRING1"] = spring1;

                symbolGeometryHelper.ActivePosition = new Position(span / 2, 0, P - pipeDiameter / 2 - B + G);
                spring1 = symbolGeometryHelper.CreateCylinder(null, C / 2, B - G);
                rotateMatrix = new Matrix4X4();
                rotateMatrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                spring1.Transform(rotateMatrix);
                m_Symbolic.Outputs["SPRING2"] = spring1;

                symbolGeometryHelper.ActivePosition = new Position(-span / 2, 0, P - pipeDiameter / 2 - B);
                spring1 = symbolGeometryHelper.CreateCylinder(null, D / 2, G);
                rotateMatrix = new Matrix4X4();
                rotateMatrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                spring1.Transform(rotateMatrix);
                m_Symbolic.Outputs["FLANGE1"] = spring1;

                symbolGeometryHelper.ActivePosition = new Position(span / 2, 0, P - pipeDiameter / 2 - B);
                spring1 = symbolGeometryHelper.CreateCylinder(null, D / 2, G);
                rotateMatrix = new Matrix4X4();
                rotateMatrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                spring1.Transform(rotateMatrix);
                m_Symbolic.Outputs["FLANGE2"] = spring1;

                symbolGeometryHelper.ActivePosition = new Position(-span / 2, 0, P - pipeDiameter / 2 - (B * 0.75));
                spring1 = symbolGeometryHelper.CreateCylinder(null, 0.7 * rodDiameter, takeOut + F - P + (B * 0.75));
                rotateMatrix = new Matrix4X4();
                rotateMatrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                spring1.Transform(rotateMatrix);
                m_Symbolic.Outputs["TOP1"] = spring1;

                symbolGeometryHelper.ActivePosition = new Position(span / 2, 0, P - pipeDiameter / 2 - (B * 0.75));
                spring1 = symbolGeometryHelper.CreateCylinder(null, 0.7 * rodDiameter, takeOut + F - P + (B * 0.75));
                rotateMatrix = new Matrix4X4();
                rotateMatrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                spring1.Transform(rotateMatrix);
                m_Symbolic.Outputs["TOP2"] = spring1;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(-span / 2, -sectionWidth - W / 2, -pipeDiameter / 2 - sectionDepth);
                spring1 = symbolGeometryHelper.CreateBox(null, span, sectionWidth * 2 + W, sectionDepth, 9);
                rotateMatrix = new Matrix4X4();
                rotateMatrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                spring1.Transform(rotateMatrix);
                m_Symbolic.Outputs["SECTIONS"] = spring1;
            }
            catch       //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrConstructOutputs, "Error in constructOutputs of Anvil_FIGB268G"));
                    return;
                }
            }
        }
        #endregion

    }

}

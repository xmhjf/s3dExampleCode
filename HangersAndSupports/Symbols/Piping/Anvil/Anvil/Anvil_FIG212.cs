//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   Anvil_FIG212.cs
//    Anvil,Ingr.SP3D.Content.Support.Symbols.Anvil_FIG212
//   Author       :  Rajeswari
//   Creation Date:  09-May-2013
//   Description:
//   
//   Anvil_FIG212.cs is same for Anvil_FIG216.cs

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   09-May-2013  Rajeswari CR-CP-222292 Convert HS_Anvil VB Project to C# .Net
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
using System;
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
    [CacheOption(CacheOptionType.Cached)]
    [SymbolVersion("1.0.0.0")]
    public class Anvil_FIG212 : HangerComponentSymbolDefinition, ICustomHgrBOMDescription
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Anvil,Ingr.SP3D.Content.Support.Symbols.Anvil_FIG212"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "PIPE_DIA", "PIPE_DIA", 0.999999)]
        public InputDouble m_dPIPE_DIA;
        [InputDouble(3, "B", "B", 0.999999)]
        public InputDouble m_dB;
        [InputDouble(4, "C", "C", 0.999999)]
        public InputDouble m_dC;
        [InputDouble(5, "D", "D", 0.999999)]
        public InputDouble m_dD;
        [InputDouble(6, "E", "E", 0.999999)]
        public InputDouble m_dE;
        [InputDouble(7, "F", "F", 0.999999)]
        public InputDouble m_dF;
        [InputDouble(8, "G1", "G1", 0.999999)]
        public InputDouble m_dG1;
        [InputDouble(9, "G2", "G2", 0.999999)]
        public InputDouble m_dG2;
        [InputDouble(10, "H", "H", 0.999999)]
        public InputDouble m_dH;
        [InputDouble(11, "FINISH", "FINISH", 1)]
        public InputDouble m_oFINISH;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("TOP1", "TOP1")]
        [SymbolOutput("TOP2", "TOP2")]
        [SymbolOutput("BOT1", "BOT1")]
        [SymbolOutput("BOT2", "BOT2")]
        [SymbolOutput("TOP_BOLT", "TOP_BOLT")]
        [SymbolOutput("BOT_BOLT", "BOT_BOLT")]
        [SymbolOutput("BODY", "BODY")]
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

                Double pipeDiameter = m_dPIPE_DIA.Value;
                Double B = m_dB.Value;
                Double C = m_dC.Value;
                Double D = m_dD.Value;
                Double E = m_dE.Value;
                Double F = m_dF.Value;
                Double G1 = m_dG1.Value;
                Double G2 = m_dG2.Value;
                Double H = m_dH.Value;

                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                Matrix4X4 matrix = new Matrix4X4();
                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                //ports

                Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;
                Port port2 = new Port(OccurrenceConnection, part, "Pin", new Position(0, 0, (E - F / 2)), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;

                //Validating Inputs
                if (D <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidDGTZero, "D value should be greater than zero"));
                    return;
                }
                if (H <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidHGTZero, "H value should be greater than zero"));
                    return;
                }
                if (G1 <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidG1GTZero, "G1 value should be greater than zero"));
                    return;
                }
                if (G2 <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidG2GTZero, "G2 value should be greater than zero"));
                    return;
                }

                symbolGeometryHelper.ActivePosition = new Position(-(C / 2 + G1), -G2 / 2, pipeDiameter / 2);
                Projection3d top1 = symbolGeometryHelper.CreateBox(null, G1, G2, D - pipeDiameter / 2, 9);
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                top1.Transform(matrix);
                m_Symbolic.Outputs["TOP1"] = top1;

                symbolGeometryHelper = new SymbolGeometryHelper();
                matrix = new Matrix4X4();
                symbolGeometryHelper.ActivePosition = new Position(C / 2, -G2 / 2, pipeDiameter / 2);
                Projection3d top2 = symbolGeometryHelper.CreateBox(null, G1, G2, D - pipeDiameter / 2, 9);
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                top2.Transform(matrix);
                m_Symbolic.Outputs["TOP2"] = top2;

                symbolGeometryHelper = new SymbolGeometryHelper();
                matrix = new Matrix4X4();
                symbolGeometryHelper.ActivePosition = new Position(-(C / 2 + G1), -G2 / 2, -H);
                Projection3d bot1 = symbolGeometryHelper.CreateBox(null, G1, G2, H - pipeDiameter / 2, 9);
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                bot1.Transform(matrix);
                m_Symbolic.Outputs["BOT1"] = bot1;

                symbolGeometryHelper = new SymbolGeometryHelper();
                matrix = new Matrix4X4();
                symbolGeometryHelper.ActivePosition = new Position(C / 2, -G2 / 2, -H);
                Projection3d bot2 = symbolGeometryHelper.CreateBox(null, G1, G2, H - pipeDiameter / 2, 9);
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                bot2.Transform(matrix);
                m_Symbolic.Outputs["BOT2"] = bot2;

                symbolGeometryHelper = new SymbolGeometryHelper();
                Vector normal = new Position(0, C / 2 + 2 * G1, E).Subtract(new Position(0, -(C / 2 + 2 * G1), E));
                symbolGeometryHelper.ActivePosition = new Position(0, -(C / 2 + 2 * G1), E);
                symbolGeometryHelper.SetOrientation(normal, normal.GetOrthogonalVector());
                Projection3d topBolt = symbolGeometryHelper.CreateCylinder(null, F / 2, normal.Length);
                m_Symbolic.Outputs["TOP_BOLT"] = topBolt;

                symbolGeometryHelper = new SymbolGeometryHelper();
                normal = new Position(0, C / 2 + 2 * G1, -B).Subtract(new Position(0, -(C / 2 + 2 * G1), -B));
                symbolGeometryHelper.ActivePosition = new Position(0, -(C / 2 + 2 * G1), -B);
                symbolGeometryHelper.SetOrientation(normal, normal.GetOrthogonalVector());
                Projection3d botBolt = symbolGeometryHelper.CreateCylinder(null, F / 2, normal.Length);
                m_Symbolic.Outputs["BOT_BOLT"] = botBolt;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 0, 1).GetOrthogonalVector());
                Projection3d body = symbolGeometryHelper.CreateCylinder(null, pipeDiameter / 2 + G1, G2);
                matrix = new Matrix4X4();
                matrix.Rotate(-Math.PI / 2, new Vector(1, 0, 0));
                matrix.Translate(new Vector(0, -G2 / 2, 0));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                body.Transform(matrix);
                m_Symbolic.Outputs["BODY"] = body;
            }
            catch//General Unhandled exception 
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrConstructOutputs, "Error in constructOutputs of Anvil_FIG212"));
                return;
            }
        }
        #endregion

        #region "ICustomHgrBOMDescription Members"

        public string BOMDescription(BusinessObject oSupportOrComponent)
        {
            string bomDescription = "";
            try
            {
                Part part = (Part)oSupportOrComponent.GetRelationship("madeFrom", "part").TargetObjects[0];
                PropertyValueCodelist finishCodelist = (PropertyValueCodelist)oSupportOrComponent.GetPropertyValue("IJOAHgrFinish", "FINISH");
                if (finishCodelist.PropValue == 0)
                    finishCodelist.PropValue = 1;

                string finish = finishCodelist.PropertyInfo.CodeListInfo.GetCodelistItem(finishCodelist.PropValue).DisplayName;

                bomDescription = part.PartDescription + ", Finish: " + finish;
            }
            catch//General Unhandled exception 
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrBOMDescription, "Error in BOM description of Anvil_FIG212"));
            }
            return bomDescription;
        }
        #endregion

    }

}

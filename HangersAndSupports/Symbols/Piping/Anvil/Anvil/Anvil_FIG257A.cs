//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   Anvil_FIG257A.cs
//   Anvil,Ingr.SP3D.Content.Support.Symbols.Anvil_FIG257A
//   Author       : Vijaya 
//   Creation Date: 1-May-2013  
//   Description: Initial Creation-CR-CP-222292

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   1-May-2013    Vijaya    CR-CP-222292 Convert HS_Anvil VB Project to C# .Net
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

    [SymbolVersion("1.0.0.0")]
    [CacheOption(CacheOptionType.Cached)]
    public class Anvil_FIG257A : HangerComponentSymbolDefinition, ICustomHgrBOMDescription
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Anvil,Ingr.SP3D.Content.Support.Symbols.Anvil_FIG257A"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"
        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "PIPE_DIA", "PIPE_DIA", 0.999999)]
        public InputDouble m_dPIPE_DIA;
        [InputDouble(3, "FINISH", "FINISH", 1)]
        public InputDouble m_oFINISH;
        #endregion

        #region "Definitions of Aspects and their Outputs"
        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("WEB", "WEB")]
        [SymbolOutput("T_BASE", "T_BASE")]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
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
                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                Matrix4X4 rotateMatrix = new Matrix4X4();

                Double pipeDiameter = m_dPIPE_DIA.Value;
                double TL = 0.3048, TW = 0.104775, thickness = 0.00635, tHeight = 0.1016;

                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                //ports
                Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;

                Port port2 = new Port(OccurrenceConnection, part, "Structure", new Position(0, 0, -(pipeDiameter / 2 + tHeight)), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;

                //Validating Inputs
                if (thickness <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidThickness, "Thickness should be greater than zero"));
                    return;
                }
                if (TL <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidTLGTZero, "TL value should be greater than zero"));
                    return;
                }
                if (TW <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidTWGTZero, "TW value should be greater than zero"));
                    return;
                }
                if (tHeight <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidT_HEIGHT, "T_HEIGHT  should be greater than zero"));
                    return;
                }
                symbolGeometryHelper.ActivePosition = new Position(-thickness / 2, -TL / 2, pipeDiameter / 2);
                rotateMatrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                rotateMatrix.Rotate(Math.PI, new Vector(1, 0, 0), new Position(0, 0, 0));
                Projection3d web = symbolGeometryHelper.CreateBox(null, thickness, TL, tHeight - thickness, 9);
                web.Transform(rotateMatrix);
                m_Symbolic.Outputs["WEB"] = web;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(-TW / 2, -TL / 2, pipeDiameter / 2 + tHeight - thickness);
                rotateMatrix = new Matrix4X4();
                rotateMatrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                rotateMatrix.Rotate(Math.PI, new Vector(1, 0, 0), new Position(0, 0, 0));
                Projection3d tBase = symbolGeometryHelper.CreateBox(null, TW, TL, thickness, 9);
                tBase.Transform(rotateMatrix);
                m_Symbolic.Outputs["T_BASE"] = tBase;
            }
            catch  //General Unhandled exception 
            {
                if (base.ToDoListMessage == null)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrConstructOutputs, "Error in constructOutputs of Anvil_FIG257A"));
                    return;
                }
            }
        }
        #endregion

        #region "ICustomHgrBOMDescription Members"
        public string BOMDescription(BusinessObject oSupportOrComponent)
        {
            string bomDescription = "";
            try
            {
                Part catalogPart = (Part)oSupportOrComponent.GetRelationship("madeFrom", "part").TargetObjects[0];
                PropertyValueCodelist finishCodelist = (PropertyValueCodelist)oSupportOrComponent.GetPropertyValue("IJOAHgrShoeFinish", "FINISH");
                if (finishCodelist.PropValue == 0)
                    finishCodelist.PropValue = 1;
                string finish = finishCodelist.PropertyInfo.CodeListInfo.GetCodelistItem((int)finishCodelist.PropValue).ShortDisplayName;

                bomDescription = catalogPart.PartDescription + ", Finish: " + finish.ToString();
                return bomDescription;
            }
            catch
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrBOMDescription, "Error in BOM description of Anvil_FIG257A"));
                return "";
            }
        }
        #endregion
    }
}

//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   LRParts_hex_nut.cs
//    LRParts,Ingr.SP3D.Content.Support.Symbols.LRParts_hex_nut
//   Author       :  Rajeswari
//   Creation Date:  26/10/2012 
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
// 26/10/2012     Rajeswari   Initial Creation
//  26/03/2013     Rajeswari DI-CP-228142  Modify the error handling for delivered H&S symbols
//  30/10/2013     Vijaya    CR-CP-242533  Provide the ability to store GType outputs as a single blob.
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
    [VariableOutputs]
    public class LRParts_hex_nut : CustomSymbolDefinition
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "LRParts,Ingr.SP3D.Content.Support.Symbols.LRParts_hex_nut"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "ROD_DIA", "ROD_DIA", 0.999999)]
        public InputDouble m_ROD_DIA;
        [InputDouble(3, "W", "W", 0.999999)]
        public InputDouble m_W;
        [InputDouble(4, "T", "T", 0.999999)]
        public InputDouble m_T;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("NUT", "NUT")]
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
                Double W = m_W.Value;
                Double T = m_T.Value;
                if (W == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, LRPartsLocalizer.GetString(LRPartsResourceIDs.ErrInvalidW, "W cannot be zero or negative"));
                    return;
                }
                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();

                Port port1 = new Port(OccurrenceConnection, part, "InThdRH", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;

                Vector normal = (new Position(0, 0, -T).Subtract(new Position(0, 0, 0)));
                symbolGeometryHelper.ActivePosition = new Position(0, 0, 0);
                symbolGeometryHelper.SetOrientation(normal, new Vector(1, 0, 0));
                Projection3d nut = symbolGeometryHelper.CreateCylinder(null, W / (4 * Math.Cos(Math.PI * 30 / 180)) + W * (Math.Tan(Math.PI * 30 / 180)) / 2, normal.Length);
                m_Symbolic.Outputs["NUT"] = nut;

            }
            catch  //General Unhandled exception 
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, LRPartsLocalizer.GetString(LRPartsResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of LRParts_hex_nut.cs."));
                return;
            }
        }

        #endregion
    }
}


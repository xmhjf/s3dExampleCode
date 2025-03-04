//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   Connection.cs
//   SupTemporarySymbols,Ingr.SP3D.Content.Support.Symbols.Connection
//   Author       :  BS
//   Creation Date:  25-03-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   25-03-2013     BS      CR-CP-222300 Converted HgrSupTemporarySymbols VB Project to C# .Net
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

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
    public class Connection : HangerComponentSymbolDefinition
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "SupTemporarySymbols,Ingr.SP3D.Content.Support.Symbols.Connection"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;

        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Logic line", "Logic line")]
        [SymbolOutput("port1", "port 1")]
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
                //==========================================
                // create hgrports as part of the output
                //==========================================
                // z-axis of port is z-axis(global)
                // x-axis of port is x-axis(global)
                Port port1 = new Port(OccurrenceConnection, part, "Connection", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["port1"] = port1;
                Line3d line = new Line3d(OccurrenceConnection, new Position(0, 0, 0), new Position(0, 0, 0.0001));
                m_Symbolic.Outputs["Logic line"] = line;                
               
            }
            catch //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, SupTemporarySymbolsLocalizer.GetString(SupTemporarySymbolsSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of Connection"));
                }
            }
        }
        #endregion
    }
}

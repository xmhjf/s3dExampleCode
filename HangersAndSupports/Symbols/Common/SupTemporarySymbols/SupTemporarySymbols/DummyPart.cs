//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   CDuctClamp.cs
//    SupTemporarySymbols,Ingr.SP3D.Content.Support.Symbols.DummyPart
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
    public class DummyPart : HangerComponentSymbolDefinition
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "SupTemporarySymbols,Ingr.SP3D.Content.Support.Symbols.DummyPart"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "Length", "Length", 1)]
        public InputDouble m_Length;

        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Line", "Line")]
        [SymbolOutput("Structure Port", "Structure Port")]
        [SymbolOutput("Route Port", "Route Port")]
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
                 //Obtain the Units Of Measure Object.  The symbol definition assumes input
                 //in inches however graphics must be constructed in Data Base units.  The
                 //Units of Measure object does the conversion for us.
                double length = m_Length.Value;

                Line3d dummyLine = new Line3d(OccurrenceConnection, new Position(0, 0, 0), new Position(0, 0, length));
                m_Symbolic.Outputs["Line"] = dummyLine;
                //==========================================
                // Hanger Ports
                //==========================================
                //create hgrports as part of the output in the case of symbolic representation
                Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, length), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Route Port"] = port1;
                Port port2 = new Port(OccurrenceConnection, part, "Structure", new Position(0, 0, 0), new Vector(0, 1, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Structure Port"] = port2;

            }
            catch //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, SupTemporarySymbolsLocalizer.GetString(SupTemporarySymbolsSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of Dummy part"));
                }
            }
        }
        #endregion
    }
}

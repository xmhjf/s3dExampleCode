//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   Utility_Connection.cs
//    Utility,Ingr.SP3D.Content.Support.Symbols.Utility_Connection
//   Author       :  Hema
//   Creation Date:  29.10.2012
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   29.10.2012      Hema     CR-CP-222288  Converted HS_Utility VB Project to C# .Net  
//	 27/03/2013		 Hema 	  DI-CP-228142  Modify the error handling for delivered H&S symbols
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
    public class Utility_Connection : HangerComponentSymbolDefinition
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Utility,Ingr.SP3D.Content.Support.Symbols.Utility_Connection"
        //----------------------------------------------------------------------------------


        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "FlexPortRotX", "FlexPortRotX", 0.999999)]
        public InputDouble m_dFlexPortRotX;
        [InputDouble(3, "FlexPortRotY", "FlexPortRotY", 0.999999)]
        public InputDouble m_dFlexPortRotY;
        [InputDouble(4, "FlexPortRotZ", "FlexPortRotZ", 0.999999)]
        public InputDouble m_dFlexPortRotZ;

        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("FlexPort", "FlexPort")]
        [SymbolOutput("ConnectionPort", "ConnectionPort")]
        public AspectDefinition m_oSymbolic;

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

                Double flexPortRotX = m_dFlexPortRotX.Value;
                Double flexPortRotY = m_dFlexPortRotY.Value;
                Double flexPortRotZ = m_dFlexPortRotZ.Value;

                //=================================================
                //Construction of Physical Aspect 
                //=================================================

                Matrix4X4 transMatrix = new Matrix4X4();
                Vector tempVector = new Vector(1, 0, 0);
                transMatrix.Rotate(flexPortRotX, tempVector);

                Vector tempVector1 = new Vector(0, 1, 0);
                transMatrix.Rotate(flexPortRotY, tempVector1);

                Vector tempVector2 = new Vector(0, 0, 1);
                transMatrix.Rotate(flexPortRotZ, tempVector2);

                //ports
                Port port1 = new Port(OccurrenceConnection, part, "FlexPort", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_oSymbolic.Outputs["FlexPort"] = port1;

                Port port2 = new Port(OccurrenceConnection, part, "Connection", new Position(0, 0, 0), new Vector(transMatrix.GetIndexValue(0), transMatrix.GetIndexValue(1), transMatrix.GetIndexValue(2)), new Vector(transMatrix.GetIndexValue(8), transMatrix.GetIndexValue(9), transMatrix.GetIndexValue(10)));
                m_oSymbolic.Outputs["ConnectionPort"] = port2;

            }
            catch
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of Utility_Connection"));
                    return;
                }
            }
        }
        #endregion
    }
}


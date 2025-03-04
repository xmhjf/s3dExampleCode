//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   Spring.cs
//    Power,Ingr.SP3D.Content.Support.Symbols.Spring
//   Author       :  Rajeswari
//   Creation Date:  12-Dec-2012
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   12-Dec-2012  Rajeswari  CR222282 .Net HS_Power project creation
//	 25-Mar-2013    Hema 	 DI-CP-228142  Modify the error handling for delivered H&S symbols
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
    [VariableOutputs]
    [CacheOption(CacheOptionType.NonCached)] 
    [SymbolVersion("1.0.0.0")]
   public class Spring : HangerComponentSymbolDefinition
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Power,Ingr.SP3D.Content.Support.Symbols.Spring"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"
        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "TakeOut", "TakeOut", 0.999999)]
        public InputDouble m_dTakeOut;
        [InputDouble(3, "A", "A", 0.999999)]
        public InputDouble m_dA;
        [InputDouble(4, "B", "B", 0.999999)]
        public InputDouble m_dB;
        [InputDouble(5, "C", "C", 0.999999)]
        public InputDouble m_dC;
        [InputDouble(6, "F", "F", 0.999999)]
        public InputDouble m_dF;
        [InputDouble(7, "D", "D", 0.999999)]
        public InputDouble m_dD;
        [InputDouble(8, "G", "G", 0.999999)]
        public InputDouble m_dG;
        [InputDouble(9, "Z", "Z", 0.999999)]
        public InputDouble m_dZ;

        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("BODY", "BODY")]
        [SymbolOutput("FLANGE", "FLANGE")]
        [SymbolOutput("BOTTOM", "BOTTOM")]
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
                               
                Double rodDiameter = m_dA.Value;
                Double B = m_dB.Value;
                Double C = m_dC.Value;
                Double F = m_dF.Value;
                Double D = m_dD.Value;
                Double G = m_dG.Value;
                Double Z = m_dZ.Value;
                Double takeOut = m_dTakeOut.Value;

                if (C <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PowerLocalizer.GetString(PowerSymbolResourceIDs.ErrInvalidCGTZero, "C should be greater than zero"));
                    return;
                }
                if (D <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PowerLocalizer.GetString(PowerSymbolResourceIDs.ErrInvalidDGTZero, "D should be greater than zero"));
                    return;
                }

                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                //ports
                
                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
               
                Port port1 = new Port(OccurrenceConnection, part, "BotInThdRH", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;
                Port port2 = new Port(OccurrenceConnection, part, "TopInThdRH", new Position(0, 0, takeOut), new Vector(1, 0, 0), new Vector(0, 0, -1));
                m_Symbolic.Outputs["Port2"] = port2;

                symbolGeometryHelper.ActivePosition = new Position(0, 0, G + takeOut - B);
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(1, 0, 0));
                Projection3d body = symbolGeometryHelper.CreateCylinder(null, C / 2.0, B - G);
                m_Symbolic.Outputs["BODY"] = body;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, 0, takeOut);
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(1, 0, 0));
                Projection3d flange = symbolGeometryHelper.CreateCylinder(null, D / 2.0,G);
                m_Symbolic.Outputs["FLANGE"] = flange;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, 0, -F);
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(1, 0, 0));
                Projection3d bottom = symbolGeometryHelper.CreateCylinder(null, 0.7 * rodDiameter, Z + (B * 0.75));
                m_Symbolic.Outputs["BOTTOM"] = bottom;
            }
            catch
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PowerLocalizer.GetString(PowerSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of Spring"));
                    return;
                }
            }
        }
        #endregion
    }
}

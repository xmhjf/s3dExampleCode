//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   SFS5394_SLEEVE.cs
//    FINL_Parts,Ingr.SP3D.Content.Support.Symbols.SFS5394_SLEEVE
//   Author       :  Vijay
//   Creation Date:  18.03.2013
//   Description: CR-CP-222272 Convert HS_FINL_Parts VB Project to C# .Net  

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   18.03.2013     Vijay   CR-CP-222272 Convert HS_FINL_Parts VB Project to C# .Net  
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.Support.Middle;

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
    public class SFS5394_SLEEVE : HangerComponentSymbolDefinition
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "FINL_Parts,Ingr.SP3D.Content.Support.Symbols.SFS5394_SLEEVE"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "D", "D", 0.999999)]
        public InputDouble m_dD;
        [InputDouble(3, "S", "S", 0.999999)]
        public InputDouble m_dS;
        [InputDouble(4, "C", "C", 0.999999)]
        public InputDouble m_dC;
        [InputDouble(5, "B", "B", 0.999999)]
        public InputDouble m_dB;
        [InputDouble(6, "L", "L", 0.999999)]
        public InputDouble m_dL;
        [InputDouble(7, "L1", "L1", 0.999999)]
        public InputDouble m_dL1;
        [InputDouble(8, "L2", "L2", 0.999999)]
        public InputDouble m_dL2;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("BOX", "BOX")]
        [SymbolOutput("CYL", "CYL")]
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
                 
               

                Double D = m_dD.Value;
                Double S = m_dS.Value;
                Double C = m_dC.Value;
                Double B = m_dB.Value;
                Double L = m_dL.Value;
                Double L1 = m_dL1.Value;
                Double L2 = m_dL2.Value;
                //ports

                Port port1 = new Port(OccurrenceConnection, part, "Hole", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;

                Port port2 = new Port(OccurrenceConnection, part, "Pin", new Position(0, 0, -C + L - L2), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;

                if (D <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, FINL_PartsLocalizer.GetString(FINL_PartsSymbolResourceIDs.ErrInvalidDGTZero, "D value should be greater than zero"));
                    return;
                }
                if (S <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, FINL_PartsLocalizer.GetString(FINL_PartsSymbolResourceIDs.ErrInvalidSGTZ, "S value should be greater than zero"));
                    return;
                }
                if (B <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, FINL_PartsLocalizer.GetString(FINL_PartsSymbolResourceIDs.ErrInvalidBGTZero, "B value should be greater than zero"));
                    return;
                }
                if (L1 <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, FINL_PartsLocalizer.GetString(FINL_PartsSymbolResourceIDs.ErrInvalidL1GTZero, "L1 value should be greater than zero"));
                    return;
                }
                if (C == 0 && L == 0 && L1 == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, FINL_PartsLocalizer.GetString(FINL_PartsSymbolResourceIDs.ErrInvalidCandLandL1, "C,L and L1 values cannot be zero"));
                    return;
                }

                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                //=================================================
                //Construction of Physical Aspect 
                //=================================================                

                symbolGeometryHelper.ActivePosition = new Position(-S, -B / 2, -C);
                Projection3d top1Box = symbolGeometryHelper.CreateBox(null, S * 2, B, L1, 9);
                m_Symbolic.Outputs["BOX"] = top1Box;

                symbolGeometryHelper = new SymbolGeometryHelper();
                Vector normal = new Position(0, 0, -C + L).Subtract(new Position(0, 0, -C + L1));
                symbolGeometryHelper.ActivePosition = new Position(0, 0, -C + L1);
                symbolGeometryHelper.SetOrientation(normal, normal.GetOrthogonalVector());
                Projection3d topBoltCylinder = symbolGeometryHelper.CreateCylinder(null, D / 2 + S, normal.Length);
                m_Symbolic.Outputs["CYL"] = topBoltCylinder;
                }
            catch
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, FINL_PartsLocalizer.GetString(FINL_PartsSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of SFS5394_SLEEVE.cs."));
                    return;
                }
            }
        }
        #endregion
    }
}

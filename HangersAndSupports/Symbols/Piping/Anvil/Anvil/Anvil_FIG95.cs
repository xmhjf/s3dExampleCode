//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   Anvil_FIG95.cs
//    Anvil,Ingr.SP3D.Content.Support.Symbols.Anvil_FIG95
//   Author       :  Rajeswari
//   Creation Date:  13-05-2013
//   Description:
//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   13-05-2013      Rajeswari  CR-CP-233113 Convert HS_Anvil VB Project to C# .Net
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
using System;
using System.Collections.ObjectModel;
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
    public class Anvil_FIG95 : HangerComponentSymbolDefinition
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Anvil,Ingr.SP3D.Content.Support.Symbols.Anvil_FIG95"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "ROD_DIA", "ROD_DIA", 0.999999)]
        public InputDouble m_dROD_DIA;
        [InputDouble(3, "C", "C", 0.999999)]
        public InputDouble m_dC;
        [InputDouble(4, "D", "D", 0.999999)]
        public InputDouble m_dD;
        [InputDouble(5, "FLANGE_T", "FLANGE_T", 0.999999)]
        public InputDouble m_dFLANGE_T;
        [InputDouble(6, "FINISH", "FINISH", 1)]
        public InputDouble m_oFINISH;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("BODY", "BODY")]
        [SymbolOutput("BOLT", "BOLT")]
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

                Double A = m_dROD_DIA.Value;
                Double C = m_dC.Value;
                Double D = m_dD.Value;
                Double flangeT = m_dFLANGE_T.Value;
                Double H = C - A;
                Double E = H / 2;
                const double const_1 = 0.01905, const_2 = 0.009525, const_3 = 0.00762;
                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                Matrix4X4 rotateMatrix = new Matrix4X4();
                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                //ports

                Port port1 = new Port(OccurrenceConnection, part, "InThdRH", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;
                port1 = new Port(OccurrenceConnection, part, "Structure", new Position(0, E, flangeT / 2), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port1;

                //Validating Inputs
                if (A <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidAGTZero, "A value should be greater than zero"));
                    return;
                }
                if (D == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidDNZero, "D value cannot be zero"));
                    return;
                }
                Collection<Position> pointCollection = new Collection<Position>();
                pointCollection.Add(new Position(-A, -A, 0));
                pointCollection.Add(new Position(-A, E, 0));
                pointCollection.Add(new Position(-A, E, const_1));
                pointCollection.Add(new Position(-A, -A, const_1));
                pointCollection.Add(new Position(-A, -A, D / 2 + const_2));
                pointCollection.Add(new Position(-A, H, D / 2 + const_2));
                pointCollection.Add(new Position(-A, H, const_2 - D / 2));
                pointCollection.Add(new Position(-A, -A, const_2 - D / 2));
                pointCollection.Add(new Position(-A, -A, 0));
                Projection3d body = new Projection3d(new LineString3d(pointCollection), new Vector(1, 0, 0), 2 * A, true);
                m_Symbolic.Outputs["BODY"] = body;

                symbolGeometryHelper.ActivePosition = new Position(0, 0, const_3);
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 0, 1).GetOrthogonalVector());
                Projection3d bolt = symbolGeometryHelper.CreateCylinder(null, A / 2, 1.1 * D);
                rotateMatrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                bolt.Transform(rotateMatrix);
                m_Symbolic.Outputs["BOLT"] = bolt;

                symbolGeometryHelper = new SymbolGeometryHelper();
                Vector normal = new Position(0, 0, D / 2 + const_2 + A).Subtract(new Position(0, 0, D / 2 + const_2));
                symbolGeometryHelper.ActivePosition = new Position(0, 0, D / 2 + const_2);
                symbolGeometryHelper.SetOrientation(normal, normal.GetOrthogonalVector());
                Projection3d nut = symbolGeometryHelper.CreateCylinder(null, 0.75 * A, normal.Length);
                m_Symbolic.Outputs["NUT"] = nut;
            }
            catch//General Unhandled exception 
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrConstructOutputs, "Error in constructOutputs of Anvil_FIG95"));
                return;
            }
        }
        #endregion

    }

}

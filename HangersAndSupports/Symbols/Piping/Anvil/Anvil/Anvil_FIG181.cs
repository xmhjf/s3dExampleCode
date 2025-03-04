//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   Anvil_FIG181.cs
//    Anvil,Ingr.SP3D.Content.Support.Symbols.Anvil_FIG181
//   Author       :  Rajeswari
//   Creation Date:  09-May-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   09-May-2013  Rajeswari CR-CP-222292 Convert HS_Anvil VB Project to C# .Net
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
    [CacheOption(CacheOptionType.Cached)]
    [SymbolVersion("1.0.0.0")]
    public class Anvil_FIG181 : HangerComponentSymbolDefinition
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Anvil,Ingr.SP3D.Content.Support.Symbols.Anvil_FIG181"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "E", "E", 0.999999)]
        public InputDouble m_dE;
        [InputDouble(3, "B", "B", 0.999999)]
        public InputDouble m_dB;
        [InputDouble(4, "D", "D", 0.999999)]
        public InputDouble m_dD;
        [InputDouble(5, "C", "C", 0.999999)]
        public InputDouble m_dC;
        [InputDouble(6, "F", "F", 0.999999)]
        public InputDouble m_dF;
        [InputDouble(7, "G1", "G1", 0.999999)]
        public InputDouble m_dG1;
        [InputDouble(8, "G2", "G2", 0.999999)]
        public InputDouble m_dG2;
        [InputDouble(9, "FINISH", "FINISH", 1)]
        public InputDouble m_oFINISH;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("BODY", "BODY")]
        [SymbolOutput("ROLL", "ROLL")]
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

                Double E = m_dE.Value;
                Double B = m_dB.Value;
                Double D = m_dD.Value;
                Double C = m_dC.Value;
                Double F = m_dF.Value;
                Double G1 = m_dG1.Value;
                Double G2 = m_dG2.Value;

                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                //ports

                Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;
                Port port2 = new Port(OccurrenceConnection, part, "Hole", new Position(0, 0, E), new Vector(1, 0, 0), new Vector(0, 0, -1));
                m_Symbolic.Outputs["Port2"] = port2;
                
                //Validating Inputs
                if (F <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidFGTZero, "F value should be greater than zero"));
                    return;
                }
                if (G2 == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidG2NZero, "G2 value cannot be zero"));
                    return;
                }
                Collection<ICurve> lineCollection = new Collection<ICurve>();
                lineCollection.Add(new Line3d(new Position(-G2 / 2, -C / 2 - G1, (-D) - F * 1.5), new Position(-G2 / 2, -C / 2 - G1, B - D - Math.Tan(30 * Math.PI / 180) * (C / 2 - G1 * 2))));
                lineCollection.Add(new Line3d(new Position(-G2 / 2, -C / 2 - G1, B - D - Math.Tan(30 * Math.PI / 180) * (C / 2 - G1 * 2)), new Position(-G2 / 2, -C / 2 + Math.Cos(30 * Math.PI / 180) * (C / 2 - G1 * 2), B - D)));
                lineCollection.Add(new Line3d(new Position(-G2 / 2, -C / 2 + Math.Cos(30 * Math.PI / 180) * (C / 2 - G1 * 2), B - D), new Position(-G2 / 2, C / 2 - Math.Cos(30 * Math.PI / 180) * (C / 2 - G1 * 2), B - D)));
                lineCollection.Add(new Line3d(new Position(-G2 / 2, C / 2 - Math.Cos(30 * Math.PI / 180) * (C / 2 - G1 * 2), B - D), new Position(-G2 / 2, C / 2 + G1, B - D - Math.Tan(30 * Math.PI / 180) * (C / 2 - G1 * 2))));
                lineCollection.Add(new Line3d(new Position(-G2 / 2, C / 2 + G1, B - D - Math.Tan(30 * Math.PI / 180) * (C / 2 - G1 * 2)), new Position(-G2 / 2, C / 2 + G1, -D - F * 1.5)));
                lineCollection.Add(new Line3d(new Position(-G2 / 2, C / 2 + G1, -D - F * 1.5), new Position(-G2 / 2, C / 2, -D - F * 1.5)));
                lineCollection.Add(new Line3d(new Position(-G2 / 2, C / 2, -D - F * 1.5), new Position(-G2 / 2, C / 2, B - D - Math.Tan(30 * Math.PI / 180) * (C / 2 - G1 * 2) - Math.Tan(30 * Math.PI / 180) * G1)));
                lineCollection.Add(new Line3d(new Position(-G2 / 2, C / 2, B - D - Math.Tan(30 * Math.PI / 180) * (C / 2 - G1 * 2) - Math.Tan(30 * Math.PI / 180) * G1), new Position(-G2 / 2, C / 2 - Math.Cos(30 * Math.PI / 180) * (C / 2 - G1 * 2) - Math.Tan(30 * Math.PI / 180) * G1, B - D - G1)));
                lineCollection.Add(new Line3d(new Position(-G2 / 2, C / 2 - Math.Cos(30 * Math.PI / 180) * (C / 2 - G1 * 2) - Math.Tan(30 * Math.PI / 180) * G1, B - D - G1), new Position(-G2 / 2, -C / 2 + Math.Cos(30 * Math.PI / 180) * (C / 2 - G1 * 2) + Math.Tan(30 * Math.PI / 180) * G1, B - D - G1)));
                lineCollection.Add(new Line3d(new Position(-G2 / 2, -C / 2 + Math.Cos(30 * Math.PI / 180) * (C / 2 - G1 * 2) + Math.Tan(30 * Math.PI / 180) * G1, B - D - G1), new Position(-G2 / 2, -C / 2, B - D - Math.Tan(30 * Math.PI / 180) * (C / 2 - G1 * 2) - Math.Tan(30 * Math.PI / 180) * G1)));
                lineCollection.Add(new Line3d(new Position(-G2 / 2, -C / 2, B - D - Math.Tan(30 * Math.PI / 180) * (C / 2 - G1 * 2) - Math.Tan(30 * Math.PI / 180) * G1), new Position(-G2 / 2, -C / 2, -D - F * 1.5)));
                lineCollection.Add(new Line3d(new Position(-G2 / 2, -C / 2, -D - F * 1.5), new Position(-G2 / 2, -C / 2 - G1, -D - F * 1.5)));

                Projection3d body = new Projection3d(new ComplexString3d(lineCollection), new Vector(1, 0, 0), G2, true);
                m_Symbolic.Outputs["BODY"] = body;

                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                Vector normal = new Position(0, C / 2 + G1 * 4, -D).Subtract(new Position(0, -C / 2 - G1 * 4, -D));
                symbolGeometryHelper.ActivePosition = new Position(0, -C / 2 - G1 * 4, -D);
                symbolGeometryHelper.SetOrientation(normal, normal.GetOrthogonalVector());
                Projection3d roll = symbolGeometryHelper.CreateCylinder(null, F / 2, normal.Length);
                m_Symbolic.Outputs["ROLL"] = roll;
            }
            catch//General Unhandled exception 
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrConstructOutputs, "Error in constructOutputs of Anvil_FIG181"));
                return;
            }
        }
        #endregion

    }

}

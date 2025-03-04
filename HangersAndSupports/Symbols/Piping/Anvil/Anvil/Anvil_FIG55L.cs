//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   Anvil_FIG55L.cs
//   Anvil,Ingr.SP3D.Content.Support.Symbols.Anvil_FIG55L
//   Author       :  Hema
//   Creation Date:  03-05-2013
//   Description:    Converted Anvil_Parts VB Project to C#.Net Project
//   
//   Anvil_FIG55L.cs is same for Anvil_FIG55S.cs

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   03-05-2013     Hema     CR-CP-222292 Convert HS_Anvil VB Project to C# .Net
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

    public class Anvil_FIG55L : HangerComponentSymbolDefinition
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Anvil,Ingr.SP3D.Content.Support.Symbols.Anvil_FIG55L"
        //----------------------------------------------------------------------------------


        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "W", "W", 0.999999)]
        public InputDouble m_dW;
        [InputDouble(3, "F", "F", 0.999999)]
        public InputDouble m_dF;
        [InputDouble(4, "H", "H", 0.999999)]
        public InputDouble m_dH;
        [InputDouble(5, "T", "T", 0.999999)]
        public InputDouble m_dT;
        [InputDouble(6, "FINISH", "FINISH", 1)]
        public InputDouble m_oFINISH;

        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("BODY", "BODY")]
        [SymbolOutput("HOLE", "HOLE")]
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

                Double holeSize = m_dF.Value;
                Double W = m_dW.Value;
                Double F = m_dF.Value;
                Double H = m_dH.Value;
                Double T = m_dT.Value;

                //Initializing SymbolGeometryHelper
                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                Matrix4X4 matrix = new Matrix4X4();

                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                //ports
                Port port1 = new Port(OccurrenceConnection, part, "Structure", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;
                Port port2 = new Port(OccurrenceConnection, part, "Hole", new Position(0, 0, -(H + holeSize / 2)), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;

                //Validating Inputs
                if (T == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidTNZero, "T value cannot be zero"));
                    return;
                }
                if (holeSize <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidHoleSizeGTZero, "Hole size should be greater than zero"));
                    return;
                }
                Collection<ICurve> curveCollection = new Collection<ICurve>();
               
                symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                Arc3d arc = symbolGeometryHelper.CreateArc(null, W / 2, Math.PI);
                matrix.Rotate(Math.PI, new Vector(0, 0, 1));
                matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0));
                matrix.Translate(new Vector(0, -T / 2, -H));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                arc.Transform(matrix);
                curveCollection.Add(arc);

                curveCollection.Add(new Line3d(new Position(-T / 2, W / 2, -H), new Position(-T / 2, W / 2, 0)));
                curveCollection.Add(new Line3d(new Position(-T / 2, W / 2, 0), new Position(-T / 2, -W / 2, 0)));
                curveCollection.Add(new Line3d(new Position(-T / 2, -W / 2, 0), new Position(-T / 2, -W / 2, -H)));

                Projection3d body = new Projection3d(new ComplexString3d(curveCollection), new Vector(1, 0, 0), T, true);
                m_Symbolic.Outputs["BODY"] = body;

                symbolGeometryHelper = new SymbolGeometryHelper();
                Vector normal = new Position(T / 2, 0, -H).Subtract(new Position(-T / 2, 0, -H));
                symbolGeometryHelper.ActivePosition = new Position(-T / 2, 0, -H);
                symbolGeometryHelper.SetOrientation(normal, normal.GetOrthogonalVector());
                Projection3d hole = symbolGeometryHelper.CreateCylinder(null, holeSize / 2, normal.Length);
                m_Symbolic.Outputs["HOLE"] = hole;
            }
            catch       //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrConstructOutputs, "Error in constructOutputs of Anvil_FIG55"));
                    return;
                }
            }
        }
        #endregion

    }

}

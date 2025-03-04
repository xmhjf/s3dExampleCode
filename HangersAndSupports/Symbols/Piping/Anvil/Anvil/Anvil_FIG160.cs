//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   Anvil_FIG160.cs
//    Anvil,Ingr.SP3D.Content.Support.Symbols.Anvil_FIG160
//   Author       :  Vijay
//   Creation Date:  30-04-2013
//   Description:
//   
//   Anvil_FIG160.cs is same for Anvil_FIG161.cs,Anvil_FIG162.cs,Anvil_FIG163.cs,Anvil_FIG164.cs,Anvil_FIG165.cs

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   30-04-2013     Vijay    CR-CP-222292 Convert HS_Anvil VB Project to C# .Net 
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using System.Collections.ObjectModel;
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
    public class Anvil_FIG160 : HangerComponentSymbolDefinition
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Anvil,Ingr.SP3D.Content.Support.Symbols.Anvil_FIG160"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "PIPE_DIA", "PIPE_DIA", 0.999999)]
        public InputDouble m_dPIPE_DIA;
        [InputDouble(3, "C", "C", 0.999999)]
        public InputDouble m_dC;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("BODY", "BODY")]
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
                
                Double pipeDiameter = m_dPIPE_DIA.Value;
                Double C = m_dC.Value;
                Double L = 0.3048;

                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                
                Collection<ICurve> curveCollection = new Collection<ICurve>();
                Matrix4X4 rotateMatrix = new Matrix4X4();

                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                //ports

                Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;

                Port port2 = new Port(OccurrenceConnection, part, "Structure", new Position(0, 0, -C), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;

                //Validating Inputs
                if (C == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidCNZero, "C value should not be lessthan zero."));
                    return;
                }

                curveCollection.Add(new Line3d(new Position(-L / 2, pipeDiameter / 2 * Math.Sin(30 * (Math.PI / 180)), -pipeDiameter / 2 * Math.Cos(30 * (Math.PI / 180))), new Position(-L / 2, C * Math.Sin(30 * (Math.PI / 180)), -C * Math.Cos(30 * (Math.PI / 180)))));

                symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                Arc3d innerArc = symbolGeometryHelper.CreateArc(null, C, 2 * Math.PI / 6);
                rotateMatrix.Rotate(2 * Math.PI / 6, new Vector(0, 0, 1));
                rotateMatrix.Rotate(Math.PI / 2, new Vector(1, 0, 0));
                rotateMatrix.Translate(new Vector(0, -L / 2, 0));
                rotateMatrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                rotateMatrix.Rotate(Math.PI, new Vector(1, 0, 0), new Position(0, 0, 0));
                innerArc.Transform(rotateMatrix);
                curveCollection.Add(innerArc);

                curveCollection.Add(new Line3d(new Position(-L / 2, -pipeDiameter / 2 * Math.Sin(30 * (Math.PI / 180)), -pipeDiameter / 2 * Math.Cos(30 * (Math.PI / 180))), new Position(-L / 2, -C * Math.Sin(30 * (Math.PI / 180)), -C * Math.Cos(30 * (Math.PI / 180)))));

                Projection3d body = new Projection3d(new ComplexString3d(curveCollection), new Vector(1, 0, 0), L, true);
                m_Symbolic.Outputs["BODY"] = body;
            }
            catch  //General Unhandled exception 
            {
                if (base.ToDoListMessage == null)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrConstructOutputs, "Error in constructOutputs of Anvil_FIG160"));
                    return;
                }
            }
        }
        #endregion
    }
}

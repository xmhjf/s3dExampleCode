//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   LRParts_FIG55NS.cs
//    LRParts,Ingr.SP3D.Content.Support.Symbols.LRParts_FIG55NS
//   Author       :  Hema
//   Creation Date: 18-10-2012 
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   18-10-2012      Hema    Initial Creation
//  26/03/2013     Rajeswari DI-CP-228142  Modify the error handling for delivered H&S symbols
//  30/10/2013     Vijaya    CR-CP-242533  Provide the ability to store GType outputs as a single blob.
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
    [VariableOutputs]
    public class LRParts_FIG55NS : CustomSymbolDefinition
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "LRParts,Ingr.SP3D.Content.Support.Symbols.LRParts_FIG55NS"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "W", "W", 0.999999)]
        public InputDouble m_W;
        [InputDouble(3, "F", "F", 0.999999)]
        public InputDouble m_F;
        [InputDouble(4, "H", "H", 0.999999)]
        public InputDouble m_H;
        [InputDouble(5, "T", "T", 0.999999)]
        public InputDouble m_T;
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
                Double W = m_W.Value;
                Double holeSize = m_F.Value;
                Double H = m_H.Value;
                Double T = m_T.Value;

                if (T == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, LRPartsLocalizer.GetString(LRPartsResourceIDs.ErrInvalidT_55, "T cannot be zero"));
                    return;
                }
                if (W == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, LRPartsLocalizer.GetString(LRPartsResourceIDs.ErrInvalidW_55, "W cannot be zero"));
                    return;
                }
                if (holeSize <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, LRPartsLocalizer.GetString(LRPartsResourceIDs.ErrInvalidF, "F cannot be zero or negative"));
                    return;
                }
                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                Matrix4X4 matrix = new Matrix4X4();
                //================================================= 
                //Construction of Physical Aspect 
                //=================================================

                //ports

                Port  Port1 = new Port(OccurrenceConnection, part, "Structure", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, -1));
                m_Symbolic.Outputs["Port1"] = Port1;
                Port Port2 = new Port(OccurrenceConnection, part, "Hole", new Position(0, 0, -(H + holeSize / 2)), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = Port2;

                Line3d line1;
                Collection<ICurve> curveCollection = new Collection<ICurve>();        

                line1 = new Line3d(new Position(-T / 2, W / 2, -H), new Position(-T / 2, W / 2, 0));
                curveCollection.Add(line1);
                line1 = new Line3d(new Position(-T / 2, W / 2, 0), new Position(-T / 2, -W / 2, 0));
                curveCollection.Add(line1);
                line1 = new Line3d(new Position(-T / 2, -W / 2, 0), new Position(-T / 2, -W / 2, -H));
                curveCollection.Add(line1);

                symbolGeometryHelper.ActivePosition = new Position(0, 0, 0);
                symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                Arc3d arc = symbolGeometryHelper.CreateArc(null, W / 2, Math.PI);
                matrix.Rotate(Math.PI, new Vector(0, 0, 1));
                matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0));
                matrix.Translate(new Vector(0, -T / 2, -H));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                arc.Transform(matrix);
                curveCollection.Add(arc);

                ComplexString3d lineString = new ComplexString3d(curveCollection);
                Vector lineVector = new Vector(1, 0, 0);
                Projection3d body = new Projection3d(lineString, lineVector, T, true);
                m_Symbolic.Outputs["BODY"] = body;

                Vector normal = new Position(T / 2, 0, -H).Subtract(new Position(-T / 2, 0, -H));
                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(-T / 2, 0, -H);
                symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 0, 1));
                Projection3d cylinder = (Projection3d)symbolGeometryHelper.CreateCylinder(null, holeSize / 2, normal.Length);
                m_Symbolic.Outputs["HOLE"] = cylinder;
            }
            catch  //General Unhandled exception 
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, LRPartsLocalizer.GetString(LRPartsResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of LRParts_FIG55NS.cs."));
                return;
            }
        }

        #endregion
    }
}

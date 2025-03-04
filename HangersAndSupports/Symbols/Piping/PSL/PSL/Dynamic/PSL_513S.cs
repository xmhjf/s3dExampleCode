//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2014R1, Intergraph Corporation. All rights reserved.
//
//   PSL_513S.cs
//    PSL,Ingr.SP3D.Content.Support.Symbols.PSL_513S
//   Author       :  Hema
//   Creation Date:  21-08-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   21-08-2013     Hema    CR-CP-232036,232037,232038,232039,232040 Convert HS_PSL VB Project to C# .Net    
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.ReferenceData.Middle;
using System.Collections.ObjectModel;

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
    public class PSL_513S : HangerComponentSymbolDefinition
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "PSL,Ingr.SP3D.Content.Support.Symbols.PSL_513S"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart PartInput;
        [InputDouble(2, "S", "S", 0.999999)]
        public InputDouble S;
        [InputDouble(3, "T", "T", 0.999999)]
        public InputDouble T;
        [InputDouble(4, "Q", "Q", 0.999999)]
        public InputDouble Q;
        [InputDouble(5, "C", "C", 0.999999)]
        public InputDouble C;
        [InputDouble(6, "O", "O", 0.999999)]
        public InputDouble O;
        [InputDouble(7, "PIPE_DIA", "PIPE_DIA", 0.999999)]
        public InputDouble PIPE_DIA;
        [InputDouble(8, "L", "L", 0.999999)]
        public InputDouble L;
        [InputDouble(9, "Z", "Z", 0.999999)]
        public InputDouble Z;
        [InputDouble(10, "D", "D", 0.999999)]
        public InputDouble D;
        [InputDouble(11, "E", "E", 0.999999)]
        public InputDouble E;
        [InputDouble(12, "F", "F", 0.999999)]
        public InputDouble F;
        [InputDouble(13, "I", "I", 0.999999)]
        public InputDouble I;
        [InputDouble(14, "U", "U", 0.999999)]
        public InputDouble U;
        [InputDouble(15, "R", "R", 0.999999)]
        public InputDouble R;

        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Symbolic Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("FILLER", "FILLER")]
        [SymbolOutput("BEND", "BEND")]
        [SymbolOutput("R", "R")]
        [SymbolOutput("L", "L")]
        [SymbolOutput("TOP", "TOP")]
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
                //-----------------------------------------------------------------------
                // Construction of Physical Aspect 
                //----------------------------------------------------------------------
                Part part = (Part)PartInput.Value;
                Double r = R.Value;
                Double s = S.Value;
                Double t = T.Value;
                Double q = Q.Value;
                Double c = C.Value;
                Double o = O.Value;
                Double pipeDiameter = PIPE_DIA.Value;
                Double l = L.Value;
                Double z = Z.Value;
                Double d = D.Value;
                Double e = E.Value;
                Double f = F.Value;
                Double i = I.Value;
                Double u = U.Value;
                Double rodDiameter = I.Value;

                //Intializing SymbolGeomHelper
                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                Matrix4X4 matrix = new Matrix4X4();
                Collection<ICurve> curveCollection = new Collection<ICurve>();

                //Ports
                Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;
                Port port2 = new Port(OccurrenceConnection, part, "Structure", new Position(0, 0, c - o), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;

                //Validating Inputs
                if (e == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidENEZ, "E value cannot be zero"));
                    return;
                }
                if (rodDiameter <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidRodDiameterGTZ, "ROD_DIA value should be greater than zero"));
                    return;
                }
                if (z == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidZNEZ, "Z value cannot be zero"));
                    return;
                }
                if (d == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidDNEZ, "D value cannot be zero"));
                    return;
                }
                curveCollection.Add(new Line3d(new Position(-e / 2.0, -u / 2.0 - o, 0), new Position(-e / 2.0, -pipeDiameter / 2.0, 0)));

                symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                Arc3d arc = symbolGeometryHelper.CreateArc(null, pipeDiameter / 2.0, Math.PI);
                matrix.Rotate(Math.PI, new Vector(0, 0, 1));
                matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0));
                matrix.Translate(new Vector(0, -e / 2.0, 0));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                arc.Transform(matrix);
                curveCollection.Add(arc);

                curveCollection.Add(new Line3d(new Position(-e / 2.0, pipeDiameter / 2.0, 0), new Position(-e / 2.0, u / 2.0 + o, 0)));

                Projection3d filler = new Projection3d(new ComplexString3d(curveCollection), new Vector(1, 0, 0), e, true);
                m_Symbolic.Outputs["FILLER"] = filler;

                Revolution3d bend = new Revolution3d((new Circle3d(new Position(u / 2.0, 0, 0), new Vector(0, 0, 1), rodDiameter / 2.0)), new Vector(0, -1, 0), new Position(0, 0, 0), Math.PI * 180 / 180, true);
                matrix = new Matrix4X4();
                matrix.Rotate(Math.PI, new Vector(1, 0, 0));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                bend.Transform(matrix);
                m_Symbolic.Outputs["BEND"] = bend;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 0, 1).GetOrthogonalVector());
                Projection3d rightCylinder = symbolGeometryHelper.CreateCylinder(null, rodDiameter / 2.0, z);
                matrix = new Matrix4X4();
                matrix.Translate(new Vector(u / 2.0, 0, 0));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                rightCylinder.Transform(matrix);
                m_Symbolic.Outputs["R"] = rightCylinder;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 0, 1).GetOrthogonalVector());
                Projection3d leftCylinder = symbolGeometryHelper.CreateCylinder(null, rodDiameter / 2.0, z);
                matrix = new Matrix4X4();
                matrix.Translate(new Vector(-u / 2.0, 0, 0));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                leftCylinder.Transform(matrix);
                m_Symbolic.Outputs["L"] = leftCylinder;

                Collection<Position> pointCollection = new Collection<Position>();

                pointCollection.Add(new Position(-d / 2.0, -f / 2.0, s));
                pointCollection.Add(new Position(-d / 2.0, -f / 2.0, s + t));
                pointCollection.Add(new Position(-d / 2.0, -r / 2.0, s + t));
                pointCollection.Add(new Position(-d / 2.0, -q / 2.0, c - o));
                pointCollection.Add(new Position(-d / 2.0, q / 2.0, c - o));
                pointCollection.Add(new Position(-d / 2.0, r / 2.0, s + t));
                pointCollection.Add(new Position(-d / 2.0, f / 2.0, s + t));
                pointCollection.Add(new Position(-d / 2.0, f / 2.0, s));
                pointCollection.Add(new Position(-d / 2.0, -f / 2.0, s));

                Projection3d top = new Projection3d(new LineString3d(pointCollection), new Vector(1, 0, 0), d, true);
                m_Symbolic.Outputs["TOP"] = top;
            }
            catch       //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of PSL_513S."));
                    return;
                }
            }
        }
        #endregion
    }
}

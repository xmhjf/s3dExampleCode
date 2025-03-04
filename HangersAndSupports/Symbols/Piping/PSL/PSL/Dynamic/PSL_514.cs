//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2014R1, Intergraph Corporation. All rights reserved.
//
//   PSL_514.cs
//    PSL,Ingr.SP3D.Content.Support.Symbols.PSL_514
//   Author       :  Rajeswari
//   Creation Date:  22-08-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   22-08-2013   Rajeswari  CR-CP-232036,232037,232038,232039,232040 Convert HS_PSL VB Project to C# .Net    
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Support.Middle;
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
    public class PSL_514 : HangerComponentSymbolDefinition
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "PSL,Ingr.SP3D.Content.Support.Symbols.PSL_514"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"
        [InputCatalogPart(1)]
        public InputCatalogPart PartInput;
        [InputDouble(2, "AJ", "AJ", 0.999999)]
        public InputDouble AJ;
        [InputDouble(3, "AH", "AH", 0.999999)]
        public InputDouble AH;
        [InputDouble(4, "A", "A", 0.999999)]
        public InputDouble A;
        [InputDouble(5, "C", "C", 0.999999)]
        public InputDouble C;
        [InputDouble(6, "PIPE_DIA", "PIPE_DIA", 0.999999)]
        public InputDouble PIPE_DIA;
        [InputDouble(7, "B", "B", 0.999999)]
        public InputDouble B;
        [InputDouble(8, "AK", "AK", 0.999999)]
        public InputDouble AK;
        [InputDouble(9, "D", "D", 0.999999)]
        public InputDouble D;
        [InputDouble(10, "E", "E", 0.999999)]
        public InputDouble E;
        [InputDouble(11, "F", "F", 0.999999)]
        public InputDouble F;
        [InputDouble(12, "CA", "CA", 0.999999)]
        public InputDouble CA;
        [InputDouble(13, "CB", "CB", 0.999999)]
        public InputDouble CB;
        [InputDouble(14, "CC", "CC", 0.999999)]
        public InputDouble CC;
        [InputDouble(15, "DB", "DB", 0.999999)]
        public InputDouble DB;
        [InputDouble(16, "AE", "AE", 0.999999)]
        public InputDouble AE;
        [InputDouble(17, "AP", "AP", 0.999999)]
        public InputDouble AP;
        [InputDouble(18, "AR", "AR", 0.999999)]
        public InputDouble AR;
        [InputDouble(19, "AF", "AF", 0.999999)]
        public InputDouble AF;
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
        [SymbolOutput("CYL1", "CYL1")]
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
                Part part = (Part)PartInput.Value;
                Double pipeDiameter = PIPE_DIA.Value;
                Double a = A.Value;
                Double b = B.Value;
                Double c = C.Value;
                Double d = D.Value;
                Double e = E.Value;
                Double f = F.Value;
                Double rodDiameter = CA.Value;
                Double cb = CB.Value;
                Double cc = CC.Value;
                Double db = DB.Value;
                Double ae = AE.Value;
                Double ap = AP.Value;
                Double ar = AR.Value;
                Double af = AF.Value;
                Double ak = AK.Value;
                Double aj = AJ.Value;
                Double ah = AH.Value;

                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                //ports

                Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;

                Port port2 = new Port(OccurrenceConnection, part, "Pin", new Position(0, 0, c - a / 2), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;

                if (pipeDiameter == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidpipeDiameterNEZ, "PIPE_DIA value cannot be zero"));
                    return;
                }
                if (e == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidENEZ, "E value cannot be zero"));
                    return;
                }
                if (rodDiameter <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidRodDiameterGTZ, "ROD_DIA should be greater than zero"));
                    return;
                }
                if (cc == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidCCNEZ, "CC value cannot be zero"));
                    return;
                }
                if (d == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidDNEZ, "D value cannot be zero"));
                    return;
                }
                Collection<ICurve> curveCollection = new Collection<ICurve>();
                Matrix4X4 matrix = new Matrix4X4();
                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();

                curveCollection.Add(new Line3d(new Position(-e / 2.0, -pipeDiameter / 2.0 - db, 0), new Position(-e / 2.0, -pipeDiameter / 2.0, 0)));
                symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                Arc3d arc = symbolGeometryHelper.CreateArc(null, pipeDiameter / 2.0, Math.PI);
                matrix.Rotate(Math.PI, new Vector(0, 0, 1));
                matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0));
                matrix.Translate(new Vector(0, -e / 2.0, 0));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                arc.Transform(matrix);
                curveCollection.Add(arc);
                curveCollection.Add(new Line3d(new Position(-e / 2.0, pipeDiameter / 2.0, 0), new Position(-e / 2.0, pipeDiameter / 2.0 + db, 0)));
                Projection3d filler = new Projection3d(new ComplexString3d(curveCollection), new Vector(1, 0, 0), e, false);
                m_Symbolic.Outputs["FILLER"] = filler;

                Revolution3d bend = new Revolution3d((new Circle3d(new Position(cb / 2.0, 0, 0), new Vector(0, 0, 1), rodDiameter / 2.0)), new Vector(0, -1, 0), new Position(0, 0, 0), (Math.Atan(1) * 4.0) * 180 / 180, true);
                matrix = new Matrix4X4();
                matrix.Rotate(Math.PI, new Vector(1, 0, 0));
                matrix.Translate(new Vector(0, 0, 0));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                bend.Transform(matrix);
                m_Symbolic.Outputs["BEND"] = bend;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 0, 1).GetOrthogonalVector());
                Projection3d r = symbolGeometryHelper.CreateCylinder(null, rodDiameter / 2.0, cc);
                matrix = new Matrix4X4();
                matrix.Translate(new Vector(cb / 2.0, 0, 0));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                r.Transform(matrix);
                m_Symbolic.Outputs["R"] = r;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 0, 1).GetOrthogonalVector());
                Projection3d l = symbolGeometryHelper.CreateCylinder(null, rodDiameter / 2.0, cc);
                matrix = new Matrix4X4();
                matrix.Translate(new Vector(-cb / 2.0, 0, 0));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                l.Transform(matrix);
                m_Symbolic.Outputs["L"] = l;

                Collection<Position> pointCollection = new Collection<Position>();
                pointCollection.Add(new Position(-d / 2.0, -f / 2.0, ae));
                pointCollection.Add(new Position(-d / 2.0, -f / 2.0, ae + af + ap + ar));
                pointCollection.Add(new Position(-d / 2.0, -ak / 2.0, ae + af + ap + ar));
                pointCollection.Add(new Position(-d / 2.0, -aj / 2.0, c + ah));
                pointCollection.Add(new Position(-d / 2.0, aj / 2.0, c + ah));
                pointCollection.Add(new Position(-d / 2.0, ak / 2.0, ae + af + ap + ar));
                pointCollection.Add(new Position(-d / 2.0, f / 2.0, ae + af + ap + ar));
                pointCollection.Add(new Position(-d / 2.0, f / 2.0, ae));
                pointCollection.Add(new Position(-d / 2.0, -f / 2.0, ae));

                Projection3d top = new Projection3d(new LineString3d(pointCollection), new Vector(1, 0, 0), d, true);
                m_Symbolic.Outputs["TOP"] = top;
            }
            catch       //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of PSL_514.cs"));
                    return;
                }
            }
        }

        #endregion

    }
}

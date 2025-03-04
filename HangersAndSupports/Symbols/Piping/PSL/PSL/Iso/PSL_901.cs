//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2014R1, Intergraph Corporation. All rights reserved.
//
//   PSL_901.cs
//    PSL,Ingr.SP3D.Content.Support.Symbols.PSL_901
//   Author       :  Rajeswari
//   Creation Date:  21-08-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   21-08-2013   Rajeswari  CR-CP-232036,232037,232038,232039,232040 Convert HS_PSL VB Project to C# .Net    
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
    public class PSL_901 : HangerComponentSymbolDefinition
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "PSL,Ingr.SP3D.Content.Support.Symbols.PSL_901"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"
        [InputCatalogPart(1)]
        public InputCatalogPart PartInput;
        [InputDouble(2, "GAGE", "GAGE", 0.999999)]
        public InputDouble GAGE;
        [InputDouble(3, "D", "D", 0.999999)]
        public InputDouble D;
        [InputDouble(4, "E", "E", 0.999999)]
        public InputDouble E;
        [InputDouble(5, "A", "A", 0.999999)]
        public InputDouble A;
        [InputDouble(6, "F", "F", 0.999999)]
        public InputDouble F;
        [InputDouble(7, "C", "C", 0.999999)]
        public InputDouble C;
        [InputDouble(8, "B", "B", 0.999999)]
        public InputDouble B;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Symbolic Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("BEND", "BEND")]
        [SymbolOutput("R", "R")]
        [SymbolOutput("L", "L")]
        [SymbolOutput("R2", "R2")]
        [SymbolOutput("L2", "L2")]
        [SymbolOutput("BASE", "BASE")]
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
                Double gage = GAGE.Value;
                Double b = B.Value;
                Double d = D.Value;
                Double e = E.Value;
                Double a = A.Value;
                Double f = F.Value;
                Double c = C.Value;

                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                Matrix4X4 matrix = new Matrix4X4();
                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                //ports
                Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;

                Port port2 = new Port(OccurrenceConnection, part, "Structure", new Position(0, 0, a / 2 + f), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;

                if (e <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidEGTZ, "E value should be greater than zero"));
                    return;
                }
                if (f <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidFGTZ, "F value should be greater than zero"));
                    return;
                }
                if (b <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidBGTZ, "B value should be greater than zero"));
                    return;
                }
                if (d == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidDNEZ, "D value cannot be zero"));
                    return;
                }
                if (a == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidANEZ, "A value cannot be zero"));
                    return;
                }
                if (c <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidCGTZ, "C value should be greater than zero"));
                    return;
                }
                Revolution3d bend = new Revolution3d((new Circle3d(new Position(c / 2, 0, 0), new Vector(0, 0, 1), c / 2 - a / 2)), new Vector(0, -1, 0), new Position(0, 0, 0), (Math.Atan(1) * 4) * 180 / 180, true);
                matrix.Rotate(Math.PI, new Vector(1, 0, 0));
                matrix.Translate(new Vector(0, gage, 0));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                bend.Transform(matrix);
                m_Symbolic.Outputs["BEND"] = bend;

                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 0, 1).GetOrthogonalVector());
                Projection3d r = symbolGeometryHelper.CreateCylinder(null, c / 2.0 - a / 2.0, a / 2.0);
                matrix = new Matrix4X4();
                matrix.Translate(new Vector(c / 2.0, gage, 0));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                r.Transform(matrix);
                m_Symbolic.Outputs["R"] = r;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 0, 1).GetOrthogonalVector());
                Projection3d l = symbolGeometryHelper.CreateCylinder(null, c / 2.0 - a / 2.0, a / 2.0);
                matrix = new Matrix4X4();
                matrix.Translate(new Vector(-c / 2.0, gage, 0));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                l.Transform(matrix);
                m_Symbolic.Outputs["L"] = l;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 0, 1).GetOrthogonalVector());
                Projection3d r2 = symbolGeometryHelper.CreateCylinder(null, b / 2.0, d);
                matrix = new Matrix4X4();
                matrix.Translate(new Vector(c / 2.0, gage, 0));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                r2.Transform(matrix);
                m_Symbolic.Outputs["R2"] = r2;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 0, 1).GetOrthogonalVector());
                Projection3d l2 = symbolGeometryHelper.CreateCylinder(null, b / 2.0, d);
                matrix = new Matrix4X4();
                matrix.Translate(new Vector(-c / 2.0, gage, 0));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                l2.Transform(matrix);
                m_Symbolic.Outputs["L2"] = l2;

                symbolGeometryHelper = new SymbolGeometryHelper();
                Projection3d baseBox = symbolGeometryHelper.CreateBox(null, c + b + (c / 2.0 - a / 2.0) * 2.0, e, f, 9);
                matrix = new Matrix4X4();
                matrix.Translate(new Vector(-(c + b + (c / 2.0 - a / 2.0) * 2.0) / 2.0, -e / 2.0 + gage, a / 2.0));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                baseBox.Transform(matrix);
                m_Symbolic.Outputs["BASE"] = baseBox;
            }
            catch       //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of PSL_901.cs"));
                    return;
                }
            }
        }
        #endregion

    }
}

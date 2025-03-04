//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2014R1, Intergraph Corporation. All rights reserved.
//
//   PSL_239.cs
//    PSL,Ingr.SP3D.Content.Support.Symbols.PSL_239
//   Author       :  Vijay
//   Creation Date:  23-08-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   23-08-2013     Vijay    CR-CP-232036,232037,232038,232039,232040 Convert HS_PSL VB Project to C# .Net    
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
    public class PSL_239 : HangerComponentSymbolDefinition
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "PSL,Ingr.SP3D.Content.Support.Symbols.PSL_239"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"
        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "ROD_DIA", "ROD_DIA", 0.999999)]
        public InputDouble ROD_DIA;
        [InputDouble(3, "A", "A", 0.999999)]
        public InputDouble A;
        [InputDouble(4, "B", "B", 0.999999)]
        public InputDouble B;
        [InputDouble(5, "D", "D", 0.999999)]
        public InputDouble D;
        [InputDouble(6, "C", "C", 0.999999)]
        public InputDouble C;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Symbolic Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("TOP", "TOP")]
        [SymbolOutput("TOP_CYL_1", "TOP_CYL_1")]
        [SymbolOutput("TOP_CYL_2", "TOP_CYL_2")]
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

                Double rodDiameter = ROD_DIA.Value;
                Double a = A.Value;
                Double b = B.Value;
                Double c = C.Value;
                Double d = D.Value;

                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                Matrix4X4 matrix = new Matrix4X4();

                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                //ports

                Port port1 = new Port(OccurrenceConnection, part, "InThdRH", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;

                Port port2 = new Port(OccurrenceConnection, part, "Eye", new Position(0, 0, rodDiameter / 2.0 - b), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;

                //Validating Inputs
                if (c <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidCGTZ, "C value should be greater than zero"));
                    return;
                }

                double angle = Math.Atan((a / 2.0 - rodDiameter / 2.0) / (d + b - (a / 2.0) - 4.0)) * 180.0 / Math.PI;
                double y2 = (a / 2.0 + c / 2.0) * Math.Cos(angle * Math.PI / 180);
                double y1 = rodDiameter / 2.0 + c - (y2 * (c / 2.0) / (a / 2.0 + c / 2.0));
                double z1 = (rodDiameter / 2.0) + d - (Math.Sqrt(((a / 2.0 + c / 2.0) * (a / 2.0 + c / 2.0)) - (y2 * y2)) * (c / 2.0) / (a / 2.0 + c / 2.0));
                double z2 = rodDiameter / 2.0 - b + a / 2.0 + Math.Sqrt((a / 2.0 + c / 2.0) * (a / 2.0 + c / 2.0) - y2 * y2);

                symbolGeometryHelper.ActivePosition = new Position(0, 0, rodDiameter / 2.0);
                Ellipse3d curve = (Ellipse3d)symbolGeometryHelper.CreateEllipse(null, (rodDiameter / 2.0 + c), c, 2 * Math.PI);
                Projection3d top = new Projection3d(curve, new Vector(0, 0, 1), d, true);
                m_Symbolic.Outputs["TOP"] = top;

                symbolGeometryHelper = new SymbolGeometryHelper();
                Vector normal = new Position(y2, 0, z2).Subtract(new Position(y1, 0, z1));
                symbolGeometryHelper.ActivePosition = new Position(y1, 0, z1);
                symbolGeometryHelper.SetOrientation(normal, normal.GetOrthogonalVector());
                Projection3d topCylinder1 = symbolGeometryHelper.CreateCylinder(null, c / 2.0, normal.Length);
                m_Symbolic.Outputs["TOP_CYL_1"] = topCylinder1;

                symbolGeometryHelper = new SymbolGeometryHelper();
                normal = new Position(-y2, 0, z2).Subtract(new Position(-y1, 0, z1));
                symbolGeometryHelper.ActivePosition = new Position(-y1, 0, z1);
                symbolGeometryHelper.SetOrientation(normal, normal.GetOrthogonalVector());
                Projection3d topCylinder2 = symbolGeometryHelper.CreateCylinder(null, c / 2.0, normal.Length);
                m_Symbolic.Outputs["TOP_CYL_2"] = topCylinder2;

                matrix.Rotate((180.0 - angle) * Math.PI / 180.0, new Vector(0, 1, 0));
                matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1));
                matrix.Translate(new Vector(0, 0, rodDiameter / 2.0 - b + a / 2.0));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                Revolution3d bend = new Revolution3d((new Circle3d(new Position(a / 2.0 + c / 2.0, 0, 0), new Vector(0, 0, 1), c / 2.0)), new Vector(0, -1, 0), new Position(0, 0, 0), (182.0 + (angle * 2.0)) * Math.PI / 180.0, true);
                bend.Transform(matrix);
                m_Symbolic.Outputs["BOTTOM"] = bend;
            }
            catch  //General Unhandled exception 
            {
                if (base.ToDoListMessage == null)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of PSL_239."));
                    return;
                }
            }
        }

        #endregion

    }
}

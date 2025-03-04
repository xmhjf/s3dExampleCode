//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2014R1, Intergraph Corporation. All rights reserved.
//
//   PSL_235.cs
//    PSL,Ingr.SP3D.Content.Support.Symbols.PSL_235
//   Author       :  Vijay
//   Creation Date:  22-08-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   22-08-2013     Vijay    CR-CP-232036,232037,232038,232039,232040 Convert HS_PSL VB Project to C# .Net    
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
    public class PSL_235 : HangerComponentSymbolDefinition
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "PSL,Ingr.SP3D.Content.Support.Symbols.PSL_235"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"
        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "ROD_DIA", "ROD_DIA", 0.999999)]
        public InputDouble ROD_DIA;
        [InputDouble(3, "B", "B", 0.999999)]
        public InputDouble B;
        [InputDouble(4, "C", "C", 0.999999)]
        public InputDouble C;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Symbolic Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("TOP", "TOP")]
        [SymbolOutput("BOTTOM", "BOTTOM")]
        [SymbolOutput("RIGHT", "RIGHT")]
        [SymbolOutput("LEFT", "LEFT")]
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

                Double b = B.Value;
                Double c = C.Value;
                Double rodDiameter = ROD_DIA.Value;
                Double opening = 0.152;
                Double rto = 0.075;

                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                Matrix4X4 matrix = new Matrix4X4();

                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                //ports

                Port port1 = new Port(OccurrenceConnection, part, "InThdRH", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;

                Port port2 = new Port(OccurrenceConnection, part, "InThdLH", new Position(0, 0, rto), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;

                //Validating Inputs
                if (rodDiameter <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidRodDiameterGTZ, "ROD_DIA value should be greater than zero"));
                    return;
                }

                Vector normal = new Position(0, 0, opening / 2.0 + rto / 2.0).Subtract(new Position(0, 0, opening / 2.0 + rto / 2.0 + b));
                symbolGeometryHelper.ActivePosition = new Position(0, 0, opening / 2.0 + rto / 2.0 + b);
                symbolGeometryHelper.SetOrientation(normal, normal.GetOrthogonalVector());
                Projection3d topCylinder = symbolGeometryHelper.CreateCylinder(null, (rodDiameter * 1.5) / (4.0 * Math.Cos((30 * Math.PI / 180))) + (rodDiameter * 1.5) * Math.Tan((30 * Math.PI / 180)) / 2.0, normal.Length);
                m_Symbolic.Outputs["TOP"] = topCylinder;

                symbolGeometryHelper = new SymbolGeometryHelper();
                normal = new Position(0, 0, -opening / 2.0 + rto / 2.0 - b).Subtract(new Position(0, 0, -opening / 2.0 + rto / 2.0));
                symbolGeometryHelper.ActivePosition = new Position(0, 0, -opening / 2.0 + rto / 2.0);
                symbolGeometryHelper.SetOrientation(normal, normal.GetOrthogonalVector());
                Projection3d bottomCylinder = symbolGeometryHelper.CreateCylinder(null, (rodDiameter * 1.5) / (4.0 * Math.Cos((30 * Math.PI / 180))) + (rodDiameter * 1.5) * Math.Tan((30 * Math.PI / 180)) / 2.0, normal.Length);
                m_Symbolic.Outputs["BOTTOM"] = bottomCylinder;

                symbolGeometryHelper = new SymbolGeometryHelper();
                matrix.Translate(new Vector(c * 0.2 + rodDiameter * 0.3, -0.6 * rodDiameter, -opening / 2.0 + rto / 2.0 - b));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                Projection3d rightBox = symbolGeometryHelper.CreateBox(null, (c - rodDiameter) * 0.3, 1.2 * rodDiameter, opening + b * 2.0, 9);
                rightBox.Transform(matrix);
                m_Symbolic.Outputs["RIGHT"] = rightBox;

                symbolGeometryHelper = new SymbolGeometryHelper();
                matrix = new Matrix4X4();
                matrix.Translate(new Vector(c * 0.2 + rodDiameter * 0.3, -0.6 * rodDiameter, -opening / 2.0 + rto / 2.0 - b));
                matrix.Rotate(Math.PI, new Vector(0, 0, 1), new Position(0, 0, 0));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                Projection3d leftBox = symbolGeometryHelper.CreateBox(null, (c - rodDiameter) * 0.3, 1.2 * rodDiameter, opening + b * 2.0, 9);
                leftBox.Transform(matrix);
                m_Symbolic.Outputs["LEFT"] = leftBox;
            }
            catch  //General Unhandled exception 
            {
                if (base.ToDoListMessage == null)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of PSL_235."));
                    return;
                }
            }
        }

        #endregion

    }
}

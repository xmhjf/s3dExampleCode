//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2014R1, Intergraph Corporation. All rights reserved.
//
//   PSL_SB.cs
//    PSL,Ingr.SP3D.Content.Support.Symbols.PSL_SB
//   Author       :  Vijaya
//   Creation Date:  23-08-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   23-08-2013     Vijaya    CR-CP-232036,232037,232038,232039,232040 Convert HS_PSL VB Project to C# .Net    
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
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
    public class PSL_SB : HangerComponentSymbolDefinition
    {

        //DefinitionName/ProgID of this symbol is "PSL_VAR,Ingr.SP3D.Content.Support.Symbols.PSL_SB"

        #region "Definition of Inputs"
        [InputCatalogPart(1)]
        public InputCatalogPart PartInput;
        [InputDouble(2, "F", "F", 0.999999)]
        public InputDouble F;
        [InputDouble(3, "E", "E", 0.999999)]
        public InputDouble E;
        [InputDouble(4, "C", "C", 0.999999)]
        public InputDouble C;
        [InputDouble(5, "B", "B", 0.999999)]
        public InputDouble B;
        [InputDouble(6, "G", "G", 0.999999)]
        public InputDouble G;
        [InputDouble(7, "L", "L", 0.999999)]
        public InputDouble L;
        [InputDouble(8, "H", "H", 0.999999)]
        public InputDouble H;
        [InputDouble(9, "ROD_DIA", "ROD_DIA", 0.999999)]
        public InputDouble ROD_DIA;
        [InputDouble(10, "A", "A", 0.999999)]
        public InputDouble A;
        [InputDouble(11, "D", "D", 0.999999)]
        public InputDouble D;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Symbolic Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("CYL1", "CYL1")]
        [SymbolOutput("CYL2", "CYL2")]
        [SymbolOutput("BOX1", "BOX1")]
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

                Double rodDiameter = ROD_DIA.Value, a = A.Value, b = B.Value, c = C.Value, d = D.Value, e = E.Value, f = F.Value, g = G.Value, l = L.Value, h = H.Value;

                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                Matrix4X4 matrix = new Matrix4X4();
                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                //ports
                Port port1 = new Port(OccurrenceConnection, part, "Hole", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;

                Port port2 = new Port(OccurrenceConnection, part, "InThdRH", new Position(0, 0, g + l + h / 2 + c - b), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;
                //Validating Inputs
                if (d <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidDGTZ, "D value should be greater than zero"));
                    return;
                }
                if (rodDiameter <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidRodDiameterGTZ, "ROD_DIA should be greater than zero"));
                    return;
                }
                if (a <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidAGTZ, "A value should be greater than zero"));
                    return;
                }
                Vector normal = new Position(0, 0, g + l + h / 2).Subtract(new Position(0, 0, g + h / 2));
                symbolGeometryHelper.ActivePosition = new Position(0, 0, g + h / 2);
                symbolGeometryHelper.SetOrientation(normal, normal.GetOrthogonalVector());
                Projection3d cylinder1 = symbolGeometryHelper.CreateCylinder(null, d / 2, normal.Length);
                m_Symbolic.Outputs["CYL1"] = cylinder1;

                symbolGeometryHelper = new SymbolGeometryHelper();
                normal = new Position(0, 0, g + l + c + rodDiameter / 2 + h / 2).Subtract(new Position(0, 0, l + g + h / 2));
                symbolGeometryHelper.ActivePosition = new Position(0, 0, l + g + h / 2);
                symbolGeometryHelper.SetOrientation(normal, normal.GetOrthogonalVector());
                Projection3d cylinder2 = symbolGeometryHelper.CreateCylinder(null, rodDiameter, normal.Length);
                m_Symbolic.Outputs["CYL2"] = cylinder2;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(-a / 2, -rodDiameter / 2, h / 2 - f);
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                Projection3d box1 = symbolGeometryHelper.CreateBox(null, a, rodDiameter, f + g, 9);
                box1.Transform(matrix);
                m_Symbolic.Outputs["BOX1"] = box1;
            }
            catch
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of PSL_SB.cs"));
                    return;
                }
            }
        }

        #endregion
    }
}

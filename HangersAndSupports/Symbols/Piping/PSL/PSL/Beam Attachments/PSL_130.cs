//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2014R1, Intergraph Corporation. All rights reserved.
//
//   PSL_130.cs
//    PSL,Ingr.SP3D.Content.Support.Symbols.PSL_130
//   Author       :  Vijay
//   Creation Date:  21-08-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   21-08-2013     Vijay    CR-CP-232036,232037,232038,232039,232040 Convert HS_PSL VB Project to C# .Net    
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
    public class PSL_130 : HangerComponentSymbolDefinition, ICustomHgrBOMDescription
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "PSL,Ingr.SP3D.Content.Support.Symbols.PSL_130"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "A", "A", 0.999999)]
        public InputDouble A;
        [InputDouble(3, "B", "B", 0.999999)]
        public InputDouble B;
        [InputDouble(4, "D", "D", 0.999999)]
        public InputDouble D;
        [InputDouble(5, "C", "C", 0.999999)]
        public InputDouble C;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Symbolic Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("BODY", "BODY")]
        [SymbolOutput("R_F_HOLE", "R_F_HOLE")]
        [SymbolOutput("R_B_HOLE", "R_B_HOLE")]
        [SymbolOutput("L_F_HOLE", "L_F_HOLE")]
        [SymbolOutput("L_B_HOLE", "L_B_HOLE")]
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

                Double holeSize = D.Value;
                Double t = C.Value;
                Double width = A.Value;
                Double b = B.Value;

                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                Matrix4X4 matrix = new Matrix4X4();

                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                //ports

                Port port1 = new Port(OccurrenceConnection, part, "Other", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;

                Port port2 = new Port(OccurrenceConnection, part, "Other", new Position(0, 0, t), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;

                //Validating Inputs
                if (t <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidCGTZ, "C value should be greater than zero"));
                    return;
                }
                if (width <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidAGTZ, "A value should be greater than zero"));
                    return;
                }
                if (holeSize <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidDGTZ, "D value should be greater than zero"));
                    return;
                }

                matrix.Rotate(Math.PI/2, new Vector(0, 1, 0), new Position(0, 0, 0));
                matrix.Translate(new Vector(-width / 2, -width / 2, t));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                Projection3d bodyBox = symbolGeometryHelper.CreateBox(null, t, width, width, 9);
                bodyBox.Transform(matrix);
                m_Symbolic.Outputs["BODY"] = bodyBox;

                matrix = new Matrix4X4();
                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 0, 1).GetOrthogonalVector());
                matrix.Translate(new Vector(b / 2, -b / 2, -0.001));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                Projection3d rfHoleCylinder = symbolGeometryHelper.CreateCylinder(null, holeSize / 2, t + 0.002);
                rfHoleCylinder.Transform(matrix);
                m_Symbolic.Outputs["R_F_HOLE"] = rfHoleCylinder;

                matrix = new Matrix4X4();
                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 0, 1).GetOrthogonalVector());
                matrix.Translate(new Vector(b / 2, b / 2, -0.001));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                Projection3d rbHoleCylinder = symbolGeometryHelper.CreateCylinder(null, holeSize / 2, t + 0.002);
                rbHoleCylinder.Transform(matrix);
                m_Symbolic.Outputs["R_B_HOLE"] = rbHoleCylinder;

                matrix = new Matrix4X4();
                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 0, 1).GetOrthogonalVector());
                matrix.Translate(new Vector(-b / 2, -b / 2, -0.001));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                Projection3d lfHoleCylinder = symbolGeometryHelper.CreateCylinder(null, holeSize / 2, t + 0.002);
                lfHoleCylinder.Transform(matrix);
                m_Symbolic.Outputs["L_F_HOLE"] = lfHoleCylinder;

                matrix = new Matrix4X4();
                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 0, 1).GetOrthogonalVector());
                matrix.Translate(new Vector(-b / 2, b / 2, -0.001));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                Projection3d lbHoleCylinder = symbolGeometryHelper.CreateCylinder(null, holeSize / 2, t + 0.002);
                lbHoleCylinder.Transform(matrix);
                m_Symbolic.Outputs["L_B_HOLE"] = lbHoleCylinder;
            }
            catch  //General Unhandled exception 
            {
                if (base.ToDoListMessage == null)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of PSL_130."));
                    return;
                }
            }
        }

        #endregion


        #region ICustomHgrBOMDescription Members

        public string BOMDescription(BusinessObject oSupportOrComponent)
        {
            string bomDescription = "";
            try
            {
                Part part = (Part)oSupportOrComponent.GetRelationship("madeFrom", "part").TargetObjects[0];
                bomDescription = "PSL " + (string)((PropertyValueString)part.GetPropertyValue("IJUAHgrPSL_PART_NUMBER", "PART_NUMBER")).PropValue + "  Ceiling Plate";
                return bomDescription;
            }
            catch
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrBOMDescription, "Error in BOMDescription of PSL_130.cs."));
                return "";
            }
        }

        #endregion
    }
}

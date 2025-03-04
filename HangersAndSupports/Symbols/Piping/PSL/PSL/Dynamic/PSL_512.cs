//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2014R1, Intergraph Corporation. All rights reserved.
//
//   PSL_512.cs
//    PSL,Ingr.SP3D.Content.Support.Symbols.PSL_512
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
    public class PSL_512 : HangerComponentSymbolDefinition
    {
        //DefinitionName/ProgID of this symbol is "PSL,Ingr.SP3D.Content.Support.Symbols.PSL_512"

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart PartInput;
        [InputDouble(2, "F", "F", 0.999999)]
        public InputDouble F;
        [InputDouble(3, "G", "G", 0.999999)]
        public InputDouble G;
        [InputDouble(4, "PIN_DIA", "PIN_DIA", 0.999999)]
        public InputDouble PIN_DIA;
        [InputDouble(5, "E", "E", 0.999999)]
        public InputDouble E;
        [InputDouble(6, "C", "C", 0.999999)]
        public InputDouble C;
        [InputDouble(7, "H", "H", 0.999999)]
        public InputDouble H;

        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Symbolic Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("TOP", "TOP")]
        [SymbolOutput("LEFT", "LEFT")]
        [SymbolOutput("RIGHT", "RIGHT")]
        [SymbolOutput("LEFT_CYL", "LEFT_CYL")]
        [SymbolOutput("RIGHT_CYL", "RIGHT_CYL")]
        [SymbolOutput("PIN", "PIN")]
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
                Double f = F.Value;
                Double g = G.Value;
                Double pinDiameter = PIN_DIA.Value;
                Double e = E.Value;
                Double c = C.Value;
                Double h = H.Value;
                Double t = (g - c) / 2.0;

                //Intializing SymbolGeomHelper
                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                Matrix4X4 matrix = new Matrix4X4();

                //Ports
                Port port1 = new Port(OccurrenceConnection, part, "Structure", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;
                Port port2 = new Port(OccurrenceConnection, part, "Pin", new Position(0, 0, e), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;

                //Validating Inputs
                if (g <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidGGTZ, "G value should be greater than zero"));
                    return;
                }
                if (h <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidHGTZ, "H value should be greater than zero"));
                    return;
                }
                if (t <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidTGTZ, "T value should be greater than zero"));
                    return;
                }
                if (f == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidFNEZ, "F value cannot be zero"));
                    return;
                }
                if (pinDiameter <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidPinDiameter, "PIN_DIA should be greater than zero"));
                    return;
                }
                if (e == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidENEZ, "E value cannot be zero"));
                    return;
                }
                symbolGeometryHelper.ActivePosition = new Position(-g / 2.0, -h / 2.0, 0);
                Projection3d top = symbolGeometryHelper.CreateBox(null, g, h, t, 9);
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                top.Transform(matrix);
                m_Symbolic.Outputs["TOP"] = top;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(g / 2.0 - t, -h / 2.0, t);
                Projection3d left = symbolGeometryHelper.CreateBox(null, t, h, e - t, 9);
                matrix = new Matrix4X4();
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                left.Transform(matrix);
                m_Symbolic.Outputs["LEFT"] = left;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(-g / 2.0, -h / 2.0, t);
                Projection3d right = symbolGeometryHelper.CreateBox(null, t, h, e - t, 9);
                matrix = new Matrix4X4();
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                right.Transform(matrix);
                m_Symbolic.Outputs["RIGHT"] = right;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 0, 1).GetOrthogonalVector());
                Projection3d leftCylinder = symbolGeometryHelper.CreateCylinder(null, f, t);
                matrix = new Matrix4X4();
                matrix.Rotate(Math.PI / 2, new Vector(0, 1, 0));
                matrix.Translate(new Vector(g / 2.0 - t, 0, e));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                leftCylinder.Transform(matrix);
                m_Symbolic.Outputs["LEFT_CYL"] = leftCylinder;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 0, 1).GetOrthogonalVector());
                Projection3d rightCylinder = symbolGeometryHelper.CreateCylinder(null, f, t);
                matrix = new Matrix4X4();
                matrix.Rotate(-Math.PI / 2, new Vector(0, 1, 0));
                matrix.Translate(new Vector(-g / 2.0 + t, 0, e));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                rightCylinder.Transform(matrix);
                m_Symbolic.Outputs["RIGHT_CYL"] = rightCylinder;

                symbolGeometryHelper = new SymbolGeometryHelper();
                Vector normal = new Position(0, -g / 2.0 - t, e).Subtract(new Position(0, g / 2.0 + t, e));
                symbolGeometryHelper.ActivePosition = new Position(0, g / 2.0 + t, e);
                symbolGeometryHelper.SetOrientation(normal, normal.GetOrthogonalVector());
                Projection3d pin = symbolGeometryHelper.CreateCylinder(null, pinDiameter / 2.0, normal.Length);
                m_Symbolic.Outputs["PIN"] = pin;
            }
            catch       //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of PSL_512."));
                    return;
                }
            }
        }
        #endregion
    }
}

//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   SFS5390.cs
//   FINL_Parts,Ingr.SP3D.Content.Support.Symbols.SFS5390
//   Author       :  Hema
//   Creation Date:  18-03-2013
//   Description:    Converted FINL_Parts VB Project to C#.Net Project 

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   18-03-2013      Hema    Converted FINL_Parts VB Project to C#.Net Project 
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
    [SymbolVersion("1.0.0.0")]
    [CacheOption(CacheOptionType.Cached)]
    public class SFS5390 : HangerComponentSymbolDefinition
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "FINL_Parts,Ingr.SP3D.Content.Support.Symbols.SFS5390"
        //----------------------------------------------------------------------------------


        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "ROD_DIA", "ROD_DIA", 0.999999)]
        public InputDouble m_dROD_DIA;
        [InputDouble(3, "A", "A", 0.999999)]
        public InputDouble m_dA;
        [InputDouble(4, "B", "B", 0.999999)]
        public InputDouble m_dB;
        [InputDouble(5, "C", "C", 0.999999)]
        public InputDouble m_dC;
        [InputDouble(6, "D", "D", 0.999999)]
        public InputDouble m_dD;
        [InputDouble(7, "G", "G", 0.999999)]
        public InputDouble m_dG;
        [InputDouble(8, "H", "H", 0.999999)]
        public InputDouble m_dH;
        [InputDouble(9, "L", "L", 0.999999)]
        public InputDouble m_dL;
        [InputDouble(10, "M", "M", 0.999999)]
        public InputDouble m_dM;
        [InputDouble(11, "N", "N", 0.999999)]
        public InputDouble m_dN;
        [InputDouble(12, "S", "S", 0.999999)]
        public InputDouble m_dS;
        [InputDouble(13, "T", "T", 0.999999)]
        public InputDouble m_dT;
        [InputDouble(14, "X", "X", 0.999999)]
        public InputDouble m_dX;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("Left", "Left")]
        [SymbolOutput("Right", "Right")]
        [SymbolOutput("LeftCyl", "LeftCyl")]
        [SymbolOutput("RightCyl", "RightCyl")]
        [SymbolOutput("Pin", "Pin")]
        [SymbolOutput("Bottom", "Bottom")]
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
                 
               
               
                Double rod_Diameter = m_dROD_DIA.Value;
                Double A = m_dA.Value;
                Double B = m_dB.Value;
                Double C = m_dC.Value;
                Double D = m_dD.Value;
                Double G = m_dG.Value;
                Double H = m_dH.Value;
                Double L = m_dL.Value;
                Double M = m_dM.Value;
                Double N = m_dN.Value;
                Double S = m_dS.Value;
                Double T = m_dT.Value;
                Double X = m_dX.Value;
                Matrix4X4 matrix = new Matrix4X4();

                //ports
                Port port1 = new Port(OccurrenceConnection, part, "InThdRH", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;
                Port port2 = new Port(OccurrenceConnection, part, "Pin", new Position(0, 0, -X + A), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;

                if (H <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, FINL_PartsLocalizer.GetString(FINL_PartsSymbolResourceIDs.ErrInvalidHGTZero, "H value should be greater than zero"));
                    return;
                }
                if (N <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, FINL_PartsLocalizer.GetString(FINL_PartsSymbolResourceIDs.ErrInvalidNGTZero, "N value should be greater than zero"));
                    return;
                }
                if (A <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, FINL_PartsLocalizer.GetString(FINL_PartsSymbolResourceIDs.ErrInvalidAGTZero, "A value should be greater than zero"));
                    return;
                }
                if (M <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, FINL_PartsLocalizer.GetString(FINL_PartsSymbolResourceIDs.ErrInvalidMGTZero, "M value should be greater than zero"));
                    return;
                }
                if (T == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, FINL_PartsLocalizer.GetString(FINL_PartsSymbolResourceIDs.ErrInvalidTNZero, "T value should be greater than zero"));
                    return;
                }
                if (D <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, FINL_PartsLocalizer.GetString(FINL_PartsSymbolResourceIDs.ErrInvalidDGTZero, "D value should be greater than zero"));
                    return;
                }
                if (B == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, FINL_PartsLocalizer.GetString(FINL_PartsSymbolResourceIDs.ErrInvalidBNZero, "B value cannot be zero"));
                    return;
                }
                if (C <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, FINL_PartsLocalizer.GetString(FINL_PartsSymbolResourceIDs.ErrInvalidCGTZero, "C value should be greater than zero"));
                    return;
                }

                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                
                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(-G / 2 - H, -N / 2, -X);
                Projection3d left = (Projection3d)symbolGeometryHelper.CreateBox(null, H, N, A, 9);
                m_Symbolic.Outputs["Left"] = left;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(G / 2, -N / 2, -X);
                Projection3d right = (Projection3d)symbolGeometryHelper.CreateBox(null, H, N, A, 9);
                m_Symbolic.Outputs["Right"] = right;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0,0,0);
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 0, 1).GetOrthogonalVector());
                Projection3d leftCylinder = (Projection3d)symbolGeometryHelper.CreateCylinder(null, M / 2, T);
                matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0));
                matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1));
                matrix.Translate(new Vector(-S / 2 - T, 0, -X + A));
                leftCylinder.Transform(matrix);
                m_Symbolic.Outputs["LeftCyl"] = leftCylinder;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0,0,0);
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 0, 1).GetOrthogonalVector());
                Projection3d rightCylinder = (Projection3d)symbolGeometryHelper.CreateCylinder(null, M / 2, T);
                matrix = new Matrix4X4();
                matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0));
                matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1));
                matrix.Translate(new Vector(S / 2, 0, -X + A));
                rightCylinder.Transform(matrix);
                m_Symbolic.Outputs["RightCyl"] = rightCylinder;

                symbolGeometryHelper = new SymbolGeometryHelper();
                Vector normal = new Position(L / 2, 0, -X + A).Subtract(new Position(-L / 2, 0, -X + A));
                symbolGeometryHelper.ActivePosition = new Position(-L / 2, 0, -X + A);
                symbolGeometryHelper.SetOrientation(normal, normal.GetOrthogonalVector());
                Projection3d pin = (Projection3d)symbolGeometryHelper.CreateCylinder(null, D / 2, normal.Length);
                m_Symbolic.Outputs["Pin"] = pin;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, 0, -X);
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 0, 1).GetOrthogonalVector());
                Projection3d bottom = (Projection3d)symbolGeometryHelper.CreateCylinder(null, C / 2, B);
                m_Symbolic.Outputs["Bottom"] = bottom;
            }
            catch       //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, FINL_PartsLocalizer.GetString(FINL_PartsSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of SFS5390."));
                    return;
                }
            }
        }
        #endregion
    }
}

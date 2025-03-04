//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   SFS5389.cs
//   FINL_Parts,Ingr.SP3D.Content.Support.Symbols.SFS5389
//   Author       :  Hema
//   Creation Date:  18-03-2013 
//   Description:    Converted FINL_Parts VB Project to C#.Net Project 

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   18-03-2013      Hema    Converted FINL_Parts VB Project to C#.Net Project 
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using System.Collections.ObjectModel;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
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
    [SymbolVersion("1.0.0.0")]
    [CacheOption(CacheOptionType.Cached)]
    public class SFS5389 : HangerComponentSymbolDefinition
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "FINL_Parts,Ingr.SP3D.Content.Support.Symbols.SFS5389"
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
        [InputDouble(9, "S", "S", 0.999999)]
        public InputDouble m_dS;
        [InputDouble(10, "X", "X", 0.999999)]
        public InputDouble m_dX;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("Cyl", "Cyl")]
        [SymbolOutput("Body", "Body")]
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
                Double S = m_dS.Value;
                Double X = m_dX.Value;

                //ports
                Port port1 = new Port(OccurrenceConnection, part, "InThdRH", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;
                Port port2 = new Port(OccurrenceConnection, part, "Eye", new Position(0, 0, X - B - C), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;


                if (S == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, FINL_PartsLocalizer.GetString(FINL_PartsSymbolResourceIDs.ErrInvalidSNEZ, "S value cannot be zero"));
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

                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                
                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, 0, X - B);
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(1, 0, 0));
                Projection3d cylinder = (Projection3d)symbolGeometryHelper.CreateCylinder(null, D / 2, B);
                m_Symbolic.Outputs["Cyl"] = cylinder;

                Collection<ICurve> curveCollection = new Collection<ICurve>();
                Matrix4X4 matrix = new Matrix4X4();

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, 0, 0);
                symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                Arc3d arc = symbolGeometryHelper.CreateArc(null, H + G / 2, Math.PI);
                matrix.Rotate(Math.PI, new Vector(0, 0, 1));
                matrix.Rotate(3 * (Math.PI / 2), new Vector(0, 0, 1));
                matrix.Rotate(-Math.PI / 2, new Vector(0, 1, 0));
                matrix.Translate(new Vector(-S / 2, 0, X - B - C));
                arc.Transform(matrix);
                curveCollection.Add(arc);

                curveCollection.Add(new Line3d(new Position(-S / 2, -G / 2 - H, -(-X + B + C)), new Position(-S / 2, -G / 2 - H, X)));
                curveCollection.Add(new Line3d(new Position(-S / 2, -G / 2 - H, X), new Position(-S / 2, -G / 2, X)));
                curveCollection.Add(new Line3d(new Position(-S / 2, -G / 2, X), new Position(-S / 2, -G / 2, -(-X + B + C))));

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, 0, 0);
                symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                Arc3d arc1 = symbolGeometryHelper.CreateArc(null, G / 2, Math.PI);
                matrix = new Matrix4X4();
                matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1));
                matrix.Rotate(-Math.PI / 2, new Vector(0, 1, 0));
                matrix.Translate(new Vector(-S / 2, 0, X - B - C));
                arc1.Transform(matrix);
                curveCollection.Add(arc1);

                curveCollection.Add(new Line3d(new Position(-S / 2, G / 2, -(-X + B + C)), new Position(-S / 2, G / 2, X)));
                curveCollection.Add(new Line3d(new Position(-S / 2, G / 2, X), new Position(-S / 2, G / 2 + H, X)));
                curveCollection.Add(new Line3d(new Position(-S / 2, G / 2 + H, X), new Position(-S / 2, G / 2 + H, -(-X + B + C))));

                Projection3d body = new Projection3d(new ComplexString3d(curveCollection), new Vector(1, 0, 0), S, true);
                m_Symbolic.Outputs["Body"] = body;
            }
            catch       //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, FINL_PartsLocalizer.GetString(FINL_PartsSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of SFS5389."));
                    return;
                }
            }
        }
        #endregion
    }
}

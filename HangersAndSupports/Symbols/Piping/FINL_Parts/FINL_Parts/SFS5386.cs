//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   SFS5386.cs
//   FINL_Parts,Ingr.SP3D.Content.Support.Symbols.SFS5386
//   Author       :  Hema
//   Creation Date:  18-03-2013 
//   Description:    Converted FINL_Parts VB Project to C#.Net Project 

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   18-03-2013      Hema    Converted FINL_Parts VB Project to C#.Net Project 
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
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
    public class SFS5386 : HangerComponentSymbolDefinition
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "FINL_Parts,Ingr.SP3D.Content.Support.Symbols.SFS5386"
        //----------------------------------------------------------------------------------


        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "L", "L", 0.999999)]
        public InputDouble m_dL;
        [InputDouble(3, "Thickness", "Thickness", 0.999999)]
        public InputDouble m_dThickness;
        [InputDouble(4, "PlateL", "PlateL", 0.999999)]
        public InputDouble m_dPlateL;
        [InputDouble(5, "HoleDia", "HoleDia", 0.999999)]
        public InputDouble m_dHoleDia;
        [InputDouble(6, "Width", "Width", 0.999999)]
        public InputDouble m_dWidth;
        [InputDouble(7, "Radius", "Radius", 0.999999)]
        public InputDouble m_dRadius;
        [InputDouble(8, "TakeOut", "TakeOut", 0.999999)]
        public InputDouble m_dTakeOut;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("LeftPlate", "LeftPlate")]
        [SymbolOutput("RightPlate", "RightPlate")]
        [SymbolOutput("BottPlate", "BottPlate")]
        [SymbolOutput("Hole", "Hole")]
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
                 
               
               
                Double length = m_dL.Value;
                Double thickness = m_dThickness.Value;
                Double plateL = m_dPlateL.Value;
                Double holeDiameter = m_dHoleDia.Value;
                Double width = m_dWidth.Value;
                Double radius = m_dRadius.Value;
                Double takeOut = m_dTakeOut.Value;
                Double height = 0.12;
                Matrix4X4 matrix = new Matrix4X4();

                //ports
                Port port1 = new Port(OccurrenceConnection, part, "Structure", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;
                Port port2 = new Port(OccurrenceConnection, part, "Hole", new Position(0, 0, -(height - takeOut)), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;
               
                if (length <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, FINL_PartsLocalizer.GetString(FINL_PartsSymbolResourceIDs.ErrInvalidLengthGTZero, "Length should be greater than zero"));
                    return;
                }
                if (thickness == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, FINL_PartsLocalizer.GetString(FINL_PartsSymbolResourceIDs.ErrInvalidThickness, "Thickness cannot be zero"));
                    return;
                }
                if (holeDiameter <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, FINL_PartsLocalizer.GetString(FINL_PartsSymbolResourceIDs.ErrInvalidHoleDiameterGTZero, "HoleDia value should be greater than zero"));
                    return;
                }
                if (height <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, FINL_PartsLocalizer.GetString(FINL_PartsSymbolResourceIDs.ErrInvalidHeightGTZero, "Height should be greater than zero"));
                    return;
                }

                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                
                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(-width / 2 - thickness, -length / 2, height);
                Projection3d bottomPlate = (Projection3d)symbolGeometryHelper.CreateBox(null, width + 2 * thickness, length, thickness, 9);
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                matrix.Rotate(Math.PI, new Vector(1, 0, 0));
                bottomPlate.Transform(matrix);
                m_Symbolic.Outputs["BottPlate"] = bottomPlate;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(-length / 2, -height, width / 2);
                Projection3d leftPlate = (Projection3d)symbolGeometryHelper.CreateBox(null, length, height, thickness, 9);
                matrix = new Matrix4X4();
                matrix.Rotate(Math.PI/2, new Vector(1, 0, 0));
                leftPlate.Transform(matrix);
                m_Symbolic.Outputs["LeftPlate"] = leftPlate;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(-length / 2, -height, -width / 2 - thickness);
                Projection3d rightPlate = (Projection3d)symbolGeometryHelper.CreateBox(null, length, height, thickness, 9);
                matrix = new Matrix4X4();
                matrix.Rotate(Math.PI/2, new Vector(1, 0, 0));
                rightPlate.Transform(matrix);
                m_Symbolic.Outputs["RightPlate"] = rightPlate;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, 0, -height - thickness);
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), (new Vector(0, 0, 1)).GetOrthogonalVector());
                Projection3d hole = (Projection3d)symbolGeometryHelper.CreateCylinder(null, holeDiameter / 2, thickness);
                matrix = new Matrix4X4();
                matrix.Rotate(3*Math.PI / 2, new Vector(0, 0, 1));
                hole.Transform(matrix);
                m_Symbolic.Outputs["Hole"] = hole;
            }
            catch       //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, FINL_PartsLocalizer.GetString(FINL_PartsSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of SFS5386."));
                    return;
                }
            }
        }
        #endregion
    }
}

//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   StTrayCirRail.cs
//    TrayShip,Ingr.SP3D.Content.Support.Symbols.StTrayCirRail
//    Author       :  Manikanth
//    Creation Date:  16-01-2013
//    Description:   CR-CP-222297 .Net Hs_TrayShip Creation 


//dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   16-01-2013  Manikanth   CR-CP-222297 .Net Hs_TrayShip Creation 
//	 22/03/2013		Vijay 	 DI-CP-228142  Modify the error handling for delivered H&S symbols
//   04/June/2013   Manikanth TR-CP-234520 Implemented TDL Issues
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.Support.Middle;
using System.Collections.ObjectModel;

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
    [CacheOption(CacheOptionType.NonCached)]
    [SymbolVersion("1.0.0.0")]
    [VariableOutputs]
    public class TeeTray : HangerComponentSymbolDefinition
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "TrayShip,Ingr.SP3D.Content.Support.Symbols.TeeTray"
        //----------------------------------------------------------------------------------


        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "Radius", "Radius", 0.999999)]
        public InputDouble m_Radius;
        [InputDouble(3, "Thickness", "Thickness", 0.999999)]
        public InputDouble m_Thickness;
        [InputDouble(4, "Width", "Width", 0.999999)]
        public InputDouble m_Width;
        [InputDouble(5, "Width2", "Width2", 0.999999)]
        public InputDouble m_Width2;
        [InputDouble(6, "Length", "Length", 0.999999)]
        public InputDouble m_Length;
        [InputDouble(7, "BeginAngle", "BeginAngle", 0.999999)]
        public InputDouble m_BeginAngle;
        [InputDouble(8, "EndAngle", "EndAngle", 0.999999)]
        public InputDouble m_EndAngle;
        [InputString(9, "BOM_DESC", "BOM_DESC", "")]
        public InputString m_BOM_DESC;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("BODY", "BODY")]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("Port3", "Port3")]
        [SymbolOutput("Port4", "Port4")]
        [SymbolOutput("Port5", "Port5")]
        [SymbolOutput("Port6", "Port6")]
        [SymbolOutput("Port7", "Port7")]
        [SymbolOutput("CYLINDER", "CYLINDER")]
        [SymbolOutput("LEFTBEND", "LEFTBEND")]
        [SymbolOutput("RIGHTBEND", "RIGHTBEND")]
        [SymbolOutput("CYLINDER1", "CYLINDER1")]
        [SymbolOutput("CYLINDER2", "CYLINDER2")]
        [SymbolOutput("LEGS1_", "LEGS1_")]
        [SymbolOutput("LEGS2_", "LEGS2_")]
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


                double radius = m_Radius.Value;
                double thickness = m_Thickness.Value;
                double width = m_Width.Value;
                double width1 = m_Width2.Value;
                double length = m_Length.Value;
                double beginAngle = m_BeginAngle.Value;
                double endAngle = m_EndAngle.Value;
                double spacing = 0.3;
                double leftOffet;

                if (width == 0 || width1 == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, TrayShipLocalizer.GetString(TrayShiplResourceIDs.ErrWidthETZArguments, "Width cannot be zero"));
                    return;
                }
                if (thickness == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, TrayShipLocalizer.GetString(TrayShiplResourceIDs.ErrNETZeroThicknessArguments, "Thickness cannot be zero"));
                    return;
                }
                if (radius <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, TrayShipLocalizer.GetString(TrayShiplResourceIDs.ErrInvalidRadius, "Radius should be greater than zero"));
                    return;
                }
                if (width1 + 2 * radius > length)
                {
                    length = width1 + 2 * radius;
                    length = Math.Round(length, 2);
                }

                //ports

                Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;
                Port port2 = new Port(OccurrenceConnection, part, "CableTray1", new Position(-length / 2, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;
                Port port3 = new Port(OccurrenceConnection, part, "CableTray2", new Position(length / 2, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port3"] = port3;

                int noOfRungH = (int)(length / spacing) + 1;
                leftOffet = (length - noOfRungH - 1) * spacing / 2;
                int noOfRungV = (int)(radius / spacing) + 1;


                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                Collection<ICurve> curveCollection = new Collection<ICurve>();
                Matrix4X4 matrix = new Matrix4X4();
                Line3d line;
                line = new Line3d(new Position(-length / 2, -width / 2, 0), new Position(length / 2, -width / 2, 0));
                curveCollection.Add(line);
                line = new Line3d(new Position(-length / 2, -width / 2, 0), new Position(-length / 2, width / 2, 0));
                curveCollection.Add(line);
                if (length > width1 + 2 * radius)
                {
                    line = new Line3d(new Position(-length / 2, width / 2, 0), new Position(-(width1 / 2 + radius), width / 2, 0));
                    curveCollection.Add(line);

                    line = new Line3d(new Position(length / 2, width / 2, 0), new Position(width1 / 2 + radius, width / 2, 0));
                    curveCollection.Add(line);
                }
                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, 0, 0);
                symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                matrix.SetIdentity();
                matrix.Rotate(-Math.PI / 2, new Vector(0, 0, 1));
                Arc3d arc1 = symbolGeometryHelper.CreateArc(null, radius, Math.PI / 2);
                arc1.Transform(matrix);
                matrix = new Matrix4X4();
                matrix.SetIdentity();
                matrix.Translate(new Vector(-(width1 / 2 + radius), width / 2 + radius, 0));
                arc1.Transform(matrix);
                curveCollection.Add(arc1);

                line = new Line3d(new Position(-width1 / 2, width / 2 + radius, 0), new Position(width1 / 2, width / 2 + radius, 0));
                curveCollection.Add(line);

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, 0, 0);
                symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                matrix.SetIdentity();
                matrix.Rotate(Math.PI, new Vector(0, 0, 1));
                Arc3d arc2 = symbolGeometryHelper.CreateArc(null, radius, Math.PI / 2);
                arc2.Transform(matrix);
                matrix = new Matrix4X4();
                matrix.SetIdentity();
                matrix.Translate(new Vector(width1 / 2 + radius, width / 2 + radius, 0));
                arc2.Transform(matrix);
                curveCollection.Add(arc2);

                line = new Line3d(new Position(length / 2, width / 2, 0), new Position(length / 2, -width / 2, 0));
                curveCollection.Add(line);
                ComplexString3d complexString = new ComplexString3d(curveCollection);
                Projection3d body = new Projection3d(complexString, new Vector(0, 0, thickness), thickness, true);
                m_Symbolic.Outputs["BODY"] = body;
            }
            catch //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, TrayShipLocalizer.GetString(TrayShiplResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of TeeTray.cs."));
                    return;
                }
            }
        }
        #endregion

    }

}

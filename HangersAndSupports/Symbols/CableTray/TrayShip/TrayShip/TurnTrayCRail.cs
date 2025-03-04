//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   StTrayCirRail.cs
//    TrayShip,Ingr.SP3D.Content.Support.Symbols.StTrayCirRail
//    Author       :  Manikanth
//    Creation Date:  19-01-2013
//    Description:   CR-CP-222297 .Net Hs_TrayShip Creation 


//dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   19-01-2013  Manikanth   CR-CP-222297 .Net Hs_TrayShip Creation 
//	 22/03/2013		Vijay 	 DI-CP-228142  Modify the error handling for delivered H&S symbols
//   04/June/2013   Manikanth TR-CP-234520 Implemented TDL Issues
//   11-Dec-2014     PVK     TR-CP-264951	Resolve P3 coverity issues found in November 2014 report
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
    public class TurnTrayCRail : HangerComponentSymbolDefinition
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "TrayShip,Ingr.SP3D.Content.Support.Symbols.TurnTrayCRail"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "InsideRadius", "InsideRadius", 0.999999)]
        public InputDouble m_InsideRadius;
        [InputDouble(3, "Thickness", "Thickness", 0.999999)]
        public InputDouble m_Thickness;
        [InputDouble(4, "Width", "Width", 0.999999)]
        public InputDouble m_Width;
        [InputDouble(5, "Angle", "Angle", 0.999999)]
        public InputDouble m_dAngle;
        [InputDouble(6, "BeginAngle", "BeginAngle", 0.999999)]
        public InputDouble m_BeginAngle;
        [InputDouble(7, "EndAngle", "EndAngle", 0.999999)]
        public InputDouble m_EndAngle;
        [InputDouble(8, "D", "D", 0.999999)]
        public InputDouble m_D;
        [InputDouble(9, "PlateType", "PlateType", 1)]
        public InputDouble m_PlateType;
        [InputString(10, "BOM_DESC", "BOM_DESC", "")]
        public InputString m_BOM_DESC;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("Port3", "Port3")]
        [SymbolOutput("Port4", "Port4")]
        [SymbolOutput("BODY", "BODY")]
        [SymbolOutput("TOPBEND", "TOPBEND")]
        [SymbolOutput("BOTTOMBEND", "BOTTOMBEND")]
        [SymbolOutput("LEGS_", "LEGS_")]
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
                double radius = m_InsideRadius.Value;
                double thickness = m_Thickness.Value;
                double width = m_Width.Value;
                double angle = m_dAngle.Value;
                double beginAngle = m_BeginAngle.Value;
                double endAngle = m_EndAngle.Value;
                double d = m_D.Value;
                int iNoofrung;
                double spacing = 0.3;
                const double z = 0.005;
                int pType = (int)m_PlateType.Value;

                if (HgrCompareDoubleService.cmpdbl(angle % (2 * Math.PI) , 0)==true)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, TrayShipLocalizer.GetString(TrayShiplResourceIDs.ErrAnglesArguments, "Angle cannot be zero and multiple of 360°"));
                    return;
                }
                if (width <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, TrayShipLocalizer.GetString(TrayShiplResourceIDs.ErrNETZeroWidthArguments, "Width should be greater than zero"));
                    return;
                }
                if (thickness <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, TrayShipLocalizer.GetString(TrayShiplResourceIDs.ErrThicknessArguments, "Thickness should be greater than zero"));
                    return;
                }
                if (d <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, TrayShipLocalizer.GetString(TrayShiplResourceIDs.ErrDArguments, "Circular Coil Diameter should be greater than zero"));
                    return;
                }
                if (radius <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, TrayShipLocalizer.GetString(TrayShiplResourceIDs.ErrInvalidRadius, "Radius should be greater than zero"));
                    return;
                }
                //ports

                Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;
                Port port2 = new Port(OccurrenceConnection, part, "CableTray1", new Position((radius + width / 2) * Math.Cos(beginAngle), (radius + width / 2) * Math.Sin(beginAngle), 0), new Vector(Math.Cos(beginAngle), Math.Sin(beginAngle), 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;
                Port port3 = new Port(OccurrenceConnection, part, "CableTray2", new Position((radius + width / 2) * Math.Cos(angle - endAngle), (radius + width / 2) * Math.Sin(angle - endAngle), 0), new Vector(Math.Cos(angle - endAngle), Math.Sin(angle - endAngle), 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port3"] = port3;
                Port port4 = new Port(OccurrenceConnection, part, "CableTray3", new Position((radius + width / 2.0) * Math.Cos(angle / 2), (radius + width / 2) * Math.Sin(angle / 2), 0), new Vector(Math.Cos(angle / 2), Math.Sin(angle / 2), 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port4"] = port4;

                double newAngle;
                newAngle = Math.PI + angle;

                Revolution3d bend = new Revolution3d(OccurrenceConnection, (new Circle3d(new Position(0, (radius + (d / 2)), thickness / 2), new Vector(1, 0, 0), d / 2)), new Vector(0, 0, -1), new Position(0, 0, thickness / 2), angle, true);
                m_Symbolic.Outputs["TOPBEND"] = bend;

                Revolution3d bend1 = new Revolution3d(OccurrenceConnection, (new Circle3d(new Position(0, radius + width - d / 2, thickness / 2), new Vector(1, 0, 0), d / 2)), new Vector(0, 0, -1), new Position(0, 0, thickness / 2), angle, true);
                m_Symbolic.Outputs["BOTTOMBEND"] = bend1;

                Line3d line;
                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                Collection<ICurve> curveCollection = new Collection<ICurve>();
                Matrix4X4 matrix = new Matrix4X4();
                iNoofrung = (int)(angle / (spacing / (radius + width / 2))) + 1;
                if (iNoofrung == 0)
                {
                    iNoofrung = 1;
                    beginAngle = angle / 2;
                }
                else
                {
                    beginAngle = (angle - (spacing / (radius + width / 2)) * (iNoofrung - 1)) / 2;
                }
                if (pType == 1)
                {

                    line = new Line3d(new Position(radius + d, 0, 0), new Position(radius - d + width, 0, 0));
                    curveCollection.Add(line);

                    symbolGeometryHelper = new SymbolGeometryHelper();
                    symbolGeometryHelper.ActivePosition = new Position(0, 0, 0);
                    symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                    Arc3d outerArc = symbolGeometryHelper.CreateArc(null, radius + d, angle);
                    outerArc.Transform(matrix);
                    matrix = new Matrix4X4();
                    matrix.SetIdentity();
                    matrix.Translate(new Vector(0, 0, 0));
                    outerArc.Transform(matrix);
                    curveCollection.Add(outerArc);

                    symbolGeometryHelper = new SymbolGeometryHelper();
                    symbolGeometryHelper.ActivePosition = new Position(0, 0, 0);
                    symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                    Arc3d arc = symbolGeometryHelper.CreateArc(null, radius + width - d, angle);
                    arc.Transform(matrix);
                    matrix = new Matrix4X4();
                    matrix.SetIdentity();
                    matrix.Translate(new Vector(0, 0, 0));
                    arc.Transform(matrix);
                    curveCollection.Add(arc);

                    line = new Line3d(new Position((radius + d) * Math.Cos(angle), (radius + d) * Math.Sin(angle), 0), new Position((radius - d + width) * Math.Cos(angle), (radius - d + width) * Math.Sin(angle), 0));
                    curveCollection.Add(line);

                    ComplexString3d complexString = new ComplexString3d(curveCollection);
                    Projection3d body = new Projection3d(complexString, new Vector(0, 0, thickness), thickness, true);
                    m_Symbolic.Outputs["BODY"] = body;
                }
                else
                {
                    double depth;
                    double posAngle = 0;
                    depth = thickness * 3 / 4;
                    for (int i = 1; i <= iNoofrung; i++)
                    {
                        posAngle = beginAngle + (i - 1) * (spacing / (radius + width / 2)) - depth / (radius + d) / 2;
                        symbolGeometryHelper = new SymbolGeometryHelper();
                        symbolGeometryHelper.ActivePosition = new Position(0, depth / 2, 0);
                        double tmpLength = width - 2 * d;
                        if (tmpLength > 0)
                        {
                            symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 0, 1));
                        }
                        else if (tmpLength < 0)
                        {
                            symbolGeometryHelper.SetOrientation(new Vector(-1, 0, 0), new Vector(0, 0, 1));
                            tmpLength = Math.Abs(tmpLength);
                        }
                        Projection3d body1 = (Projection3d)symbolGeometryHelper.CreateBox(null, tmpLength, thickness, depth);
                        matrix = new Matrix4X4();
                        matrix.SetIdentity();
                        matrix.Rotate(posAngle, new Vector(0, 0, 1));
                        body1.Transform(matrix);

                        matrix = new Matrix4X4();
                        matrix.SetIdentity();
                        matrix.Translate(new Vector((radius + d) * Math.Cos(posAngle), (radius + d) * Math.Sin(posAngle), thickness / 2));
                        body1.Transform(matrix);
                        m_Symbolic.Outputs["LEGS_"] = body1;

                    }

                    Point3d po = new Point3d(new Position(0, 0, z));
                    m_Symbolic.Outputs["BODY"] = po;
                }

                
            }
            catch //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, TrayShipLocalizer.GetString(TrayShiplResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of TurnTrayCRail.cs."));
                    return;
                }
            }
        }
        #endregion
    }
}

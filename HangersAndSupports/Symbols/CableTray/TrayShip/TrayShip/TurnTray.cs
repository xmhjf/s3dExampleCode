//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   StTrayCirRail.cs
//    TrayShip,Ingr.SP3D.Content.Support.Symbols.StTrayCirRail
//    Author       :  Manikanth
//    Creation Date:  18-01-2013
//    Description:   CR-CP-222297 .Net Hs_TrayShip Creation 


//dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   18-01-2013  Manikanth   CR-CP-222297 .Net Hs_TrayShip Creation 
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
    public class TurnTray : HangerComponentSymbolDefinition
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "TrayShip,Ingr.SP3D.Content.Support.Symbols.TurnTray"
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
        public InputDouble m_Angle;
        [InputDouble(6, "BeginAngle", "BeginAngle", 0.999999)]
        public InputDouble m_BeginAngle;
        [InputDouble(7, "EndAngle", "EndAngle", 0.999999)]
        public InputDouble m_EndAngle;


        [InputString(8, "BOM_DESC", "BOM_DESC", "")]
        public InputString m_BOM_DESC;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]

        [SymbolOutput("Arc1", "Arc1")]
        [SymbolOutput("Line1", "Line1")]
        [SymbolOutput("Arc2", "Arc2")]
        [SymbolOutput("Line2", "Line2")]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("Port3", "Port3")]
        [SymbolOutput("Port4", "Port4")]
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
                double angle = m_Angle.Value;
                double beginAngle = m_BeginAngle.Value;
                double endAngle = m_EndAngle.Value;

                if (HgrCompareDoubleService.cmpdbl(angle % (2 * Math.PI), 0) == true)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, TrayShipLocalizer.GetString(TrayShiplResourceIDs.ErrAnglesArguments, "Angle cannot be zero and multiple of 360°"));
                    return;
                }
                if (HgrCompareDoubleService.cmpdbl(thickness, 0) == true)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, TrayShipLocalizer.GetString(TrayShiplResourceIDs.ErrNETZeroThicknessArguments, "Thickness cannot be zero"));
                    return;
                }
                if (radius <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, TrayShipLocalizer.GetString(TrayShiplResourceIDs.ErrInvalidRadius, "Radius should be greater than zero"));
                    return;
                }
                if (HgrCompareDoubleService.cmpdbl(width, 0) == true)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, TrayShipLocalizer.GetString(TrayShiplResourceIDs.ErrWidthETZArguments, "Width cannot be zero"));
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


                Line3d line;
                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                Collection<ICurve> curveCollection = new Collection<ICurve>();
                Matrix4X4 matrix = new Matrix4X4();

                line = new Line3d(new Position(radius, 0, 0), new Position(radius + width, 0, 0));
                curveCollection.Add(line);

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, 0, 0);
                symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                Arc3d outerArc = symbolGeometryHelper.CreateArc(null, radius, angle);
                outerArc.Transform(matrix);
                matrix = new Matrix4X4();
                matrix.SetIdentity();
                matrix.Translate(new Vector(0, 0, 0));
                outerArc.Transform(matrix);
                curveCollection.Add(outerArc);

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, 0, 0);
                symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                Arc3d arc = symbolGeometryHelper.CreateArc(null, radius + width, angle);
                arc.Transform(matrix);
                matrix = new Matrix4X4();
                matrix.SetIdentity();
                matrix.Translate(new Vector(0, 0, 0));
                arc.Transform(matrix);
                curveCollection.Add(arc);

                line = new Line3d(new Position(radius * Math.Cos(angle), radius * Math.Sin(angle), 0), new Position((radius + width) * Math.Cos(angle), (radius + width) * Math.Sin(angle), 0));
                curveCollection.Add(line);

                ComplexString3d complexString = new ComplexString3d(curveCollection);
                Projection3d body = new Projection3d(complexString, new Vector(0, 0, thickness), thickness, true);
                m_Symbolic.Outputs["BODY"] = body;
            }
            catch //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, TrayShipLocalizer.GetString(TrayShiplResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of TurnTray.cs."));
                    return;
                }
            }
        }
        #endregion

    }

}


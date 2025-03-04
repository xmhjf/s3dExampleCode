//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   StTrayCirRail.cs
//    TrayShip,Ingr.SP3D.Content.Support.Symbols.StTrayCirRail
//    Author       :  Manikanth
//    Creation Date:  14-01-2013
//    Description:   CR-CP-222297 .Net Hs_TrayShip Creation 


//dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   14-01-2013  Manikanth   CR-CP-222297 .Net Hs_TrayShip Creation 
//	 22/03/2013		Vijay 	 DI-CP-228142  Modify the error handling for delivered H&S symbols
//   11-Dec-2014     PVK     TR-CP-264951	Resolve P3 coverity issues found in November 2014 report
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
    [CacheOption(CacheOptionType.NonCached)]
    [SymbolVersion("1.0.0.0")]
    [VariableOutputs]
    public class StTrayCirRail : HangerComponentSymbolDefinition
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "TrayShip,Ingr.SP3D.Content.Support.Symbols.StTrayCirRail"
        //----------------------------------------------------------------------------------


        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "Thickness", "Thickness", 0.999999)]
        public InputDouble m_Thickness;
        [InputDouble(3, "Width", "Width", 0.999999)]
        public InputDouble m_Width;
        [InputDouble(4, "Length", "Length", 0.999999)]
        public InputDouble m_Length;
        [InputDouble(5, "D", "D", 0.999999)]
        public InputDouble m_D;
        [InputDouble(6, "PlateType", "PlateType", 1)]
        public InputDouble m_PlateType;
        [InputString(7, "BOM_DESC", "BOM_DESC", " ")]
        public InputString m_BOM_DESC;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("Port3", "Port3")]
        [SymbolOutput("BODY", "BODY")]
        [SymbolOutput("CYLINDER1", "CYLINDER1")]
        [SymbolOutput("CYLINDER2", "CYLINDER2")]
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
               

                double thickness = m_Thickness.Value;
                double width = m_Width.Value;
                double length = m_Length.Value;
                double d = m_D.Value;
                int pType = (int)m_PlateType.Value;
                double spacing = 0.3;
                double leftOffset;
                const double z = 0.005;

                if (width <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, TrayShipLocalizer.GetString(TrayShiplResourceIDs.ErrNETZeroWidthArguments, "Width should be greater than zero"));
                    return;
                }
                if (length <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, TrayShipLocalizer.GetString(TrayShiplResourceIDs.ErrLengthArguments, "Length should be greater than zero"));
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

                //ports

                Port port1 = new Port(OccurrenceConnection, part, "CableTray1", new Position(0, 0, 0), new Vector(0, 1, 0), new Vector(-1, 0, 0));
                m_Symbolic.Outputs["Port1"] = port1;

                Port port2 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, length / 2), new Vector(0, 1, 0), new Vector(-1, 0, 0));
                m_Symbolic.Outputs["Port2"] = port2;

                Port port3 = new Port(OccurrenceConnection, part, "CableTray2", new Position(0, 0, length), new Vector(0, 1, 0), new Vector(-1, 0, 0));
                m_Symbolic.Outputs["Port3"] = port3;


                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(-thickness / 2, -(width + d) / 2, 0);
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(1, 0, 0));
                Projection3d left = (Projection3d)symbolGeometryHelper.CreateCylinder(null, d / 2, length);
                m_Symbolic.Outputs["CYLINDER1"] = left;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(-thickness / 2, (width + d) / 2, 0);
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(1, 0, 0));
                Projection3d right = (Projection3d)symbolGeometryHelper.CreateCylinder(null, d / 2, length);
                m_Symbolic.Outputs["CYLINDER2"] = right;

                int noOfRung = (int)(length / spacing) + 1;
                if (noOfRung == 0)
                {
                    noOfRung = 1;
                    leftOffset = length / 2;
                }
                else
                {
                    leftOffset = (length - (noOfRung - 1) * spacing) / 2;
                }
                if (pType == 1)
                {
                    symbolGeometryHelper = new SymbolGeometryHelper();
                    symbolGeometryHelper.ActivePosition = new Position(-thickness / 2, 0, 0);
                    symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(1, 0, 0));
                    Projection3d body = (Projection3d)symbolGeometryHelper.CreateBox(null, length, thickness, width);
                    m_Symbolic.Outputs["BODY"] = body;

                }
                else
                {
                    Double rungpos;
                    Double Depth;
                    Point3d p1 = new Point3d(new Position(0, (width + d) / 2, z));
                    m_Symbolic.Outputs["BODY"] = p1;

                    Depth = thickness * 3 / 4;
                    for (int i = 1; i <= noOfRung; i++)
                    {
                        symbolGeometryHelper = new SymbolGeometryHelper();
                        rungpos = leftOffset + (i - 1) * spacing - Depth / 2;
                        symbolGeometryHelper.ActivePosition = new Position(-thickness / 2, 0, rungpos);
                        symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(1, 0, 0));
                        Projection3d legs_ = (Projection3d)symbolGeometryHelper.CreateBox(null, Depth, thickness, width);
                        m_Symbolic.Outputs["LEGS_"] = legs_;

                    }
                }

            }
            catch //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, TrayShipLocalizer.GetString(TrayShiplResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of StTrayCirRail.cs."));
                    return;
                }
            }
        }
        #endregion

    }

}

//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   StTray.cs
//    TrayShip,Ingr.SP3D.Content.Support.Symbols.StTray
//   Author       :  Manikanth
//   Creation Date:  13-01-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   13-01-2013  Manikanth   CR-CP-222297 .Net Hs_TrayShip Creation 
//	 22/03/2013		Vijay 	 DI-CP-228142  Modify the error handling for delivered H&S symbols
//   04/June/2013   Manikanth TR-CP-234520 Implemented TDL Issues
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
    public class StTray : HangerComponentSymbolDefinition
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "TrayShip,Ingr.SP3D.Content.Support.Symbols.StTray"
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
        [InputString(5, "BOM_DESC", "BOM_DESC", "")]
        public InputString m_oOM_DESC;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("BODY", "BODY")]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("Port3", "Port3")]
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

                //ports

                Port port1 = new Port(OccurrenceConnection, part, "CableTray1", new Position(0, 0, 0), new Vector(0, 1, 0), new Vector(-1, 0, 0));
                m_Symbolic.Outputs["Port1"] = port1;

                Port port2 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, length / 2), new Vector(0, 1, 0), new Vector(-1, 0, 0));
                m_Symbolic.Outputs["Port2"] = port2;

                Port port3 = new Port(OccurrenceConnection, part, "CableTray2", new Position(0, 0, length), new Vector(0, 1, 0), new Vector(-1, 0, 0));
                m_Symbolic.Outputs["Port3"] = port3;

                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, 0, 0);
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(1, 0, 0));
                Projection3d body = (Projection3d)symbolGeometryHelper.CreateBox(null, length, thickness, width);
                m_Symbolic.Outputs["BODY"] = body;
            }
            catch //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, TrayShipLocalizer.GetString(TrayShiplResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of StTray.cs."));
                    return;
                }
            }
        }
        #endregion

    }

}

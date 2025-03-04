//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   StTrayCirRail.cs
//    TrayShip,Ingr.SP3D.Content.Support.Symbols.StTrayCirRail
//    Author       :  Manikanth
//    Creation Date:  15-01-2013
//    Description:   CR-CP-222297 .Net Hs_TrayShip Creation 


//dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   15-01-2013  Manikanth   CR-CP-222297 .Net Hs_TrayShip Creation 
//	 22/03/2013		Vijay 	 DI-CP-228142  Modify the error handling for delivered H&S symbols
//   04/June/2013   Manikanth TR-CP-234520 Implemented TDL Issues
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
    [CacheOption(CacheOptionType.NonCached)]
    [SymbolVersion("1.0.0.0")]
    [VariableOutputs]
    public class RectStrap : HangerComponentSymbolDefinition, ICustomHgrWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "TrayShip,Ingr.SP3D.Content.Support.Symbols.RectStrap"
        //----------------------------------------------------------------------------------


        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "Width", "Width", 0.999999)]
        public InputDouble m_Width;
        [InputDouble(3, "Depth", "Depth", 0.999999)]
        public InputDouble m_Depth;
        [InputDouble(4, "Height", "Height", 0.999999)]
        public InputDouble m_Height;
        [InputDouble(5, "Thickness", "Thickness", 0.999999)]
        public InputDouble m_Thickness;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("TOP", "TOP")]
        [SymbolOutput("LEFT", "LEFT")]
        [SymbolOutput("RIGHT", "RIGHT")]
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
               
                double width = m_Width.Value;
                double depth = m_Depth.Value;
                double height = m_Height.Value;
                double thickness = m_Thickness.Value;

                if (width < 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, TrayShipLocalizer.GetString(TrayShiplResourceIDs.ErrWidthArguments, "Width cannot be less than zero"));
                    return;
                }
                if (depth <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, TrayShipLocalizer.GetString(TrayShiplResourceIDs.ErrDepthArguments, "Depth should be greater than zero"));
                    return;
                }
                if (height <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, TrayShipLocalizer.GetString(TrayShiplResourceIDs.ErrHeightArguments, "Height should be greater than zero"));
                    return;
                }
                if (thickness <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, TrayShipLocalizer.GetString(TrayShiplResourceIDs.ErrThicknessArguments, "Thickness should be greater than zero"));
                    return;
                }

                //ports

                Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;
                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, 0, height - thickness);
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 1, 0));
                Projection3d top = (Projection3d)symbolGeometryHelper.CreateBox(null, thickness, width + 2 * thickness, depth);
                m_Symbolic.Outputs["TOP"] = top;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, -width / 2 - thickness / 2, -thickness);
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 1, 0));
                Projection3d left = (Projection3d)symbolGeometryHelper.CreateBox(null, height, thickness, depth);
                m_Symbolic.Outputs["LEFT"] = left;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, width / 2 + thickness / 2, -thickness);
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 1, 0));
                Projection3d right = (Projection3d)symbolGeometryHelper.CreateBox(null, height, thickness, depth);
                m_Symbolic.Outputs["RIGHT"] = right;
            }
            catch //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, TrayShipLocalizer.GetString(TrayShiplResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of RectStrap.cs."));
                    return;
                }
            }
        }
        #endregion

        #region "ICustomHgrWeightCG Members"

        void ICustomHgrWeightCG.WeightCG(SupportComponent supportComponent, ref double weight, ref double cogX, ref double cogY, ref double cogZ)
        {
            try
            {
                double thickNess;
                double width;
                double height;
                double depth;
                width = (double)((PropertyValueDouble)supportComponent.GetPropertyValue("IJOAHgrTrayWidth", "Width")).PropValue;
                thickNess = (double)((PropertyValueDouble)supportComponent.GetPropertyValue("IJOAHgrTrayThickness", "Thickness")).PropValue;
                height = (double)((PropertyValueDouble)supportComponent.GetPropertyValue("IJOAHgrTrayHeight", "Height")).PropValue;
                depth = (double)((PropertyValueDouble)supportComponent.GetPropertyValue("IJOAHgrTrayDepth", "Depth")).PropValue;

                weight = (width * depth * thickNess) + (2 * (thickNess * height * depth)) * 7900;
                cogX = 0;
                cogY = 0;
                cogZ = 0;

            }
            catch //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, TrayShipLocalizer.GetString(TrayShiplResourceIDs.ErrWeightCG, "Error in WeightCG of RectStrap.cs."));
                }
            }
        }
        #endregion

    }

}

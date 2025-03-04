//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   CCircle.cs
//    HgrSupPad,Ingr.SP3D.Content.Support.Symbols.CCircle
//   Author       :  Rajeswari
//   Creation Date:  17-Dec-2012
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   17-Dec-2012   Rajeswari CR-CP-222301 .Net HgrSupPad project creation
//  18/03/2013     Rajeswari DI-CP-228142  Modify the error handling for delivered H&S symbols
//  6/11/2013      Vijaya    CR-CP-242533  Provide the ability to store GType outputs as a single blob for H&S .net symbols  
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
    [VariableOutputs]
    [CacheOption(CacheOptionType.NonCached)]
    [SymbolVersion("1.0.0.0")]
    public class CCircle : HangerComponentSymbolDefinition
    {
        #region "Definition of Inputs"
        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "PadDiameterLimit", "PadDiameterLimit", 0.999999)]
        public InputDouble m_dPadDiameterLimit;
        [InputDouble(3, "Thickness", "Thickness", 0.999999)]
        public InputDouble m_dThickness;
        [InputDouble(4, "Width", "Width", 0.999999)]
        public InputDouble m_dWidth;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("HgrPort_1", "Description for HgrPort")]
        [SymbolOutput("HgrPort_2", "Description for HgrPort")]
        [SymbolOutput("Circle", "Circle")]
        
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
                Double padDiameterLimit = m_dPadDiameterLimit.Value;
                Double thickness =m_dThickness.Value;
                Double width = m_dWidth.Value;
                Double padDiameter;

                if (padDiameterLimit < 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, HgrSupPadLocalizer.GetString(HgrSupPadSymbolsResourceIDs.ErrInvalidpadDiameterLimit, "Pad Diameter Limit cannot be negative"));
                    return;
                }

                if (thickness == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, HgrSupPadLocalizer.GetString(HgrSupPadSymbolsResourceIDs.ErrInvalidThickness, "Thickness cannot be zero"));
                    return;
                }
                padDiameter = padDiameterLimit + 2 * width;

                //--------------------------------------------------
                //Create the Circle and Projection
                //--------------------------------------------------
                Circle3d createCircle = new Circle3d(new Position(0, 0, 0), new Vector(0, 0, 1), padDiameter * 0.5);
                Projection3d circle = new Projection3d(createCircle, new Vector(0, 0, 1), thickness, true);
                m_Symbolic.Outputs["Circle"] = circle;

                //CREATE HANGER PORTS
                Port port1 = new Port(OccurrenceConnection, part, "HgrPort_1", new Position(0, 0, 0), new Vector(1.0, 0, 0.0), new Vector(0, 0.0, 1.0));
                m_Symbolic.Outputs["HgrPort_1"] = port1;
                Port port2 = new Port(OccurrenceConnection, part, "HgrPort_2", new Position(0, 0, thickness), new Vector(1.0, 0, 0.0), new Vector(0, 0.0, 1.0));
                m_Symbolic.Outputs["HgrPort_2"] = port2;
            }
            catch
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, HgrSupPadLocalizer.GetString(HgrSupPadSymbolsResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of CCircle.cs."));
                    return;
                }
            }
        }
        #endregion
    }
}

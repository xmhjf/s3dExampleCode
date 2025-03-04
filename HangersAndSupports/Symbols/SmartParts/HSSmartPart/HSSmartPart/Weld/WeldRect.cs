//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   WeldRect.cs
//   HSSmartPart,Ingr.SP3D.Content.Support.Symbols.WeldRect
//   Author       :  Shilpi
//   Creation Date:  5.October.2012
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   5.October.2012  Shilpi     CR-CP-221374 Converted the HS_Weld VB Project to C#.Net
//   9.Jan.2013      Hema       CR-CP-221374 Modified the WeightCG Implementation
//   25/Mar/2013     Sridhar    DI-CP-228142  Modify the error handling for delivered H&S symbols
//   30/10/2013     Hema    CR-CP-242533  Provide the ability to store GType outputs as a single blob.
//   04/12/2015     Ramya       TR 280973	Weight and CG calculations have to be removed for Weld
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.Support.Middle;

namespace Ingr.SP3D.Content.Support.Symbols
{
    //-----------------------------------------------------------------------------------
    //Namespace of this class is Ingr.SP3D.Support.Content.Symbols
    //It is recommended that customers specify namespace of their symbols to be
    //CompanyName.SP3D.Content.Specialization.
    //It is also recommended that if customers want to change this symbol to suit their
    //requirements, they should change namespace/symbol name so the identity of the modified
    //symbol will be different from the one delivered by Intergraph.
    //-----------------------------------------------------------------------------------
    [CacheOption(CacheOptionType.Cached)]
    [SymbolVersion("1.0.0.0")]
    [VariableOutputs]
    public class WeldRect : HangerComponentSymbolDefinition
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "HSSmartPart,Ingr.SP3D.Content.Support.Symbols.WeldRect"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "Width", "Width", 0)]
        public InputDouble m_dWidth;
        [InputDouble(3, "Height", "Height", 0)]
        public InputDouble m_dHeight;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Weld", "Weld")]
        [SymbolOutput("LINE1", "LINE1")]
        [SymbolOutput("LINE2", "LINE2")]
        [SymbolOutput("LINE3", "LINE3")]
        [SymbolOutput("LINE4", "LINE4")]
        public AspectDefinition m_PhysicalAspect;

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
                Double width = m_dWidth.Value;
                Double height = m_dHeight.Value;

                Port weld = new Port(OccurrenceConnection, part, "Other", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_PhysicalAspect.Outputs["Weld"] = weld;

                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                //ports

                Line3d line;
                Collection<ICurve> lineCollection = new Collection<ICurve>();

                line = new Line3d(new Position(0, 0, 0), new Position(width, 0, 0));
                m_PhysicalAspect.Outputs["Line1"] = line;
                line = new Line3d( new Position(width, 0, 0), new Position(width, height, 0));
                m_PhysicalAspect.Outputs["Line2"] = line;
                line = new Line3d( new Position(width, height, 0), new Position(0, height, 0));
                m_PhysicalAspect.Outputs["Line3"] = line;
                line = new Line3d( new Position(0, height, 0), new Position(0, 0, 0));
                m_PhysicalAspect.Outputs["Line4"] = line;
            }
            catch //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of Weld Rect"));
                }
            }
        }
        #endregion

    }
}

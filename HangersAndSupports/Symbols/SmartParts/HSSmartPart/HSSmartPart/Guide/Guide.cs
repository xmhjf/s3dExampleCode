//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.

//   Guide.cs
//   Author       :  Hema
//   Creation Date:  27.Jan.2013 
//   Description:    Converted Guide Smartpart VB Project to C# .Net 

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   27.Jan.2013     Hema   CR-CP-222484 Converted Guide Smartpart VB Project to C# .Net 
//   25/Mar/2013     Hema   DI-CP-228142  Modify the error handling for delivered H&S symbols
//   31/10/2013     Rajeswari    CR-CP-242533  Provide the ability to store GType outputs as a single blob.
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using System.Collections.Generic;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.Support.Middle;

namespace Ingr.SP3D.Content.Support.Symbols
{
    //-----------------------------------------------------------------------------------
    //Namespace of this class is  Ingr.SP3D.Content.Support.Symbols
    //It is recommended that customers specify namespace of their symbols to be
    //CompanyName.SP3D.Content.Specialization.
    //It is also recommended that if customers want to change this symbol to suit their
    //requirements, they should change namespace/symbol name so the identity of the modified
    //symbol will be different from the one delivered by Intergraph.
    //-----------------------------------------------------------------------------------

    [SymbolVersion("1.0.0.0")]
    [VariableOutputs]
    [CacheOption(CacheOptionType.Cached)]

    public class Guide : SmartPartComponentDefinition
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "HSSmartPart, Ingr.SP3D.Content.Support.Symbols.Guide"
        //----------------------------------------------------------------------------------


        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "PipeOD", "PipeOD", 0)]
        public InputDouble m_dPipeOD;
        [InputDouble(3, "ShoeHeight", "ShoeHeight", 0)]
        public InputDouble m_dShoeHeight;
        [InputDouble(4, "Width4", "Width4", 0)]
        public InputDouble m_dWidth4;
        [InputDouble(5, "Length4", "Length4", 0)]
        public InputDouble m_dLength4;
        [InputDouble(6, "Thickness4", "Thickness4", 0)]
        public InputDouble m_dThickness4;
        [InputDouble(7, "Width5", "Width5", 0)]
        public InputDouble m_dWidth5;
        [InputDouble(8, "Length5", "Length5", 0)]
        public InputDouble m_dLength5;
        [InputDouble(9, "Thickness5", "Thickness5", 0)]
        public InputDouble m_dThickness5;
        [InputDouble(10, "Width6", "Width6", 0)]
        public InputDouble m_dWidth6;
        [InputDouble(11, "Length6", "Length6", 0)]
        public InputDouble m_dLength6;
        [InputDouble(12, "Multi1Qty", "Multi1Qty", 0)]
        public InputDouble m_dMulti1Qty;
        [InputDouble(13, "Multi1LocateBy", "Multi1LocateBy", 0)]
        public InputDouble m_dMulti1LocateBy;
        [InputDouble(14, "Multi1Location", "Multi1Location", 0)]
        public InputDouble m_dMulti1Location;
        [InputDouble(15, "Offset3", "Offset3", 0)]
        public InputDouble m_dOffset3;
        [InputDouble(16, "Offset4", "Offset4", 0)]
        public InputDouble m_dOffset4;
        [InputDouble(17, "CPXOffset", "CPXOffset", 0, true)]
        public InputDouble m_dCPXOffset;
        [InputDouble(18, "CPYOffset", "CPYOffset", 0, true)]
        public InputDouble m_dCPYOffset;
        [InputDouble(19, "GuideHeight", "GuideHeight", 0, true)]
        public InputDouble m_dGuideHeight;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Route", "Route")]
        [SymbolOutput("Structure", "Structure")]
        [SymbolOutput("Surface", "Surface")]
        [SymbolOutput("Block", "Block")]
        public AspectDefinition m_PhysicalAspect;

        #endregion
        #region "Definition of Additional Inputs"
        public override IEnumerable<Input> AdditionalInputs
        {
            get
            {
                int endIndex;
                List<Input> additionalInputs = new List<Input>();
                AddGuideInputs(20, out endIndex, additionalInputs);
                return additionalInputs;
            }
        }
        #endregion
        #region "Definition of Additional Outputs"
        public override IEnumerable<OutputDefinition> AdditionalOutputs(string aspectName)
        {
            List<OutputDefinition> additionalOutputs = new List<OutputDefinition>();
            if (aspectName == "Symbolic")
            {
                AddGuideOutputs(additionalOutputs);
            }
            return additionalOutputs;
        }
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

                Double pipeOD = m_dPipeOD.Value;
                Double shoeHeight = m_dShoeHeight.Value;
                Double width4 = m_dWidth4.Value;
                Double length4 = m_dLength4.Value;
                Double thickness4 = m_dThickness4.Value;
                Double width5 = m_dWidth5.Value;
                Double length5 = m_dLength5.Value;
                Double thickness5 = m_dThickness5.Value;
                Double width6 = m_dWidth6.Value;
                Double length6 = m_dLength6.Value;
                Double multi1Qty = m_dMulti1Qty.Value;
                Double multi1LocateBy = m_dMulti1LocateBy.Value;
                Double multi1Location = m_dMulti1Location.Value;
                Double offset3 = m_dOffset3.Value;
                Double offset4 = m_dOffset4.Value;
                Matrix4X4 matrix = new Matrix4X4();

                //Loading the Guide Additional Inputs
                GuideInputs guide = LoadGuideData(20, part);
                if (base.ToDoListMessage != null)
                    if (base.ToDoListMessage.Type == ToDoMessageTypes.ToDoMessageError)
                        return;
                //Initializing SymbolGeometryHelper
                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();

                //=================================================
                //Construction of Physical Aspect 
                //=================================================

                //ports
                Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, guide.Thickness1 + thickness4 + shoeHeight + pipeOD / 2), new Vector(0, 1, 0), new Vector(0, 0, -1));
                m_PhysicalAspect.Outputs["Route"] = port1;
                Port port2 = new Port(OccurrenceConnection, part, "Structure", new Position(0, 0, 0), new Vector(0, 1, 0), new Vector(0, 0, 1));
                m_PhysicalAspect.Outputs["Structure"] = port2;
                Port port3 = new Port(OccurrenceConnection, part, "Surface", new Position(0, 0, guide.Thickness1 + thickness4), new Vector(0, 1, 0), new Vector(0, 0, 1));
                m_PhysicalAspect.Outputs["Surface"] = port3;

                //Add the Guide
                matrix.SetIdentity();
                matrix.Rotate((Math.PI / 2), new Vector(0, 0, 1));

                AddGuide(guide, matrix, m_PhysicalAspect.Outputs, "Guide");

                // Add The Block
                if (width4 > 0 && length4 > 0 && thickness4 > 0)
                {
                    symbolGeometryHelper.ActivePosition = new Position(0, 0, guide.Thickness1);
                    symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 1, 0));
                    Projection3d block = (Projection3d)symbolGeometryHelper.CreateBox(null, thickness4, length4, width4);
                    m_PhysicalAspect.Outputs["Block"] = block;
                }

                //Add the Gussets
                if (width5 > 0 && length5 > 0 && thickness5 > 0)
                {
                    matrix = new Matrix4X4();
                    matrix.SetIdentity();
                    matrix.Rotate((Math.PI / 2), new Vector(0, 0, 1));

                    Vector vector = matrix.Transform(new Vector(0, -offset3, offset4));
                    matrix.Translate(vector);

                    AddGussetsWithChamferByRow(guide.Length2, multi1Qty, multi1LocateBy, multi1Location, width5, length5, thickness5, width6, length6, matrix, m_PhysicalAspect.Outputs, "Gussets", 1);

                    matrix = new Matrix4X4();
                    matrix.SetIdentity();
                    matrix.Rotate(3 * (Math.PI / 2), new Vector(0, 0, 1));

                    Vector vector1 = matrix.Transform(new Vector(0, -offset3, offset4));
                    matrix.Translate(vector1);

                    AddGussetsWithChamferByRow(guide.Length2, multi1Qty, multi1LocateBy, multi1Location, width5, length5, thickness5, width6, length6, matrix, m_PhysicalAspect.Outputs, "Gussets", 2);
                }
            }
            catch //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of Guide"));
                }
            }
        }
        #endregion
    }
}

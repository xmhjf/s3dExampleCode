//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

//   Copyright (c) 2009, Intergraph Corporation. All rights reserved.

//   ProgID:         HS_S3DWeld.WeldSection
//   Author:
//   Creation Date:

//   Description:
//       WeldSection

//   dd.mmm.yyyy     who         change description
//   -----------     ---         ------------------

//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using System.Collections.Generic;
using System.Linq;
using System.Collections.ObjectModel;
using Ingr.SP3D.Structure.Middle.Services;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.ReferenceData.Middle.Services;

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
    public class WeldSection : HangerComponentSymbolDefinition
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "HSSmartPart,Ingr.SP3D.Content.Support.Symbols.WeldSection"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputString(2, "StandardName", "StandardName", "No Value")]
        public InputString m_sStandardName;
        [InputString(3, "SectionType", "SectionType", "No Value")]
        public InputString m_sSectionType;
        [InputString(4, "SectionName", "SectionName", "No Value")]
        public InputString m_sSectionName;
        [InputDouble(5, "WeldCardinalPoint", "WeldCardinalPoint", 0)]
        public InputDouble m_dWeldCardinalPoint;
        [InputDouble(6, "Rotation", "Rotation", 0)]
        public InputDouble m_dRotation;
        [InputDouble(7, "Reflect", "Reflect", 0)]
        public InputDouble m_dReflect;

        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("WeldPort", "WeldPort")]
        [SymbolOutput("WeldPath", "WeldPath")]
        public AspectDefinition m_PhysicalAspect;

        #endregion

        #region "Construct Outputs"

        /// <summary>
        /// Construct symbol outputs in aspects.
        /// </summary>
        /// <remarks></remarks>
        /// 

        protected override void ConstructOutputs()
        {
            try
            {
                Part part = (Part)m_PartInput.Value;
                String sStandardName = m_sStandardName.Value;
                String sSectionType = m_sSectionType.Value;
                String sSectionName = m_sSectionName.Value;
                int lWeldCP = (int)m_dWeldCardinalPoint.Value;
                Double dRotation = m_dRotation.Value;
                Boolean bReflect = Convert.ToBoolean(m_dReflect.Value);

                Port weld = new Port(OccurrenceConnection, part, "Other", new Position(0, 0, 0), new Vector(0, 1, 0), new Vector(0, 0, -1));
                m_PhysicalAspect.Outputs["WeldPort"] = weld;

                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                CrossSectionServices crossSectionServices = new CrossSectionServices();
                CatalogStructHelper catalogStructHelper = new CatalogStructHelper();
                CrossSection crossSection = catalogStructHelper.GetCrossSection(sStandardName, sSectionType, sSectionName);

                ReadOnlyCollection<SectionGeometry> sectionGeometryColl = crossSectionServices.GetSectionGeometry(crossSection);
                Collection<ICurve> colCurves;
                sectionGeometryColl[0].OuterGeometry.GetCurves(out colCurves);
                ComplexString3d crossSectionOuterGeometry = new ComplexString3d(colCurves);

                double CPX, CPY;
                crossSectionServices.GetCardinalPointOffset(crossSection, lWeldCP, out CPX, out CPY);

                Matrix4X4 matrix = new Matrix4X4(); 
                matrix.SetIdentity();
                if (bReflect)
                    matrix.Rotate(Math.PI, new Vector(0, 1, 0));
                matrix.Rotate(dRotation, new Vector(0, 0, 1));
                matrix.Translate(matrix.Transform(new Vector(-CPX, -CPY, 0)));


                crossSectionOuterGeometry.Transform(matrix);

                m_PhysicalAspect.Outputs["WeldPath"] = crossSectionOuterGeometry;
            }
            catch //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of WeldPoint"));
                }
            }
        }
        #endregion

    }

}


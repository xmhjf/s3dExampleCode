//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.

//   GBeam.cs
//   GeneralProfiles,Ingr.SP3D.Content.Support.Symbols.GBeam
//   Author       :  Hema
//   Creation Date:  05.June.2013 
//   Description:    Converted GeneralProfileSymbols VB Project to C# .Net 

//   Change History:
//   dd.mmm.yyyy     who     change descriptions
//   -----------     ---     ------------------
//   05.June.2013    Hema   CR-CP-222298 Converted GeneralProfileSymbols VB Project to C# .Net 
//   06.June.2016    PVK    TR-CP-296065	Fix new coverity issues found in H&S Content 
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
using System;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.ReferenceData.Middle.Services;
using Ingr.SP3D.Structure.Middle.Services;
namespace Ingr.SP3D.Content.Support.Symbols
{
    //-----------------------------------------------------------------------------------
    //Namespace of this class is Ingr.SP3D.Content.Support.Symbols
    //It is recommended that customers specify namespace of their symbols to be
    //<CompanyName>.SP3D.Content.<Specialization>.
    //It is also recommended that if customers want to change this symbol to suit their
    //requirements, they should change namespace/symbol name so the identity of the modified
    //symbol will be different from the one delivered by Intergraph.
    //-----------------------------------------------------------------------------------
    [CacheOption(CacheOptionType.NonCached)]
    [SymbolVersion("1.0.0.0")]
    [VariableOutputs]
    public class GBeam : HangerComponentSymbolDefinition, ICustomWeightCG
    {
        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputString(2, "SP3D_SectionName", "Cross Section Name of the structure part", "No Value")]
        public InputString m_strSP3D_SectionName;
        [InputDouble(3, "Length", "Length of the Beam", 10)]
        public InputDouble m_dLength;
        [InputString(4, "MaterialGrade", "MaterialGrade of the Beam", "No Value")]
        public InputString m_strMaterialGrade;
        [InputDouble(5, "Orientation", "Orientation", 0)]
        public InputDouble m_dOrientation;
        [InputDouble(6, "BeginOverLength", "Begin OverLength of the Beam", 0)]
        public InputDouble m_dBeginOverLength;
        [InputDouble(7, "EndOverLength", "End OverLength of the Beam", 0)]
        public InputDouble m_dEndOverLength;
        [InputString(8, "SP3D_Standard", "SP3D_Standard", "No Value")]
        public InputString m_strSP3D_Standard;
        [InputString(9, "SP3D_ClassName", "SP3D_ClassName", "No Value")]
        public InputString m_strSP3D_ClassName;
        [InputDouble(10, "BeginMiter", "BeginMiter", 0)]
        public InputDouble m_dBeginMiter;
        [InputDouble(11, "EndMiter", "Tray Width", 0)]
        public InputDouble m_dEndMiter;


        #endregion
        #region "Definitions of Aspects and their outputs"

        // Physical Aspect 
        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("General Beam", "General Beam")]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("Port3", "Port3")]
        [SymbolOutput("Port4", "Port4")]
        [SymbolOutput("Port5", "Port5")]
        [SymbolOutput("Port6", "Port6")]
        [SymbolOutput("Port7", "Port7")]

        public AspectDefinition m_Symbolic;

        #endregion

        #region "Construction of outputs of all aspects"
        protected override void ConstructOutputs()
        {
            try
            {
                Part part = (Part)m_PartInput.Value;

                String sectionName = m_strSP3D_SectionName.Value;
                String sectionStandard = m_strSP3D_Standard.Value;
                String sectionType = m_strSP3D_ClassName.Value;
                Double length = m_dLength.Value;
                Double beginOverLength = m_dBeginOverLength.Value;
                Double endOverLength = m_dEndOverLength.Value;

                CatalogStructHelper catalogStructHelper = new CatalogStructHelper();
                
                CrossSection sectionGeometry = catalogStructHelper.GetCrossSection(sectionStandard, sectionType, sectionName);

                Double height, width, X, Y, Z;

                Double totalLength = beginOverLength + length + endOverLength;
                if (sectionGeometry == null)
                {
                    Collection<Position> pointCollection = new Collection<Position>();
                    pointCollection.Add(new Position(-0.05, 0.0, 0.0));
                    pointCollection.Add(new Position(0.05, 0.0, 0.0));
                    pointCollection.Add(new Position(0.05, 0.002, 0.0));
                    pointCollection.Add(new Position(-0.048, 0.002, 0.0));
                    pointCollection.Add(new Position(-0.048, 0.1, 0.0));
                    pointCollection.Add(new Position(-0.05, 0.1, 0.0));
                    pointCollection.Add(new Position(-0.05, 0.0, 0.0));

                    height = 0.1;
                    width = 0.1;
                    X = 0.0;
                    Y = 0.05;
                    Z = 0.0;

                    Projection3d projection = new Projection3d(new LineString3d(pointCollection), new Vector(0, 0, 1), totalLength, true);
                    m_Symbolic.Outputs["General Beam"] = projection;
                }
                else
                {
                    //retrieve bounding volume
                    width = 0;
                    height = 0;
                    X = width / 2.0;
                    Y = height / 2.0;
                    Z = 0.0;
                    CrossSectionServices crossSectionServices = new CrossSectionServices();
                    Collection<ISurface> surfaces = crossSectionServices.GetProjectionSurfacesFromCrossSection(sectionGeometry, new Line3d(new Position(X, Y, 0), new Position(X, Y, totalLength)), 5, false, 0, (SweepOptions)1);

                    m_Symbolic.Outputs["General Beam"] = surfaces[0];
                }

                Port port1 = new Port(OccurrenceConnection, part, "BeginCap", new Position(X, Y, Z + beginOverLength), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;

                Port port2 = new Port(OccurrenceConnection, part, "Top", new Position(X + width / 2.0, Y, Z + beginOverLength + length / 2.0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;

                Port port3 = new Port(OccurrenceConnection, part, "BeginCap", new Position(X - width / 2.0, Y, Z + beginOverLength + length / 2.0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port3"] = port3;

                Port port4 = new Port(OccurrenceConnection, part, "Right", new Position(X, Y + height / 2.0, Z + beginOverLength + length / 2.0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port4"] = port4;

                Port port5 = new Port(OccurrenceConnection, part, "Left", new Position(X, Y - height / 2.0, Z + beginOverLength + length / 2.0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port5"] = port5;

                Port port6 = new Port(OccurrenceConnection, part, "Neutral", new Position(X, Y, Z + beginOverLength + length / 2.0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port6"] = port6;

                Port port7 = new Port(OccurrenceConnection, part, "EndCap", new Position(X, Y, Z + beginOverLength + length), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port7"] = port7;
            }
            catch
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, GeneralProfilesLocalizer.GetString(GeneralProfilesSymbolsResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of GBeam.cs."));
                    return;
                }
            }
        }
        #endregion

        #region ICustomHgrWeightCG Members

        void ICustomWeightCG.EvaluateWeightCG(BusinessObject supportComponentBO)
        {
            try
            {
                double density = 0.25;
                Part catalogPart = (Part)supportComponentBO.GetRelationship("madeFrom", "part").TargetObjects[0];
                double length, beginOverLength, endOverLength, weight, cogX, cogY, cogZ, X = 0, Y = 0;

                try
                {
                    length = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJUAHgrOccLength", "Length")).PropValue;
                }
                catch
                {
                    length = 0;
                }
                try
                {
                    beginOverLength = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJUAHgrOccOverLength", "BeginOverLength")).PropValue;
                }
                catch
                {
                    beginOverLength = 0;
                }
                try
                {
                    endOverLength = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJUAHgrOccOverLength", "EndOverLength")).PropValue;
                }
                catch
                {
                    endOverLength = 0;
                }
                double totalLength = beginOverLength + length + endOverLength;
                CrossSection crossSection=null;
                try
                {
                    crossSection = (CrossSection)catalogPart.GetRelationship("HgrCrossSection", "CrossSection").TargetObjects[0];
                }
                catch
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, GeneralProfilesLocalizer.GetString(GeneralProfilesSymbolsResourceIDs.ErrCrossSectionNotFound, "Unable to get Cross-section object"));
                    crossSection = null;
                }
                CrossSectionServices crossSectionServices = new CrossSectionServices();
                if (crossSection != null)
                    crossSectionServices.GetCardinalPointOffset(crossSection, 10, out X, out Y);
                try
                {
                    weight = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryWeight")).PropValue;
                }
                catch
                {
                    weight = density * totalLength;
                }
                //Center of Gravity
                try
                {
                    cogX = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryCogX")).PropValue;
                }
                catch
                {
                    cogX = X;
                }
                try
                {
                    cogY = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryCogY")).PropValue;
                }
                catch
                {
                    cogY = Y;
                }
                try
                {
                    cogZ = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryCogZ")).PropValue;
                }
                catch
                {
                    cogZ = totalLength / 2.0;
                }
                SupportComponent supportComponent = (SupportComponent)supportComponentBO;
                supportComponent.SetWeightAndCOG(weight, cogX, cogY, cogZ);

            }
            catch
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, GeneralProfilesLocalizer.GetString(GeneralProfilesSymbolsResourceIDs.ErrWeightCG, "Error in EvaluateWeightCG of GBeam.cs"));
            }
        }
        #endregion
    }
}

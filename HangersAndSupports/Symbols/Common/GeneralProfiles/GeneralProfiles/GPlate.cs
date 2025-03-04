//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.

//   GPlate.cs
//   GeneralProfiles,Ingr.SP3D.Content.Support.Symbols.GPlate
//   Author       :  Hema
//   Creation Date:  05.June.2013 
//   Description:    Converted GeneralProfileSymbols VB Project to C# .Net 

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   05.June.2013    Hema   CR-CP-222298 Converted GeneralProfileSymbols VB Project to C# .Net 
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
    public class GPlate : HangerComponentSymbolDefinition
    {
        //static Boolean firstTime = true;
        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "Length", "Length Of the CableTray", 40.0)]
        public InputDouble m_dLength;
        [InputDouble(3, "Width", "Width of the CableTray", 10.0)]
        public InputDouble m_dWidth;
        [InputDouble(4, "Thickness", "Thickness of the CableTray", 10.0)]
        public InputDouble m_dThickness;

        #endregion
        #region "Definitions of Aspects and their outputs"

        // Physical Aspect 
        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("CableTray", "CableTray")]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("Port3", "Port3")]
        [SymbolOutput("Port4", "Port4")]
        [SymbolOutput("Port5", "Port5")]

        public AspectDefinition m_Symbolic;

        #endregion

        #region "Construction of outputs of all aspects"
        protected override void ConstructOutputs()
        {
            try
            {
                Part part = (Part)m_PartInput.Value;
                Double length, width, thickness;

                if (string.IsNullOrEmpty(((SupportComponent)Occurrence).Name))
                {
                    length = m_dLength.Value;
                    width = m_dWidth.Value;
                    thickness = m_dThickness.Value;
                    double pipeRadius = 0;
                    Ingr.SP3D.Support.Middle.Support support = (Ingr.SP3D.Support.Middle.Support)Occurrence.GetRelationship("SupportHasComponents", "Support").TargetObjects[0];

                    if (support.IsDesignSupportAssembly())
                    {
                        SupportedHelper supportedHelper = new SupportedHelper(support);
                        if (supportedHelper.SupportedObjectInfo(1).SupportedObjectType == SupportedObjectType.CableTray)
                            pipeRadius = ((CableTrayObjectInfo)supportedHelper.SupportedObjectInfo(1)).Width / 2.0;
                        else if (supportedHelper.SupportedObjectInfo(1).SupportedObjectType == SupportedObjectType.Pipe)
                            pipeRadius = ((PipeObjectInfo)supportedHelper.SupportedObjectInfo(1)).NominalDiameter.Size / 2.0;
                        else if (supportedHelper.SupportedObjectInfo(1).SupportedObjectType == SupportedObjectType.HVAC)
                            pipeRadius = ((DuctObjectInfo)supportedHelper.SupportedObjectInfo(1)).OutsideDiameter / 2.0;
                        else if (supportedHelper.SupportedObjectInfo(1).SupportedObjectType == SupportedObjectType.Conduit)
                            pipeRadius = ((ConduitObjectInfo)supportedHelper.SupportedObjectInfo(1)).OutsideDiameter / 2.0;

                        width = MiddleServiceProvider.UOMMgr.ConvertUnitToUnit(UnitType.Distance, 8.0 * pipeRadius, UnitName.DISTANCE_INCH, UnitName.DISTANCE_METER);
                        length = MiddleServiceProvider.UOMMgr.ConvertUnitToUnit(UnitType.Distance, 5.0 * pipeRadius, UnitName.DISTANCE_INCH, UnitName.DISTANCE_METER);
                        thickness = MiddleServiceProvider.UOMMgr.ConvertUnitToUnit(UnitType.Distance, 0.5 * pipeRadius, UnitName.DISTANCE_INCH, UnitName.DISTANCE_METER);

                        Occurrence.SetPropertyValue(width, "IJUAHgrOccGeometry", "Width");
                        Occurrence.SetPropertyValue(length, "IJUAHgrOccLength", "Length");
                        Occurrence.SetPropertyValue(thickness, "IJUAHgrThickness", "Thickness");
                    }
                    else
                    {
                        length = (double)((PropertyValueDouble)Occurrence.GetPropertyValue("IJUAHgrOccLength", "Length")).PropValue;
                        width = (double)((PropertyValueDouble)Occurrence.GetPropertyValue("IJUAHgrOccGeometry", "Width")).PropValue;
                        thickness = (double)((PropertyValueDouble)Occurrence.GetPropertyValue("IJUAHgrThickness", "Thickness")).PropValue;
                    }
                }
                else
                {
                    length = (double)((PropertyValueDouble)Occurrence.GetPropertyValue("IJUAHgrOccLength", "Length")).PropValue;
                    width = (double)((PropertyValueDouble)Occurrence.GetPropertyValue("IJUAHgrOccGeometry", "Width")).PropValue;
                    thickness = (double)((PropertyValueDouble)Occurrence.GetPropertyValue("IJUAHgrThickness", "Thickness")).PropValue;
                }
                Collection<Position> pointCollection = new Collection<Position>();
                pointCollection.Add(new Position(width / 2.0, length / 2.0, 0.0));
                pointCollection.Add(new Position(-width / 2.0, length / 2.0, 0.0));
                pointCollection.Add(new Position(-width / 2.0, -length / 2.0, 0.0));
                pointCollection.Add(new Position(width / 2.0, -length / 2.0, 0.0));
                pointCollection.Add(new Position(width / 2.0, length / 2.0, 0.0));
                pointCollection.Add(new Position(width / 2.0, length / 2.0, 0.0));

                if (thickness == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, GeneralProfilesLocalizer.GetString(GeneralProfilesSymbolsResourceIDs.ErrInvalidThickness, "Thickness cannot be zero"));
                    return;
                }
                if (width == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, GeneralProfilesLocalizer.GetString(GeneralProfilesSymbolsResourceIDs.ErrInvalidWidth, "Width cannot be zero"));
                    return;
                }
                Projection3d projection = new Projection3d(new LineString3d(pointCollection), new Vector(0, 0, 1), thickness * 1.0, true);
                m_Symbolic.Outputs["CableTray"] = projection;

                Port port1 = new Port(OccurrenceConnection, part, "Center", new Position(0, 0, thickness / 2.0), new Vector(0, 1, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;

                Port port2 = new Port(OccurrenceConnection, part, "Front", new Position(0.0, length / 2.0, thickness / 2.0), new Vector(0, 1, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;

                Port port3 = new Port(OccurrenceConnection, part, "Right", new Position(width / 2.0, 0.0, thickness / 2.0), new Vector(0, 1, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port3"] = port3;

                Port port4 = new Port(OccurrenceConnection, part, "End", new Position(0.0, -length / 2.0, thickness / 2.0), new Vector(0, 1, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port4"] = port4;

                Port port5 = new Port(OccurrenceConnection, part, "Left", new Position(-width / 2.0, 0.0, thickness / 2.0), new Vector(0, 1, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port5"] = port5;

            }
            catch
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, GeneralProfilesLocalizer.GetString(GeneralProfilesSymbolsResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of GPlate.cs."));
                    return;
                }
            }
        }
        #endregion
    }
}
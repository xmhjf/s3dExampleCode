//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   BracingPlate.cs
//   HgrOglaendParts,Ingr.SP3D.Content.Support.Symbols.BracingPlate
//   Author       :  Ramya
//   Creation Date:  6.June.2012
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   6.June.2012     Ramya     Initial Creation
//   23-01-2015      Chethan  CR-CP-241198  Convert all Oglaend Content to .NET 
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
    //<CompanyName>.SP3D.Content.Support.<Specialization>.
    //It is also recommended that if customers want to change this symbol to suit their
    //requirements, they should change namespace/symbol name so the identity of the modified
    //symbol will be different from the one delivered by Intergraph.
    //-----------------------------------------------------------------------------------
    [VariableOutputs]
    [CacheOption(CacheOptionType.Cached)]
    [SymbolVersion("1.0.0.0")]
    public class BracingPlate : CustomSymbolDefinition , ICustomHgrWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "HgrOglaendParts,Ingr.SP3D.Content.Support.Symbols.BracingPlate"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "Height1", "Height1 of L Plate", 0)]
        public InputDouble m_Height1;
        [InputDouble(3, "Width1", "Width1 of L Plate", 0)]
        public InputDouble m_Width1;
        [InputDouble(4, "Height2", "Height2 of L Plate", 0)]
        public InputDouble m_Height2;
        [InputDouble(5, "Width2", "Width2 of L Plate", 0)]
        public InputDouble m_Width2;
        [InputDouble(6, "Thickness", "Thickness of L Plate", 0)]
        public InputDouble m_Thickness;
        [InputDouble(7, "P1xOffset", "Port1 xOffset of Plate", 0)]
        public InputDouble m_P1xOffset;
        [InputDouble(8, "P1yOffset", "Port1 yOffset of Plate", 0)]
        public InputDouble m_P1yOffset;
        [InputDouble(9, "P1zOffset", "Port1 zOffset of Plate", 0)]
        public InputDouble m_P1zOffset;
        [InputDouble(10, "P2xOffset", "Port2 xOffset of Plate", 0)]
        public InputDouble m_P2xOffset;
        [InputDouble(11, "P2yOffset", "Port2 yOffset of Plate", 0)]
        public InputDouble m_P2yOffset;
        [InputDouble(12, "P2zOffset", "Port2 zOffset of Plate", 0)]
        public InputDouble m_P2zOffset;

        #endregion

        #region "Definitions of Aspects and their outputs"

        // Physical Aspect 
        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Plate1", "Plate1")]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("Plate2", "Plate2")]
        public AspectDefinition m_PhysicalAspect;

        #endregion

        #region "Construction of outputs of all aspects"
        protected override void ConstructOutputs()
        {
            try
            {
                Part part = null;
                SymbolGeometryHelper symbolGeomHlpr = new SymbolGeometryHelper();
                SP3DConnection connection = default(SP3DConnection);
                connection = OccurrenceConnection;

                part = m_PartInput.Value as Part;
                double height1 = 0;
                double width1 = 0;
                double height2 = 0;
                double thickness = 0;
                double width2 = 0;

                height1 = m_Height1.Value;
                if (height1 < 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, OglaendPartLocalizer.GetString(OglaendPartSymbolResourceIDs.ErrInvalidArguments, "Height1 value should be greater than 0"));
                    return;
                }

                width1 = m_Width1.Value;
                if (width1 < 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, OglaendPartLocalizer.GetString(OglaendPartSymbolResourceIDs.ErrInvalidArguments, "Width1 value should be greater than 0"));
                    return;
                }

                height2 = m_Height2.Value;
                if (height2 < 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, OglaendPartLocalizer.GetString(OglaendPartSymbolResourceIDs.ErrInvalidArguments, "Height2 value should be greater than 0"));
                    return;
                }

                thickness = m_Thickness.Value;
                if (thickness < 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, OglaendPartLocalizer.GetString(OglaendPartSymbolResourceIDs.ErrInvalidArguments, "Thickness value should be greater than 0"));
                    return;
                }

                width2 = m_Width2.Value;
                if (width2 < 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, OglaendPartLocalizer.GetString(OglaendPartSymbolResourceIDs.ErrInvalidArguments, "Width2 value should be greater than 0"));
                    return;
                }

                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                //Initialize SymbolGeometryHelper. Set the active position and orientation 
                symbolGeomHlpr.ActivePosition = new Position(0, 0, 0);
                symbolGeomHlpr.SetOrientation(new Vector(0, 0, 1), new Vector(1, 0, 0));

                //Create Plate1
                Line3d plate1_1, plate1_2, plate1_3, plate1_4, plate1_5;
                Collection<ICurve> plate1Coll = new Collection<ICurve>();


                plate1_1 = new Line3d(new Position(-thickness / 2, 0, 0), new Position(-thickness / 2, width1, 0));
                plate1_2 = new Line3d(new Position(-thickness / 2, width1, 0), new Position(-thickness / 2, width2, height1 - height2));
                plate1_3 = new Line3d(new Position(-thickness / 2, width2, height1 - height2), new Position(-thickness / 2, width2, height1));
                plate1_4 = new Line3d(new Position(-thickness / 2, width2, height1), new Position(-thickness / 2, 0, height1));
                plate1_5 = new Line3d(new Position(-thickness / 2, 0, height1), new Position(-thickness / 2, 0, 0));

                plate1Coll.Add(plate1_1);
                plate1Coll.Add(plate1_2);
                plate1Coll.Add(plate1_3);
                plate1Coll.Add(plate1_4);
                plate1Coll.Add(plate1_5);

                ComplexString3d plate1ComplxString = new ComplexString3d(plate1Coll);
                Vector plate1Vector = new Vector(1, 0, 0);
                Projection3d plate1Projection = new Projection3d(connection, plate1ComplxString, plate1Vector, thickness, true);

                m_PhysicalAspect.Outputs["Plate1"] = plate1Projection;

                PropertyValueInt propPlateCount = (PropertyValueInt)part.GetPropertyValue("IJHgrPlateCount", "PlateCount");
                int plateCount = 0;
                double plateOffset = 0;

                if (propPlateCount.PropValue > 0)
                {
                    plateCount = (int)propPlateCount.PropValue;
                    PropertyValueDouble propPlateOffset = (PropertyValueDouble)part.GetPropertyValue("IJHgrPlateCount", "Offset");
                    plateOffset = (double)propPlateOffset.PropValue;
                }

                if (plateCount == 2)
                {
                    //Create Plate2
                    Collection<ICurve> plate2Coll = new Collection<ICurve>();
                    Line3d plate2_1, plate2_2, plate2_3, plate2_4, plate2_5;

                    plate2_1 = new Line3d(new Position(plateOffset + thickness / 2, 0, 0), new Position(plateOffset + thickness / 2, width1, 0));
                    plate2_2 = new Line3d(new Position(plateOffset + thickness / 2, width1, 0), new Position(plateOffset + thickness / 2, width2, height1 - height2));
                    plate2_3 = new Line3d(new Position(plateOffset + thickness / 2, width2, height1 - height2), new Position(plateOffset + thickness / 2, width2, height1));
                    plate2_4 = new Line3d(new Position(plateOffset + thickness / 2, width2, height1), new Position(plateOffset + thickness / 2, 0, height1));
                    plate2_5 = new Line3d(new Position(plateOffset + thickness / 2, 0, height1), new Position(plateOffset + thickness / 2, 0, 0));

                    plate2Coll.Add(plate2_1);
                    plate2Coll.Add(plate2_2);
                    plate2Coll.Add(plate2_3);
                    plate2Coll.Add(plate2_4);
                    plate2Coll.Add(plate2_5);

                    ComplexString3d plate2ComplxString = new ComplexString3d(plate2Coll);
                    Vector plate2Vector = new Vector(1, 0, 0);
                    Projection3d plate2Projection = new Projection3d(connection, plate2ComplxString, plate2Vector, thickness, true);

                    m_PhysicalAspect.Outputs["Plate2"] = plate2Projection;
                }

                double port1OffsetX = m_P1xOffset.Value;
                double port1OffsetY = m_P1yOffset.Value;
                double port1OffsetZ = m_P1zOffset.Value;
                double port2OffsetX = m_P2xOffset.Value;
                double port2OffsetY = m_P2yOffset.Value;
                double port2OffsetZ = m_P2zOffset.Value;

                Port port1 = new Port(connection, part, "Port1", new Position((plateOffset + thickness) / 2 + port1OffsetX, width2 + port1OffsetY, height1 + port1OffsetZ), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_PhysicalAspect.Outputs["Port1"] = port1;

                Port port2 = new Port(connection, part, "Port2", new Position((plateOffset + thickness) / 2 + port2OffsetX, port2OffsetY, height1 + port2OffsetZ), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_PhysicalAspect.Outputs["Port2"] = port2;
            }
            catch
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    Type myType = this.GetType();
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, OglaendPartLocalizer.GetString(OglaendPartSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of BracingPlate.cs " + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name));
                }
            }

        #endregion
        }
  #region ICustomHgrWeightCG Members
        void ICustomHgrWeightCG.WeightCG(SupportComponent supportComponent, ref double weight, ref double cogX, ref double cogY, ref double cogZ)
        {

            ////System WCG Attributes

            Part catalogPart = (Part)supportComponent.GetRelationship("madeFrom", "part").TargetObjects[0];

            try
            {
                weight = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryWeight")).PropValue;
            }
            catch
            {
                weight = 0;
            }
            //Center of Gravity
            try
            {
                cogX = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryCogX")).PropValue;
            }
            catch
            {
                cogX = 0;
            }
            try
            {
                cogY = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryCogY")).PropValue;
            }
            catch
            {
                cogY = 0;
            }
            try
            {
                cogZ = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryCogZ")).PropValue;
            }
            catch
            {
                cogZ = 0;
            }

        }
        #endregion
    }
}


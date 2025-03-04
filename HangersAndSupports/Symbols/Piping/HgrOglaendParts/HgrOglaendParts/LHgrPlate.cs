//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   LHgrPlate.cs
//   HgrOglaendParts,Ingr.SP3D.Content.Support.Symbols.LHgrPlate
//   Author       :  VSP
//   Creation Date:  29.May.2012
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   29.May.2012     VSP     Initial Creation
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
    public class LHgrPlate : CustomSymbolDefinition, ICustomHgrWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "HgrOglaendParts,Ingr.SP3D.Content.Support.Symbols.LHgrPlate"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "FlangeWidth", "Width of Flange of L Plate", 0)]
        public InputDouble m_FlangeWidth;
        [InputDouble(3, "WebWidth", "Width of Web of L Plate", 0)]
        public InputDouble m_WebWidth;
        [InputDouble(4, "Length", "Length of L Plate", 0.5)]
        public InputDouble m_Length;
        [InputDouble(5, "OverhangFlWidth", "Overhang Flange Width of L Plate", 0)]
        public InputDouble m_OverhangFlWidth;
        [InputDouble(6, "OverhangWebWidth", "Overhang Web Width of L Plate", 0)]
        public InputDouble m_OverhangWebWidth;
        [InputDouble(7, "OverhangFlLength", "Overhang Flange Length of L Plate", 0)]
        public InputDouble m_OverhangFlLength;
        [InputDouble(8, "OverhangWebLength", "Overhang Web Length of L Plate", 0)]
        public InputDouble m_OverhangWebLength;
        [InputDouble(9, "Thickness", "Thickness of L Plate", 0)]
        public InputDouble m_Thickness;
        [InputDouble(10, "StuctureIntPosZ", "StuctureInt Port Position along Z", 0)]
        public InputDouble m_StuctureIntPosZ;
        [InputDouble(11, "Surface1PosZ", "Surface1 Port Position along Z", 0)]
        public InputDouble m_Surface1PosZ;
        [InputDouble(12, "Surface2PosZ", "Surface2 Port Position along Z", 0)]
        public InputDouble m_Surface2PosZ;
        [InputDouble(13, "HP1PosX", "HPort1 xOffset of Plate", 0)]
        public InputDouble m_HP1xOffset;
        [InputDouble(14, "HP1PosY", "HPort1 yOffset of Plate", 0)]
        public InputDouble m_HP1yOffset;
        [InputDouble(15, "HP1PosZ", "HPort1 zOffset of Plate", 0)]
        public InputDouble m_HP1zOffset;
        [InputDouble(16, "HP2PosX", "HPort2 xOffset of Plate", 0)]
        public InputDouble m_HP2xOffset;
        [InputDouble(17, "HP2PosY", "HPort2 yOffset of Plate", 0)]
        public InputDouble m_HP2yOffset;
        [InputDouble(18, "HP2PosZ", "HPort2 zOffset of Plate", 0)]
        public InputDouble m_HP2zOffset;

        #endregion

        #region "Definitions of Aspects and their outputs"

        // Physical Aspect 
        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Flange", "Flange")]
        [SymbolOutput("Web", "Web")]
        [SymbolOutput("Structure", "Structure")]
        [SymbolOutput("StructureInt", "StructureInt")]
        [SymbolOutput("Hole1", "Hole1")]
        [SymbolOutput("Hole2", "Hole2")]
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
                double length = 0;
                double flangeWidth = 0;
                double webWidth = 0;
                double thickness = 0;
                double overhangFlWidth = m_OverhangFlWidth.Value;
                double overhangWebWidth = m_OverhangWebWidth.Value;
                double overhangFlLength = m_OverhangFlLength.Value;
                double overhangWebLength = m_OverhangWebLength.Value;
                double surface1PosZ = m_Surface1PosZ.Value;
                double surface2PosZ = m_Surface2PosZ.Value;

                length = m_Length.Value;
                if (length < 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, OglaendPartLocalizer.GetString(OglaendPartSymbolResourceIDs.ErrInvalidArguments, "Length value should be greater than 0"));
                    return;
                }

                flangeWidth = m_FlangeWidth.Value;
                if (flangeWidth < 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, OglaendPartLocalizer.GetString(OglaendPartSymbolResourceIDs.ErrInvalidArguments, "Flange Width value should be greater than 0"));
                    return;
                }

                webWidth = m_WebWidth.Value;
                if (webWidth < 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, OglaendPartLocalizer.GetString(OglaendPartSymbolResourceIDs.ErrInvalidArguments, "Web Width value should be greater than 0"));
                    return;
                }

                thickness = m_Thickness.Value;
                if (thickness < 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, OglaendPartLocalizer.GetString(OglaendPartSymbolResourceIDs.ErrInvalidArguments, "Thickness value should be greater than 0"));
                    return;
                }


                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                //Initialize SymbolGeometryHelper. Set the active position and orientation 
                symbolGeomHlpr.ActivePosition = new Position(0, 0, 0);
                symbolGeomHlpr.SetOrientation(new Vector(0, 0, 1), new Vector(1, 0, 0));

                //Create the Web
                Line3d webLine1, webLine2, webLine3, webLine4;
                Collection<ICurve> webColl = new Collection<ICurve>();

                webLine1 = new Line3d(new Position(0, 0, 0), new Position(0, 0, length));
                webLine2 = new Line3d(new Position(0, 0, length), new Position(0, webWidth, length));
                webLine3 = new Line3d(new Position(0, webWidth, length), new Position(0, webWidth, 0));
                webLine4 = new Line3d(new Position(0, webWidth, 0), new Position(0, 0, 0));

                ////Add the curves into ICurves Collection
                webColl.Add(webLine1);
                webColl.Add(webLine2);
                webColl.Add(webLine3);
                webColl.Add(webLine4);

                //Create the Flange
                Line3d flangeLine1, flangeLine2, flangeLine3, flangeLine4, flangeLine5, flangeLine6;
                Collection<ICurve> flangeColl = new Collection<ICurve>();
                flangeLine1 = new Line3d(new Position(thickness, 0, 0), new Position(thickness, 0, length));

                if (overhangFlWidth > 0 && overhangFlLength > 0)
                {
                    //Create the Lines 
                    flangeLine2 = new Line3d(new Position(thickness, 0, length), new Position((flangeWidth - thickness - overhangFlWidth), 0, length));
                    flangeLine3 = new Line3d(new Position((flangeWidth - thickness - overhangFlWidth), 0, length), new Position((flangeWidth - thickness - overhangFlWidth), 0, (length + overhangFlLength)));
                    flangeLine4 = new Line3d(new Position((flangeWidth - thickness - overhangFlWidth), 0, (length + overhangFlLength)), new Position((flangeWidth - thickness), 0, (length + overhangFlLength)));
                    flangeLine5 = new Line3d(new Position((flangeWidth - thickness), 0, (length + overhangFlLength)), new Position(flangeWidth - thickness, 0, 0));
                    flangeLine6 = new Line3d(new Position(flangeWidth - thickness, 0, 0), new Position(thickness, 0, 0));

                    //Add the curves into ICurves Collection
                    flangeColl.Add(flangeLine1);
                    flangeColl.Add(flangeLine2);
                    flangeColl.Add(flangeLine3);
                    flangeColl.Add(flangeLine4);
                    flangeColl.Add(flangeLine5);
                    flangeColl.Add(flangeLine6);
                }
                else
                {
                    flangeLine2 = new Line3d(new Position(thickness, 0, length), new Position(flangeWidth - thickness, 0, length));
                    flangeLine3 = new Line3d(new Position(flangeWidth - thickness, 0, length), new Position(flangeWidth - thickness, 0, 0));
                    flangeLine4 = new Line3d(new Position(flangeWidth - thickness, 0, 0), new Position(thickness, 0, 0));

                    //Add the curves into ICurves Collection
                    flangeColl.Add(flangeLine1);
                    flangeColl.Add(flangeLine2);
                    flangeColl.Add(flangeLine3);
                    flangeColl.Add(flangeLine4);
                }

                ComplexString3d webComplxString = new ComplexString3d(webColl);
                Vector webVector = new Vector(1, 0, 0);
                Projection3d webProjection = new Projection3d(connection, webComplxString, webVector, thickness, true);

                ComplexString3d flangeComplxString = new ComplexString3d(flangeColl);
                Vector flangeVector = new Vector(0, 1, 0);
                Projection3d flangeProjection = new Projection3d(connection, flangeComplxString, flangeVector, thickness, true);

                m_PhysicalAspect.Outputs["Web"] = webProjection;
                m_PhysicalAspect.Outputs["Flange"] = flangeProjection;

                double hport1OffsetX = m_HP1xOffset.Value;
                double hport1OffsetY = m_HP1yOffset.Value;
                double hport1OffsetZ = m_HP1zOffset.Value;
                double hport2OffsetX = m_HP2xOffset.Value;
                double hport2OffsetY = m_HP2yOffset.Value;
                double hport2OffsetZ = m_HP2zOffset.Value;
                double stuctureIntPosZ = m_StuctureIntPosZ.Value;

                //Add the ports
                if (overhangFlWidth > 0 && overhangFlLength > 0)
                {
                    Port oStructurePort = new Port(connection, part, "Structure", new Position(((flangeWidth - overhangFlWidth + thickness) / 2), (webWidth / 2), length), new Vector(1, 0, 0), new Vector(0, 0, 1));
                    m_PhysicalAspect.Outputs["Structure"] = oStructurePort;
                }
                else
                {
                    Port oStructurePort = new Port(connection, part, "Structure", new Position((flangeWidth / 2), (webWidth / 2), length), new Vector(1, 0, 0), new Vector(0, 0, 1));
                    m_PhysicalAspect.Outputs["Structure"] = oStructurePort;
                }

                Port oStructureIntPort = new Port(connection, part, "StructureInt", new Position(thickness, thickness, length / 2 + stuctureIntPosZ), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_PhysicalAspect.Outputs["StructureInt"] = oStructureIntPort;
                
                Port hole1Port = new Port(connection, part, "Hole1", new Position((flangeWidth / 2) + hport1OffsetX, 0 + hport1OffsetY, length - 0.005 + hport1OffsetZ), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_PhysicalAspect.Outputs["Hole1"] = hole1Port;
                
                Port hole2Port = new Port(connection, part, "Hole2", new Position(0 + hport2OffsetX, (webWidth / 2) + hport2OffsetY, length - 0.005 + hport2OffsetZ), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_PhysicalAspect.Outputs["Hole2"] = hole2Port;

            }
            catch
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    Type myType = this.GetType();
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, OglaendPartLocalizer.GetString(OglaendPartSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of LHgrPlate.cs " + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name));
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


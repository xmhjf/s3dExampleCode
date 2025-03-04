//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   Guide.cs
//   HS_OglaendParts,Ingr.SP3D.Support.Content.Symbols.Guide
//   Author       :  PVK	
//   Creation Date:  07.June.2012
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   07.June.2012     PVS     Initial Creation
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
    //Namespace of this class is Ingr.SP3D.Support.Content.Symbols
    //It is recommended that customers specify namespace of their symbols to be
    //<CompanyName>.SP3D.Content.<Specialization>.
    //It is also recommended that if customers want to change this symbol to suit their
    //requirements, they should change namespace/symbol name so the identity of the modified
    //symbol will be different from the one delivered by Intergraph.
    //-----------------------------------------------------------------------------------
    [CacheOption(CacheOptionType.NonCached)]
    [VariableOutputs]
    public class Guide : CustomSymbolDefinition, ICustomHgrWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "HS_OglaendParts,Ingr.SP3D.Support.Content.Symbols.Guide"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_oPartInput;
        [InputDouble(2, "Width1", "Width of Plate", 0)]
        public InputDouble m_dWidth;
        [InputDouble(3, "Length1", "Length of Plate", 0)]
        public InputDouble m_dLength;
        [InputDouble(4, "Thickness1", "Thickness of Plate", 0)]
        public InputDouble m_dThickness;
        [InputDouble(5, "P1xOffset", "Port1 xOffset of Plate", 0)]
        public InputDouble m_dP1xOffset;
        [InputDouble(6, "P1yOffset", "Port1 yOffset of Plate", 0)]
        public InputDouble m_dP1yOffset;
        [InputDouble(7, "P1zOffset", "Port1 zOffset of Plate", 0)]
        public InputDouble m_dP1zOffset;
        [InputDouble(8, "P2xOffset", "Port2 xOffset of Plate", 0)]
        public InputDouble m_dP2xOffset;
        [InputDouble(9, "P2yOffset", "Port2 yOffset of Plate", 0)]
        public InputDouble m_dP2yOffset;
        [InputDouble(10, "P2zOffset", "Port2 zOffset of Plate", 0)]
        public InputDouble m_dP2zOffset;
        [InputDouble(11, "LengthE",  "E", 0)]
        public InputDouble m_dE;
        [InputDouble(12, "LengthF", "F", 0)]
        public InputDouble m_dF;
        [InputDouble(13, "LengthG", "G", 0)]
        public InputDouble m_dG;
        [InputDouble(14, "LengthH", "H", 0)]
        public InputDouble m_dH;
    
        //[InputDouble(5, "Offset1", "Offset1", 0)]
        //public InputDouble m_dOffset1;
        //[InputDouble(6, "Offset2", "Offset2", 0)]
        //public InputDouble m_dOffset2;
        #endregion

        #region "Definitions of Aspects and their outputs"
        
        // Physical Aspect 
        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Plate", "Plate")]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("Structure", "Structure")]
        public AspectDefinition m_oPhysicalAspect;

        #endregion

        #region "Construction of outputs of all aspects"
        protected override void ConstructOutputs()
        {
            try
            {
                Part oPart = null;
                SymbolGeometryHelper oSymbolGeomHlpr = new SymbolGeometryHelper();
                SP3DConnection oConnection = default(SP3DConnection);
                oConnection = OccurrenceConnection;
                //SupportComponent supportComponent = OccurrenceConnection;
                oPart = m_oPartInput.Value as Part;
                double dLength = m_dLength.Value;
                double dWidth = m_dWidth.Value;
                double dThickness = m_dThickness.Value;
                // double d = 0.106;
                double e = m_dE.Value;
                double f = m_dF.Value;
                double g = m_dG.Value;
                double h = m_dH.Value;

                //MessageBox.Show("HI");

                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                //Initialize SymbolGeometryHelper. Set the active position and orientation 
                oSymbolGeomHlpr.ActivePosition = new Position(0, 0, 0);
                oSymbolGeomHlpr.SetOrientation(new Vector(0, 0, 1), new Vector(1, 0, 0));

                //Create the Web
                Line3d oPlateLine1, oPlateLine2, oPlateLine3, oPlateLine4, oPlateLine5, oPlateLine6, oPlateLine7, oPlateLine8, oPlateLine9, oPlateLine10, oPlateLine11, oPlateLine12;
                //Collection<ICurve> oWebColl1 = new Collection<ICurve>();
                Collection<ICurve> oPlateColl1 = new Collection<ICurve>();

                oPlateLine1 = new Line3d(new Position(0, 0, 0), new Position(0, 0, dLength));
                oPlateLine2 = new Line3d(new Position(0, 0, dLength), new Position(dWidth, 0, dLength));
                oPlateLine3 = new Line3d(new Position(dWidth, 0, dLength), new Position(dWidth, 0, dLength - (h + g)));
                oPlateLine4 = new Line3d(new Position(dWidth, 0, dLength - (h + g)), new Position(dWidth - f, 0, dLength - (h + g)));
                oPlateLine5 = new Line3d(new Position(dWidth - f, 0, dLength - (h + g)), new Position(dWidth - f, 0, dLength - h));
                oPlateLine6 = new Line3d(new Position(dWidth - f, 0, dLength - h), new Position(dWidth - f - e, 0, dLength - h));
                oPlateLine7 = new Line3d(new Position(dWidth - f - e, 0, dLength - h), new Position(dWidth - f - e, 0, h));
                oPlateLine8 = new Line3d(new Position(dWidth - f - e, 0, h), new Position(dWidth - f, 0, h));
                oPlateLine9 = new Line3d(new Position(dWidth - f, 0, h), new Position(dWidth - f, 0, h + g));
                oPlateLine10 = new Line3d(new Position(dWidth - f, 0, h + g), new Position(dWidth, 0, h + g));
                oPlateLine11 = new Line3d(new Position(dWidth, 0, h + g), new Position(dWidth, 0, 0));
                oPlateLine12 = new Line3d(new Position(dWidth, 0, 0), new Position(0, 0, 0));


                ////Add the curves into ICurves Collection
                oPlateColl1.Add(oPlateLine1);
                oPlateColl1.Add(oPlateLine2);
                oPlateColl1.Add(oPlateLine3);
                oPlateColl1.Add(oPlateLine4);
                oPlateColl1.Add(oPlateLine5);
                oPlateColl1.Add(oPlateLine6);
                oPlateColl1.Add(oPlateLine7);
                oPlateColl1.Add(oPlateLine8);
                oPlateColl1.Add(oPlateLine9);
                oPlateColl1.Add(oPlateLine10);
                oPlateColl1.Add(oPlateLine11);
                oPlateColl1.Add(oPlateLine12);



                ComplexString3d oPlateComplxString = new ComplexString3d(oPlateColl1);
                Vector PlateVector = new Vector(0, 1, 0);
                Projection3d oPlateProjection = new Projection3d(oConnection, oPlateComplxString, PlateVector, dThickness, true);


                double dPort1OffsetX = m_dP1xOffset.Value;
                double dPort1OffsetY = m_dP1yOffset.Value;
                double dPort1OffsetZ = m_dP1zOffset.Value;
                double dPort2OffsetX = m_dP2xOffset.Value;
                double dPort2OffsetY = m_dP2yOffset.Value;
                double dPort2OffsetZ = m_dP2zOffset.Value;

                m_oPhysicalAspect.Outputs["Plate"] = oPlateProjection;

                Port oStructurePort = new Port(oConnection, oPart, "Structure", new Position(dWidth - (e + f), dThickness / 2, dLength / 2), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_oPhysicalAspect.Outputs["Structure"] = oStructurePort;

                Port oPort1 = new Port(oConnection, oPart, "Port1", new Position(dWidth - (e + f) + dPort1OffsetX, dPort1OffsetY, dLength / 2 + dPort1OffsetZ), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_oPhysicalAspect.Outputs["Port1"] = oPort1;

                Port oPort2 = new Port(oConnection, oPart, "Port2", new Position(dWidth - (e + f) + dPort2OffsetX, dThickness + dPort2OffsetY, dLength / 2 + dPort2OffsetZ), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_oPhysicalAspect.Outputs["Port2"] = oPort2;

            }
            catch (Exception oExc) //General Unhandled exception 
            {
                throw oExc;
            }
        }
        #endregion

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


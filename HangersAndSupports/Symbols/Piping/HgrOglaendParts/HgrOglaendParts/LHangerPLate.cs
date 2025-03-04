//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   LHangerPlate.cs
//   HS_OglaendParts,Ingr.SP3D.Content.Support.Symbols.LHangerPlate
//   Author       :  PVK
//   Creation Date:  18.SEP.2013
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
    //<CompanyName>.SP3D.Content.<Specialization>.
    //It is also recommended that if customers want to change this symbol to suit their
    //requirements, they should change namespace/symbol name so the identity of the modified
    //symbol will be different from the one delivered by Intergraph.
    //-----------------------------------------------------------------------------------
    [VariableOutputs]
    public class LHangerPlate : CustomSymbolDefinition, ICustomHgrWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "HS_OglaendParts,Ingr.SP3D.Content.Support.Symbols.LHgrPlate"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_oPartInput;
        [InputDouble(2, "FlangeWidth", "Width of Flange of L Plate", 0)]
        public InputDouble m_dFlangeWidth;
        [InputDouble(3, "WebWidth", "Width of Web of L Plate", 0)]
        public InputDouble m_dWebWidth;
        [InputDouble(4, "Length", "Length of L Plate", 0.5)]
        public InputDouble m_dLength;
        [InputDouble(5, "Thickness", "Thickness of L Plate", 0)]
        public InputDouble m_dThickness;
        [InputDouble(6, "SP1PosX", "SteelPort1 xOffset of Plate", 0)]
        public InputDouble m_dSP1xOffset;
        [InputDouble(7, "SP1PosY", "SteelPort1 yOffset of Plate", 0)]
        public InputDouble m_dSP1yOffset;
        [InputDouble(8, "SP1PosZ", "SteelPort1 zOffset of Plate", 0)]
        public InputDouble m_dSP1zOffset;
        [InputDouble(9, "SP2PosX", "SteelPort2 xOffset of Plate", 0)]
        public InputDouble m_dSP2xOffset;
        [InputDouble(10, "SP2PosY", "SteelPort2 yOffset of Plate", 0)]
        public InputDouble m_dSP2yOffset;
        [InputDouble(11, "SP2PosZ", "SteelPort2 zOffset of Plate", 0)]
        public InputDouble m_dSP2zOffset;

        #endregion

        #region "Definitions of Aspects and their outputs"

        // Physical Aspect 
        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("LSection", "Lsection")]
        [SymbolOutput("Steel1", "Steel1")]
        [SymbolOutput("Steel2", "Steel2")]
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

                oPart = m_oPartInput.Value as Part;
                double dLength = m_dLength.Value;
                double dFlangeWidth = m_dFlangeWidth.Value;
                double dWebWidth = m_dWebWidth.Value;
                double dThickness = m_dThickness.Value;


                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                //Initialize SymbolGeometryHelper. Set the active position and orientation 
                oSymbolGeomHlpr.ActivePosition = new Position(0, 0, 0);
                oSymbolGeomHlpr.SetOrientation(new Vector(0, 0, 1), new Vector(1, 0, 0));

                //Create the Web
                Line3d oLine1, oLine2, oLine3, oLine4, oLine5, oLine6;
                Collection<ICurve> oLColl = new Collection<ICurve>();

                //oLine1 = new Line3d(new Position(0, 0, 0), new Position(dWebWidth, 0, 0));
                //oLine2 = new Line3d(new Position(dWebWidth, 0, 0), new Position(dWebWidth, 0, dFlangeWidth));
                //oLine3 = new Line3d(new Position(dWebWidth, 0, dFlangeWidth), new Position(dWebWidth - dThickness, 0, dFlangeWidth));
                //oLine4 = new Line3d(new Position(dWebWidth - dThickness, 0, dFlangeWidth), new Position(dWebWidth - dThickness, 0, dThickness));
                //oLine5 = new Line3d(new Position(dWebWidth - dThickness, 0, dThickness), new Position(0, 0, dThickness));
                //oLine6 = new Line3d(new Position(0, 0, dThickness), new Position(0, 0, 0));

                oLine1 = new Line3d(new Position(0, 0, 0), new Position(0, 0, dWebWidth));
                oLine2 = new Line3d(new Position(0, 0, dWebWidth), new Position(dFlangeWidth, 0, dWebWidth));
                oLine3 = new Line3d(new Position(dFlangeWidth, 0, dWebWidth), new Position(dFlangeWidth, 0, dWebWidth - dThickness));
                oLine4 = new Line3d(new Position(dFlangeWidth, 0, dWebWidth - dThickness), new Position(dThickness, 0, dWebWidth - dThickness));
                oLine5 = new Line3d(new Position(dThickness, 0, dWebWidth - dThickness), new Position(dThickness, 0, 0));
                oLine6 = new Line3d(new Position(dThickness, 0, 0), new Position(0, 0, 0));

                ////Add the curves into ICurves Collection
                oLColl.Add(oLine1);
                oLColl.Add(oLine2);
                oLColl.Add(oLine3);
                oLColl.Add(oLine4);
                oLColl.Add(oLine5);
                oLColl.Add(oLine6);



                ComplexString3d oLComplxString = new ComplexString3d(oLColl);
                Vector LVector = new Vector(0, 1, 0);
                Projection3d oLProjection = new Projection3d(oConnection, oLComplxString, LVector, dLength, true);

                m_oPhysicalAspect.Outputs["LSection"] = oLProjection;

                double dSteelPort1OffsetX = m_dSP1xOffset.Value;
                double dSteelPort1OffsetY = m_dSP1yOffset.Value;
                double dSteelPort1OffsetZ = m_dSP1zOffset.Value;
                double dSteelPort2OffsetX = m_dSP2xOffset.Value;
                double dSteelPort2OffsetY = m_dSP2yOffset.Value;
                double dSteelPort2OffsetZ = m_dSP2zOffset.Value;


                Port oSteelPort = new Port(oConnection, oPart, "Steel1", new Position(dFlangeWidth / 2 + dSteelPort1OffsetX, (dLength / 2) + dSteelPort1OffsetY, dWebWidth + dSteelPort1OffsetZ), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_oPhysicalAspect.Outputs["Steel1"] = oSteelPort;


                Port oSteel2Port = new Port(oConnection, oPart, "Steel2", new Position(dSteelPort2OffsetX, (dLength / 2) + dSteelPort2OffsetY, dWebWidth / 2 + dSteelPort2OffsetZ), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_oPhysicalAspect.Outputs["Steel2"] = oSteel2Port;

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


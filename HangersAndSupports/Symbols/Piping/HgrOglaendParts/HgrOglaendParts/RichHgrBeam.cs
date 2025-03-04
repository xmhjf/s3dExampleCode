//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   RichHgrBeam.cs
//   HgrOglaendParts,Ingr.SP3D.Content.Support.Symbols.RichHgrBeam
//   Author       :  Chethan
//   Creation Date:  23-01-2015
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   7.June.2012     Ramya     Initial Creation
//   23-01-2015      Chethan  CR-CP-241198  Convert all Oglaend Content to .NET 
//   28-04-2015      PVK	  Resolve Coverity issues found in April
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
using System;
using System.Linq;

using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.Common.Exceptions;
using Ingr.SP3D.Structure.Middle.Services;
using Ingr.SP3D.Support.Middle;
using System.Collections.ObjectModel;

namespace Ingr.SP3D.Content.Support.Symbols
{
    //-----------------------------------------------------------------------------------
    //Namespace of this class is Ingr.SP3D.Content.Support.Symbols
    //It is recommended that customers specify namespace of their symbols to be
    //<CompanyName>.SP3D.Support.Content.<Specialization>.
    //It is also recommended that if customers want to change this symbol to suit their
    //requirements, they should change namespace/symbol name so the identity of the modified
    //symbol will be different from the one delivered by Intergraph.
    //-----------------------------------------------------------------------------------
    [VariableOutputs]
    public class RichHgrBeam : ConnectionComponentDefinition, ICustomHgrBOMDescription, ICustomHgrWeightCG
    {
        private const double LINEAR_TOLERANCE = 0.0000001;
        private const double DENSITY = 7849;  //kg/m^3
        private const long CUTBACK_STEEL = 1;
        private const long SNIPPED_STEEL = 2;
        private const long NUM_PORTS = 30;
        HangerBeamInputs oHgrBeamInput;
        CutbackSteelInputs oCutbackSteelInput;
        SnipSteelInputs oSnipSteelInput;

        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "HS_3DAPI_RichHgrBeam,Ingr.SP3D.Content.Support.Symbols.RichHgrBeam"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs

        [InputCatalogPart(1)]
        public InputCatalogPart m_oPartInput;
        [InputDouble(2, "BeginOverLength", "BeginOverLength of Beam", 0)]
        public InputDouble m_dBeginOverLength;
        [InputDouble(3, "EndOverLength", "EndOverLength of the Beam", 0)]
        public InputDouble m_dEndOverLength;
        [InputDouble(4, "Length", "Length of the Beam", 0)]
        public InputDouble m_dLength;
        [InputDouble(5, "CP1", "CP1 of BeginCapPort", 1)]
        public InputDouble m_dCP1;
        [InputDouble(6, "CP2", "CP2 of EndCapPort", 1)]
        public InputDouble m_dCP2;
        [InputDouble(7, "CP3", "CP3 of BeginFace Port", 5)]
        public InputDouble m_dCP3;
        [InputDouble(8, "CP4", "CP4 of EndCapFace Port", 5)]
        public InputDouble m_dCP4;
        [InputDouble(9, "CP5", "CP5 of Neutral Port", 5)]
        public InputDouble m_dCP5;
        [InputDouble(10, "CP6", "CP6 of Route port", 8)]
        public InputDouble m_dCP6;
        [InputDouble(11, "BeginCutbackAnchorPoint", "BeginCutbackAnchorPoint of the Beam", 2)]
        public InputDouble m_dBeginCutbackAnchorPoint;
        [InputDouble(12, "EndCutbackAnchorPoint", "EndCutbackAnchorPoint of the Beam", 2)]
        public InputDouble m_dEndCutbackAnchorPoint;
        [InputDouble(13, "BeginCapXOffset", "BeginCapXOffset of Beam", 0)]
        public InputDouble m_dBeginCapXOffset;
        [InputDouble(14, "BeginCapYOffset", "BeginCapYOffset of the Beam", 0)]
        public InputDouble m_dBeginCapYOffset;
        [InputDouble(15, "BeginCapRotZ", "BeginCapRotZ of the Beam", 0)]
        public InputDouble m_dBeginCapRotZ;
        [InputDouble(16, "EndCapXOffset", "EndCapXOffset of the Beam", 0)]
        public InputDouble m_dEndCapXOffset;
        [InputDouble(17, "EndCapYOffset", "EndCapYOffset of the Beam", 0)]
        public InputDouble m_dEndCapYOffset;
        [InputDouble(18, "EndCapRotZ", "EndCapRotZ of the Beam", 1)]
        public InputDouble m_dEndCapRotZ;
        [InputDouble(19, "FlexPortXOffset", "FlexPortXOffset of the Beam", 0)]
        public InputDouble m_dFlexPortXOffset;
        [InputDouble(20, "FlexPortYOffset", "FlexPortYOffset of the Beam", 0)]
        public InputDouble m_dFlexPortYOffset;
        [InputDouble(21, "FlexPortZOffset", "FlexPortZOffset of the Beam", 0)]
        public InputDouble m_dFlexPortZOffset;
        [InputDouble(22, "FlexPortRotX", "FlexPortRotX of the Beam", 0)]
        public InputDouble m_dFlexPortRotX;
        [InputDouble(23, "FlexPortRotY", "FlexPortRotY of the Beam", 0)]
        public InputDouble m_dFlexPortRotY;
        [InputDouble(24, "FlexPortRotZ", "FlexPortRotZ of the Beam", 0)]
        public InputDouble m_dFlexPortRotZ;
        [InputDouble(25, "CutbackBeginAngle", "CutbackBeginAngle of the Beam", 0)]
        public InputDouble m_dCutbackBeginAngle;
        [InputDouble(26, "CutbackEndAngle", "CutbackEndAngle of the Beam", 0)]
        public InputDouble m_dCutbackEndAngle;
        [InputString(27, "MaterialGrade", "MaterialGrade of the Beam", "C20")]
        public InputString m_sMaterialGrade;
        [InputString(28, "MaterialType", "MaterialType of the Beam", "Concrete")]
        public InputString m_sMaterialType;
        [InputDouble(29, "CoatingType", "CoatingType", 3)]
        public InputDouble m_sCoatingType;
        [InputDouble(30, "CoatingRequirement", "CoatingRequirement", 3)]
        public InputDouble m_sCoatingRequirement;
        [InputDouble(31, "HgrBeamType", "HgrBeamType", 1)]
        public InputDouble m_dHgrBeamType;
        [InputDouble(32, "CutbackBeginAngle1", "CutbackBeginAngle1", 0)]
        public InputDouble m_dCutbackBeginAngle1;
        [InputDouble(33, "CutbackBeginAngle2", "CutbackBeginAngle2", 0)]
        public InputDouble m_dCutbackBeginAngle2;
        [InputDouble(34, "CutbackEndAngle1", "CutbackEndAngle1", 0)]
        public InputDouble m_dCutbackEndAngle1;
        [InputDouble(35, "CutbackEndAngle2", "CutbackEndAngle2", 0)]
        public InputDouble m_dCutbackEndAngle2;
        [InputDouble(36, "BeginCutbackAnchorPoint1", "BeginCutbackAnchorPoint1", 0)]
        public InputDouble m_dBeginCutbackAnchorPoint1;
        [InputDouble(37, "BeginCutbackAnchorPoint2", "BeginCutbackAnchorPoint2", 0)]
        public InputDouble m_dBeginCutbackAnchorPoint2;
        [InputDouble(38, "EndCutbackAnchorPoint1", "EndCutbackAnchorPoint1", 0)]
        public InputDouble m_dEndCutbackAnchorPoint1;
        [InputDouble(39, "EndCutbackAnchorPoint2", "EndCutbackAnchorPoint2", 0)]
        public InputDouble m_dEndCutbackAnchorPoint2;
        [InputDouble(40, "BeginOffsetAlongFlange", "BeginOffsetAlongFlange", 0)]
        public InputDouble m_dBeginOffsetAlongFlange;
        [InputDouble(41, "BeginOffsetAlongWeb", "BeginOffsetAlongWeb", 0)]
        public InputDouble m_dBeginOffsetAlongWeb;
        [InputDouble(42, "EndOffsetAlongFlange", "EndOffsetAlongFlange", 0)]
        public InputDouble m_dEndOffsetAlongFlange;
        [InputDouble(43, "EndOffsetAlongWeb", "EndOffsetAlongWeb", 0)]
        public InputDouble m_dEndOffsetAlongWeb;
        [InputDouble(44, "FacePortOrient", "FacePortOrient", 0)]
        public InputDouble m_dFacePortOrient;
        [InputDouble(45, "VarLength", "VarLength", 0.5)]
        public InputDouble m_dVarLength;
        [InputDouble(46, "EndFlexPortXOffset", "EndFlexPortXOffset of the Beam", 0)]
        public InputDouble m_dEndFlexPortXOffset;
        [InputDouble(47, "EndFlexPortYOffset", "EndFlexPortYOffset of the Beam", 0)]
        public InputDouble m_dEndFlexPortYOffset;
        [InputDouble(48, "EndFlexPortZOffset", "EndFlexPortZOffset of the Beam", 0)]
        public InputDouble m_dEndFlexPortZOffset;
        [InputDouble(49, "EndFlexPortRotX", "EndFlexPortRotX of the Beam", 0)]
        public InputDouble m_dEndFlexPortRotX;
        [InputDouble(50, "EndFlexPortRotY", "EndFlexPortRotY of the Beam", 0)]
        public InputDouble m_dEndFlexPortRotY;
        [InputDouble(51, "EndFlexPortRotZ", "EndFlexPortRotZ of the Beam", 0)]
        public InputDouble m_dEndFlexPortRotZ;
        [InputDouble(52, "CP7", "CP7 of EndFlex Port", 8)]
        public InputDouble m_dCP7;

        #endregion

        #region "Definitions of Aspects and their outputs"

        // Physical Aspect 
        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("BeginCap", "Begin Cap")]
        [SymbolOutput("EndCap", "End Cap")]
        [SymbolOutput("BeginFace", "Begin Face")]
        [SymbolOutput("EndFace", "End Face")]
        [SymbolOutput("Neutral", "Neutral")]
        [SymbolOutput("BeginFlex", "Begin Flex")]
        [SymbolOutput("BeginCapSurface", "Begin Cap Surface")]
        [SymbolOutput("EndCapSurface", "End Cap Surface")]
        [SymbolOutput("EndFlex", "End Flex")]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("Port3", "Port3")]
        [SymbolOutput("Port4", "Port4")]
        [SymbolOutput("Port5", "Port5")]
        [SymbolOutput("Port6", "Port6")]
        [SymbolOutput("Port7", "Port7")]
        [SymbolOutput("Port8", "Port8")]
        [SymbolOutput("Port9", "Port9")]
        [SymbolOutput("Port10", "Port10")]
        [SymbolOutput("Port11", "Port11")]
        [SymbolOutput("Port12", "Port12")]
        [SymbolOutput("Port13", "Port13")]
        [SymbolOutput("Port14", "Port14")]
        [SymbolOutput("Port15", "Port15")]
        [SymbolOutput("Port16", "Port16")]
        [SymbolOutput("Port17", "Port17")]
        [SymbolOutput("Port18", "Port18")]
        [SymbolOutput("Port19", "Port19")]
        [SymbolOutput("Port20", "Port20")]
        [SymbolOutput("Port21", "Port21")]
        [SymbolOutput("Port22", "Port22")]
        [SymbolOutput("Port23", "Port23")]
        [SymbolOutput("Port24", "Port24")]
        [SymbolOutput("Port25", "Port25")]
        [SymbolOutput("Port26", "Port26")]
        [SymbolOutput("Port27", "Port27")]
        [SymbolOutput("Port28", "Port28")]
        [SymbolOutput("Port29", "Port29")]
        [SymbolOutput("Port30", "Port30")]
        public AspectDefinition m_oPhysicalAspect;

        #endregion

        #region "Construction of outputs of all aspects"
        protected override void ConstructOutputs()
        {
            try
            {
                
                //Begin Cap Port Orientation
                double dXDirX1 = 0, dXDirY1 = 0, dXDirZ1 = 0, dZDirX1 = 0, dZDirY1 = 0, dZDirZ1 = 0;

                //End Cap Port Orientation
                double dXDirX2 = 0, dXDirY2 = 0, dXDirZ2 = 0, dZDirX2 = 0, dZDirY2 = 0, dZDirZ2 = 0;

                //Begin Face Port Orientation
                double dXDirX3 = 0, dXDirY3 = 0, dXDirZ3 = 0, dZDirX3 = 0, dZDirY3 = 0, dZDirZ3 = 0;

                //End Face Port Orientation
                double dXDirX4 = 0, dXDirY4 = 0, dXDirZ4 = 0, dZDirX4 = 0, dZDirY4 = 0, dZDirZ4 = 0;

                //Route Port Orientation
                double dXDirX6 = 0, dXDirY6 = 0, dXDirZ6 = 0, dZDirX6 = 0, dZDirY6 = 0, dZDirZ6 = 0;

                //Endflex Port Orientation
                double dXDirX7 = 0, dXDirY7 = 0, dXDirZ7 = 0, dZDirX7 = 0, dZDirY7 = 0, dZDirZ7 = 0;
                
                double dSnipCBBeginAngle = 0, dSnipCBEndAngle = 0;
                AnchorPoint lSnipBeginCBAnchorPt = AnchorPoint.BottomCenter, lSnipEndCBAnchorPt = AnchorPoint.TopCenter;
                SymbolWarningException oWarningCollection = new SymbolWarningException();
                SymbolGeometryHelper oSymbolGeomHlpr = new SymbolGeometryHelper();
                SP3DConnection oConnection = OccurrenceConnection;

                Part oSectionPart = (Part)m_oPartInput.Value;
                double dBeginOverLength = m_dBeginOverLength.Value;
                double dEndOverLength = m_dEndOverLength.Value;
                long lCardinalPt1 = (long)m_dCP1.Value;
                long lCardinalPt2 = (long)m_dCP2.Value;
                long lCardinalPt3 = (long)m_dCP3.Value;
                long lCardinalPt4 = (long)m_dCP4.Value;
                long lCardinalPt5 = (long)m_dCP5.Value;
                long lCardinalPt6 = (long)m_dCP6.Value;
                AnchorPoint lBeginCutbackAnchorPoint = (AnchorPoint)m_dBeginCutbackAnchorPoint.Value;
                AnchorPoint lEndCutbackAnchorPoint = (AnchorPoint)m_dEndCutbackAnchorPoint.Value;
                double dBeginCapXOffset = m_dBeginCapXOffset.Value;
                double dBeginCapYOffset = m_dBeginCapYOffset.Value;
                double dBeginCapRotZ = m_dBeginCapRotZ.Value;
                double dEndCapXOffset = m_dEndCapXOffset.Value;
                double dEndCapYOffset = m_dEndCapYOffset.Value;
                double dEndCapRotZ = m_dEndCapRotZ.Value;
                double dFlexPortXOffset = m_dFlexPortXOffset.Value;
                double dFlexPortYOffset = m_dFlexPortYOffset.Value;
                double dFlexPortZOffset = m_dFlexPortZOffset.Value;
                double dFlexPortRotX = m_dFlexPortRotX.Value;
                double dFlexPortRotY = m_dFlexPortRotY.Value;
                double dFlexPortRotZ = m_dFlexPortRotZ.Value;
                double dCutbackBeginAngle = m_dCutbackBeginAngle.Value;
                double dCutbackEndAngle = m_dCutbackEndAngle.Value;
                long bIsCutback = (long)m_dHgrBeamType.Value;
                double dBeginAlongFlangeAngle = m_dCutbackBeginAngle1.Value;
                double dBeginAlongWebAngle = m_dCutbackBeginAngle2.Value;
                double dEndAlongFlangeAngle = m_dCutbackEndAngle1.Value;
                double dEndAlongWebAngle = m_dCutbackEndAngle2.Value;
                AnchorPoint lBeginAlongFlangeAnchrPnt = (AnchorPoint)m_dBeginCutbackAnchorPoint1.Value;
                AnchorPoint lBeginAlongWebAnchrPnt = (AnchorPoint)m_dBeginCutbackAnchorPoint2.Value;
                AnchorPoint lEndAlongFlangeAnchrPnt = (AnchorPoint)m_dEndCutbackAnchorPoint1.Value;
                AnchorPoint lEndAlongWebAnchrPnt = (AnchorPoint)m_dEndCutbackAnchorPoint2.Value;
                double dBeginAlongFlangeOffset = m_dBeginOffsetAlongFlange.Value;
                double dBeginAlongWebOffset = m_dBeginOffsetAlongWeb.Value;
                double dEndAlongFlangeOffset = m_dEndOffsetAlongFlange.Value;
                double dEndAlongWebOffset = m_dEndOffsetAlongWeb.Value;
                long lFacePortOrient = (long)m_dFacePortOrient.Value;
                long lCutLength = (long)m_dVarLength.Value;
                double dEndFlexPortXOffset = m_dEndFlexPortXOffset.Value;
                double dEndFlexPortYOffset = m_dEndFlexPortYOffset.Value;
                double dEndFlexPortZOffset = m_dEndFlexPortZOffset.Value;
                double dEndFlexPortRotX = m_dEndFlexPortRotX.Value;
                double dEndFlexPortRotY = m_dEndFlexPortRotY.Value;
                double dEndFlexPortRotZ = m_dEndFlexPortRotZ.Value;
                long lCardinalPt7 = (long)m_dCP7.Value;
                double dCPOffsetx = 0, dCPOffsety = 0;

                double dLength;
                if(m_dLength.Value == 0)
                    dLength= 0.5;
                else
                    dLength = m_dLength.Value;
                //=================================================
                // Construction of Physical Aspect
                //=================================================   
                oSymbolGeomHlpr.ActivePosition = new Position(0, 0, -dLength);
                oSymbolGeomHlpr.SetOrientation(new Vector(0, 0, 1), new Vector(1, 0, 0));
                
                //Get the Cross Section from the Relationship between the HgrPart and the CrossSection
                RelationCollection oHgrRelation;
                CrossSection oCrossSection;
                CrossSectionServices oCrossSectionServices = new CrossSectionServices();
                
                oHgrRelation = oSectionPart.GetRelationship("HgrCrossSection", "CrossSection");
                oCrossSection = (CrossSection)oHgrRelation.TargetObjects.First();
                
                //Get Section type and standard
                string sSectionType = oCrossSection.Name;
                string sSectionStd = (oCrossSection.CrossSectionClass.Name);
                double dWidth, dDepth, dFlangeThickness, dWebThickness;
                GetSectionData(oSectionPart, out dWidth, out dFlangeThickness, out dWebThickness, out dDepth);

                
                
                #region "Is Cutback"
                if (bIsCutback == CUTBACK_STEEL)
                {
                    if (dCutbackBeginAngle == 0 && dCutbackEndAngle == 0)
                    {

                        dXDirX3 = 1;
                        dXDirY3 = 0;
                        dXDirZ3 = 0;
                        dZDirX3 = 0;
                        dZDirY3 = 0;
                        dZDirZ3 = 1;
                        dXDirX4 = 1;
                        dXDirY4 = 0;
                        dXDirZ4 = 0;
                        dZDirX4 = 0;
                        dZDirY4 = 0;
                        dZDirZ4 = 1;
                    }
                    //else if (lBeginCutbackAnchorPoint == 2 || lBeginCutbackAnchorPoint == 8)
                    else if (lBeginCutbackAnchorPoint == AnchorPoint.BottomCenter || lBeginCutbackAnchorPoint == AnchorPoint.TopCenter)
                    {
                        dXDirX3 = 1;
                        dXDirY3 = 0;
                        dXDirZ3 = 0;

                        dZDirX3 = 0;
                        dZDirY3 = Math.Sin(dCutbackBeginAngle);
                        dZDirZ3 = Math.Cos(dCutbackBeginAngle);

                        if (lEndCutbackAnchorPoint == AnchorPoint.BottomCenter || lEndCutbackAnchorPoint == AnchorPoint.TopCenter)
                            //if (lEndCutbackAnchorPoint == 2 || lEndCutbackAnchorPoint == 8)
                        {
                            dXDirX4 = 1;
                            dXDirY4 = 0;
                            dXDirZ4 = 0;

                            dZDirX4 = 0;
                            dZDirY4 = Math.Sin(dCutbackEndAngle);
                            dZDirZ4 = Math.Cos(dCutbackEndAngle);
                        }
                        else
                        {
                            dXDirX4 = Math.Cos(dCutbackEndAngle);
                            dXDirY4 = 0;
                            dXDirZ4 = Math.Sin(dCutbackEndAngle);

                            dZDirX4 = -Math.Sin(dCutbackEndAngle);
                            dZDirY4 = 0;
                            dZDirZ4 = Math.Cos(dCutbackEndAngle);
                        }
                    }
                    else if (lBeginCutbackAnchorPoint == AnchorPoint.MiddleLeft || lBeginCutbackAnchorPoint == AnchorPoint.MiddleRight)
                        //else if (lBeginCutbackAnchorPoint == 4 || lBeginCutbackAnchorPoint == 6)
                    {
                        dXDirX3 = Math.Cos(dCutbackBeginAngle);
                        dXDirY3 = 0;
                        dXDirZ3 = Math.Sin(dCutbackBeginAngle);

                        dZDirX3 = -Math.Sin(dCutbackBeginAngle);
                        dZDirY3 = 0;
                        dZDirZ3 = Math.Cos(dCutbackBeginAngle);


                        if (lEndCutbackAnchorPoint == AnchorPoint.BottomCenter || lEndCutbackAnchorPoint == AnchorPoint.TopCenter)
                            //if (lEndCutbackAnchorPoint == 2 || lEndCutbackAnchorPoint == 8)
                        {
                            dXDirX4 = 1;
                            dXDirY4 = 0;
                            dXDirZ4 = 0;

                            dZDirX4 = 0;
                            dZDirY4 = Math.Sin(dCutbackEndAngle);
                            dZDirZ4 = Math.Cos(dCutbackEndAngle);
                        }
                        else
                        {

                            dXDirX4 = Math.Cos(dCutbackEndAngle);
                            dXDirY4 = 0;
                            dXDirZ4 = Math.Sin(dCutbackEndAngle);

                            dZDirX4 = -Math.Sin(dCutbackEndAngle);
                            dZDirY4 = 0;
                            dZDirZ4 = Math.Cos(dCutbackEndAngle);

                        }
                    }
                }
                #endregion
                #region "Is Snip"
                else
                {
                    #region "Set Angle"
                    if (dBeginAlongFlangeAngle < 0 || dBeginAlongWebAngle > 0 || dEndAlongFlangeAngle > 0 || dEndAlongWebAngle < 0)
                    {
                        if (dBeginAlongFlangeAngle < 0)
                            dBeginAlongFlangeAngle = Math.Abs(dBeginAlongFlangeAngle);
                        if (dBeginAlongWebAngle > 0)
                            dBeginAlongWebAngle = -dBeginAlongWebAngle;
                        if (dEndAlongFlangeAngle > 0)
                            dEndAlongFlangeAngle = -dEndAlongFlangeAngle;
                        if (dEndAlongWebAngle < 0)
                            dEndAlongWebAngle = Math.Abs(dEndAlongWebAngle);
                    }
                    #endregion
                    #region "Set SnipAngle"
                    if (lFacePortOrient == 1)
                    {
                        dSnipCBBeginAngle = -dBeginAlongFlangeAngle;
                        dSnipCBEndAngle = -dEndAlongFlangeAngle;
                        lSnipBeginCBAnchorPt = lBeginAlongFlangeAnchrPnt;
                        lSnipEndCBAnchorPt = lEndAlongFlangeAnchrPnt;
                    }
                    else if (lFacePortOrient == 2)
                    {
                        dSnipCBBeginAngle = dBeginAlongWebAngle;
                        dSnipCBEndAngle = dEndAlongWebAngle;
                        lSnipBeginCBAnchorPt = lBeginAlongWebAnchrPnt;
                        lSnipEndCBAnchorPt = lEndAlongWebAnchrPnt;
                    }
                    #endregion
                    #region "Set FacePort Orientation for 2"
                    //set the rotation for the face port
                    if (lFacePortOrient == 2)
                    {
                        if (lCardinalPt3 == 1 || lCardinalPt3 == 2 || lCardinalPt3 == 3 || lCardinalPt3 == 11)
                        {
                            if (CompareDoubleGreaterthan(dBeginAlongWebOffset, LINEAR_TOLERANCE))
                            {
                                dXDirX3 = 1;
                                dXDirY3 = 0;
                                dXDirZ3 = 0;

                                dZDirX3 = 0;
                                dZDirY3 = 0;
                                dZDirZ3 = 1;
                            }
                            else
                            {
                                dXDirX3 = 1;
                                dXDirY3 = 0;
                                dXDirZ3 = 0;

                                dZDirX3 = 0;
                                dZDirY3 = Math.Sin(dSnipCBBeginAngle);
                                dZDirZ3 = Math.Cos(dSnipCBBeginAngle);
                            }
                        }
                        else if (lCardinalPt3 == 4 || lCardinalPt3 == 5 || lCardinalPt3 == 6 || lCardinalPt3 == 15)
                        {
                            if (CompareDoubleGreaterthan(dBeginAlongWebOffset, dDepth / 2 + LINEAR_TOLERANCE))
                            {
                                dXDirX3 = 1;
                                dXDirY3 = 0;
                                dXDirZ3 = 0;

                                dZDirX3 = 0;
                                dZDirY3 = 0;
                                dZDirZ3 = 1;
                            }
                            else
                            {
                                dXDirX3 = 1;
                                dXDirY3 = 0;
                                dXDirZ3 = 0;

                                dZDirX3 = 0;
                                dZDirY3 = Math.Sin(dSnipCBBeginAngle);
                                dZDirZ3 = Math.Cos(dSnipCBBeginAngle);
                            }
                        }
                        else if (lCardinalPt3 == 10 || lCardinalPt3 == 12 || lCardinalPt3 == 13)
                        {
                            if (CompareDoubleGreaterthan(dBeginAlongWebOffset, dDepth / 2 + LINEAR_TOLERANCE))
                            {
                                dXDirX3 = 1;
                                dXDirY3 = 0;
                                dXDirZ3 = 0;

                                dZDirX3 = 0;
                                dZDirY3 = 0;
                                dZDirZ3 = 1;
                            }
                            else
                            {
                                dXDirX3 = 1;
                                dXDirY3 = 0;
                                dXDirZ3 = 0;

                                dZDirX3 = 0;
                                dZDirY3 = Math.Sin(dSnipCBBeginAngle);
                                dZDirZ3 = Math.Cos(dSnipCBBeginAngle);
                            }
                        }
                        else if (lCardinalPt3 == 7 || lCardinalPt3 == 8 || lCardinalPt3 == 9 || lCardinalPt3 == 14)
                        {
                            dXDirX3 = 1;
                            dXDirY3 = 0;
                            dXDirZ3 = 0;

                            dZDirX3 = 0;
                            dZDirY3 = Math.Sin(dSnipCBBeginAngle);
                            dZDirZ3 = Math.Cos(dSnipCBBeginAngle);
                        }

                        if (lCardinalPt4 == 1 || lCardinalPt4 == 2 || lCardinalPt4 == 3 || lCardinalPt4 == 11)
                        {
                            if (CompareDoubleGreaterthan(dEndAlongWebOffset, LINEAR_TOLERANCE))
                            {
                                dXDirX4 = 1;
                                dXDirY4 = 0;
                                dXDirZ4 = 0;

                                dZDirX4 = 0;
                                dZDirY4 = 0;
                                dZDirZ4 = 1;
                            }
                            else
                            {
                                dXDirX4 = 1;
                                dXDirY4 = 0;
                                dXDirZ4 = 0;

                                dZDirX4 = 0;
                                dZDirY4 = Math.Sin(dSnipCBEndAngle);
                                dZDirZ4 = Math.Cos(dSnipCBEndAngle);
                            }
                        }
                        else if (lCardinalPt4 == 4 || lCardinalPt4 == 5 || lCardinalPt4 == 6 || lCardinalPt4 == 15)
                        {
                            if (CompareDoubleGreaterthan(dEndAlongWebOffset, dDepth / 2 + LINEAR_TOLERANCE))
                            {
                                dXDirX4 = 1;
                                dXDirY4 = 0;
                                dXDirZ4 = 0;

                                dZDirX4 = 0;
                                dZDirY4 = 0;
                                dZDirZ4 = 1;
                            }
                            else
                            {
                                dXDirX4 = 1;
                                dXDirY4 = 0;
                                dXDirZ4 = 0;

                                dZDirX4 = 0;
                                dZDirY4 = Math.Sin(dSnipCBEndAngle);
                                dZDirZ4 = Math.Cos(dSnipCBEndAngle);
                            }

                        }
                        else if (lCardinalPt4 == 10 || lCardinalPt4 == 12 || lCardinalPt4 == 13)
                        {
                            if ((CompareDoubleGreaterthan(dEndAlongWebOffset, dDepth / 2 + LINEAR_TOLERANCE)))
                            {
                                dXDirX4 = 1;
                                dXDirY4 = 0;
                                dXDirZ4 = 0;

                                dZDirX4 = 0;
                                dZDirY4 = 0;
                                dZDirZ4 = 1;
                            }
                            else
                            {
                                dXDirX4 = 1;
                                dXDirY4 = 0;
                                dXDirZ4 = 0;

                                dZDirX4 = 0;
                                dZDirY4 = Math.Sin(dSnipCBEndAngle);
                                dZDirZ4 = Math.Cos(dSnipCBEndAngle);
                            }
                        }
                        else if (lCardinalPt4 == 7 || lCardinalPt4 == 8 || lCardinalPt4 == 9 || lCardinalPt4 == 14)
                        {
                            dXDirX4 = 1;
                            dXDirY4 = 0;
                            dXDirZ4 = 0;

                            dZDirX4 = 0;
                            dZDirY4 = Math.Sin(dSnipCBEndAngle);
                            dZDirZ4 = Math.Cos(dSnipCBEndAngle);
                        }
                    }
                    #endregion
                    #region "Set FacePort Orientation for 1"
                    else if (lFacePortOrient == 1)
                    {
                        if (lCardinalPt3 == 1 || lCardinalPt3 == 4 || lCardinalPt3 == 7 || lCardinalPt3 == 12)
                        {
                            if (CompareDoubleGreaterthan(dBeginAlongFlangeOffset, LINEAR_TOLERANCE))
                            {
                                dXDirX3 = 1;
                                dXDirY3 = 0;
                                dXDirZ3 = 0;

                                dZDirX3 = 0;
                                dZDirY3 = 0;
                                dZDirZ3 = 1;
                            }
                            else
                            {
                                dXDirX3 = Math.Cos(dSnipCBBeginAngle);
                                dXDirY3 = 0;
                                dXDirZ3 = -Math.Sin(dSnipCBBeginAngle);

                                dZDirX3 = Math.Sin(dSnipCBBeginAngle);
                                dZDirY3 = 0;
                                dZDirZ3 = Math.Cos(dSnipCBBeginAngle);
                            }
                        }
                        else if (lCardinalPt3 == 2 || lCardinalPt3 == 5 || lCardinalPt3 == 8)
                        {
                            if (CompareDoubleGreaterthan(dBeginAlongFlangeOffset, dWidth / 2 + LINEAR_TOLERANCE))
                            {
                                dXDirX3 = 1;
                                dXDirY3 = 0;
                                dXDirZ3 = 0;

                                dZDirX3 = 0;
                                dZDirY3 = 0;
                                dZDirZ3 = 1;
                            }
                            else
                            {
                                dXDirX3 = Math.Cos(dSnipCBBeginAngle);
                                dXDirY3 = 0;
                                dXDirZ3 = -Math.Sin(dSnipCBBeginAngle);

                                dZDirX3 = Math.Sin(dSnipCBBeginAngle);
                                dZDirY3 = 0;
                                dZDirZ3 = Math.Cos(dSnipCBBeginAngle);
                            }
                        }
                        else if (lCardinalPt3 == 11 || lCardinalPt3 == 10 || lCardinalPt3 == 14 || lCardinalPt3 == 15)
                        {
                            if (CompareDoubleGreaterthan(dBeginAlongFlangeOffset, dWidth / 4 + LINEAR_TOLERANCE))
                            {
                                dXDirX3 = 1;
                                dXDirY3 = 0;
                                dXDirZ3 = 0;

                                dZDirX3 = 0;
                                dZDirY3 = 0;
                                dZDirZ3 = 1;
                            }
                            else
                            {
                                dXDirX3 = Math.Cos(dSnipCBBeginAngle);
                                dXDirY3 = 0;
                                dXDirZ3 = -Math.Sin(dSnipCBBeginAngle);

                                dZDirX3 = Math.Sin(dSnipCBBeginAngle);
                                dZDirY3 = 0;
                                dZDirZ3 = Math.Cos(dSnipCBBeginAngle);
                            }
                        }
                        else if (lCardinalPt3 == 3 || lCardinalPt3 == 6 || lCardinalPt3 == 9 || lCardinalPt3 == 13)
                        {
                            dXDirX3 = Math.Cos(dSnipCBBeginAngle);
                            dXDirY3 = 0;
                            dXDirZ3 = -Math.Sin(dSnipCBBeginAngle);

                            dZDirX3 = Math.Sin(dSnipCBBeginAngle);
                            dZDirY3 = 0;
                            dZDirZ3 = Math.Cos(dSnipCBBeginAngle);
                        }

                        if (lCardinalPt4 == 1 || lCardinalPt4 == 4 || lCardinalPt4 == 7 || lCardinalPt4 == 12)
                        {
                            if (CompareDoubleGreaterthan(dEndAlongFlangeOffset, LINEAR_TOLERANCE))
                            {
                                dXDirX4 = 1;
                                dXDirY4 = 0;
                                dXDirZ4 = 0;

                                dZDirX4 = 0;
                                dZDirY4 = 0;
                                dZDirZ4 = 1;
                            }
                            else
                            {
                                dXDirX4 = Math.Cos(dSnipCBEndAngle);
                                dXDirY4 = 0;
                                dXDirZ4 = -Math.Sin(dSnipCBEndAngle);

                                dZDirX4 = Math.Sin(dSnipCBEndAngle);
                                dZDirY4 = 0;
                                dZDirZ4 = Math.Cos(dSnipCBEndAngle);
                            }
                        }
                        else if (lCardinalPt4 == 2 || lCardinalPt4 == 5 || lCardinalPt4 == 8)
                        {
                            if (CompareDoubleGreaterthan(dEndAlongFlangeOffset, dWidth / 2 + LINEAR_TOLERANCE))
                            {
                                dXDirX4 = 1;
                                dXDirY4 = 0;
                                dXDirZ4 = 0;

                                dZDirX4 = 0;
                                dZDirY4 = 0;
                                dZDirZ4 = 1;
                            }
                            else
                            {
                                dXDirX4 = Math.Cos(dSnipCBEndAngle);
                                dXDirY4 = 0;
                                dXDirZ4 = -Math.Sin(dSnipCBEndAngle);

                                dZDirX4 = Math.Sin(dSnipCBEndAngle);
                                dZDirY4 = 0;
                                dZDirZ4 = Math.Cos(dSnipCBEndAngle);
                            }
                        }
                        else if (lCardinalPt4 == 11 || lCardinalPt4 == 10 || lCardinalPt4 == 14 || lCardinalPt4 == 15)
                        {
                            if (CompareDoubleGreaterthan(dEndAlongFlangeOffset, dWidth / 4 + LINEAR_TOLERANCE))
                            {
                                dXDirX4 = 1;
                                dXDirY4 = 0;
                                dXDirZ4 = 0;

                                dZDirX4 = 0;
                                dZDirY4 = 0;
                                dZDirZ4 = 1;
                            }
                            else
                            {
                                dXDirX4 = Math.Cos(dSnipCBEndAngle);
                                dXDirY4 = 0;
                                dXDirZ4 = -Math.Sin(dSnipCBEndAngle);

                                dZDirX4 = Math.Sin(dSnipCBEndAngle);
                                dZDirY4 = 0;
                                dZDirZ4 = Math.Cos(dSnipCBEndAngle);
                            }
                        }
                        else if (lCardinalPt4 == 3 || lCardinalPt4 == 6 || lCardinalPt4 == 9 || lCardinalPt4 == 13)
                        {
                            dXDirX4 = Math.Cos(dSnipCBEndAngle);
                            dXDirY4 = 0;
                            dXDirZ4 = -Math.Sin(dSnipCBEndAngle);

                            dZDirX4 = Math.Sin(dSnipCBEndAngle);
                            dZDirY4 = 0;
                            dZDirZ4 = Math.Cos(dSnipCBEndAngle);
                        }
                    }
                    #endregion
                }
                #endregion

                ReadOnlyCollection<BusinessObject> oPorts;

                oHgrBeamInput = new HangerBeamInputs();
                oHgrBeamInput.BeginOverLength = dBeginOverLength;
                oHgrBeamInput.CardinalPoint = 1;
                oHgrBeamInput.EndOverLength = dEndOverLength;
                oHgrBeamInput.Length = dLength;
                oHgrBeamInput.Part = oSectionPart;
                oHgrBeamInput.Density = 0.25; //default value
    
                oCutbackSteelInput = new CutbackSteelInputs();
                oCutbackSteelInput.BeginAnchorPoint = lBeginCutbackAnchorPoint;
                oCutbackSteelInput.BeginOverLength = dBeginOverLength;
                oCutbackSteelInput.CutbackBeginAngle = dCutbackBeginAngle;
                oCutbackSteelInput.CutbackEndAngle = dCutbackEndAngle;
                oCutbackSteelInput.Density = 0.25; // default value
                oCutbackSteelInput.EndAnchorPoint = lEndCutbackAnchorPoint;
                oCutbackSteelInput.EndOverLength = dEndOverLength;
                oCutbackSteelInput.Length = dLength;
                oCutbackSteelInput.Part = oSectionPart;

                oSnipSteelInput = new SnipSteelInputs();
                oSnipSteelInput.BeginFlangeAnchorPoint = lBeginAlongFlangeAnchrPnt;
                oSnipSteelInput.BeginOffsetAlongFlange = dBeginAlongFlangeOffset;
                oSnipSteelInput.BeginOffsetAlongWeb = dBeginAlongWebOffset;
                oSnipSteelInput.BeginOverLength = dBeginOverLength;
                oSnipSteelInput.BeginWebAnchorPoint = lBeginAlongWebAnchrPnt;
                oSnipSteelInput.Density = 0.25; //default value
                oSnipSteelInput.EndFlangeAnchorPoint = lEndAlongFlangeAnchrPnt;
                oSnipSteelInput.EndOffsetAlongFlange = dEndAlongFlangeOffset;
                oSnipSteelInput.EndOffsetAlongWeb = dEndAlongWebOffset;
                oSnipSteelInput.EndOverLength = dEndOverLength;
                oSnipSteelInput.EndWebAnchorPoint = lEndAlongWebAnchrPnt;
                oSnipSteelInput.Length = dLength;
                oSnipSteelInput.Part = oSectionPart;
                oSnipSteelInput.SnipBeginAngleAlongFlange = dBeginAlongFlangeAngle;
                oSnipSteelInput.SnipBeginAngleAlongWeb = dBeginAlongWebAngle;
                oSnipSteelInput.SnipEndAngleAlongFlange = dEndAlongFlangeAngle;
                oSnipSteelInput.SnipEndAngleAlongWeb = dEndAlongWebAngle;


                if (bIsCutback == CUTBACK_STEEL)
                {
                    if (sSectionType == "HSSC" || sSectionType == "PIPE" || sSectionType == "CS")
                    {
                        oPorts = CreateConnectionComponentPorts(oHgrBeamInput);  
                        // oGenericHelper.ConnectionComponentPortCollection(oSectionPart, 1, dBeginOverLength, dLength, dEndOverLength);
                    }
                    else
                    {
                        //double dVal;
                        //oGenericHelper.GetDataByRule("", oSectionPart, out dVal);
                       
                        oPorts = CreateConnectionComponentPorts(oCutbackSteelInput);
                            //oGenericHelper.ConnectionComponentCutbackPortCollection(oSectionPart, dBeginOverLength, dLength, dEndOverLength, (int)lBeginCutbackAnchorPoint, (int)lEndCutbackAnchorPoint, dCutbackBeginAngle, dCutbackEndAngle);
                    }

                }
                else
                {
                    if (sSectionType == "L")
                    {
                        
                        oPorts = CreateConnectionComponentPorts(oSnipSteelInput);
                            //oGenericHelper.ConnectionComponentSnipPortCollection(oSectionPart, dBeginOverLength, dLength, dEndOverLength, (int)lBeginAlongFlangeAnchrPnt, (int)lEndAlongFlangeAnchrPnt, (int)lBeginAlongWebAnchrPnt, (int)lEndAlongWebAnchrPnt, dBeginAlongFlangeOffset, dBeginAlongWebOffset, dEndAlongFlangeAngle, dEndAlongWebOffset, dBeginAlongFlangeAngle, dBeginAlongWebAngle, dEndAlongFlangeAngle, dEndAlongWebAngle);
                    }
                    else
                    {

                        //PF_EventHandler "Snip functionality is supported only for L section", Err, MODULE, "GetConnectionPortCollForSnip", False
                        return;
                    }

                }
                #region "Set BeginCap"
                //Last two ports are reserved for CapSurfaces
                long lPortCount = oPorts.Count - 2;

                if (lPortCount > 0)
                {
                    for (int iIndex = 0; iIndex < lPortCount; iIndex++)
                    {
                        m_oPhysicalAspect.Outputs["Port" + (iIndex+1)] = oPorts[iIndex];
                    }
                }
                //Now add the Cap Surfaces
                m_oPhysicalAspect.Outputs["BeginCapSurface"] = oPorts[(int)lPortCount];
                m_oPhysicalAspect.Outputs["EndCapSurface"] = oPorts[(int)lPortCount + 1];

                oCrossSectionServices.GetCardinalPointOffset(oCrossSection, (int)lCardinalPt1,out dCPOffsetx,out dCPOffsety);//  'Begin Cap
                //MessageBox.Show("CP1 x: " + dCPOffsetx + "CP1 y: " + dCPOffsety);

                //Begin Cap and End Cap ports location offsets
                double dX, dY, dZ;

                //Route port location offsets
                double dX6, dY6, dRouteZ6;

                //Route port location offsets
                double dX7, dY7, dRouteZ7;

                //Begin Face and End Face ports location offsets
                double dX1, dY1, dBFZ1, dX2, dY2, dEFZ2;

                dX = dCPOffsetx + dBeginCapXOffset;
                dY = dCPOffsety + dBeginCapYOffset;
                dZ = 0;

                if (dBeginCapRotZ == 0)
                {
                    dXDirX1 = 1;
                    dXDirY1 = 0;
                    dXDirZ1 = 0;
                    dZDirX1 = 0;
                    dZDirY1 = 0;
                    dZDirZ1 = 1;
                }
                else
                {
                    dXDirX1 = Math.Cos(dBeginCapRotZ);
                    dXDirY1 = Math.Sin(dBeginCapRotZ);
                    dXDirZ1 = 0;
                    dZDirX1 = 0;
                    dZDirY1 = 0;
                    dZDirZ1 = 1;
                }
                //BeginCap Port
                Port oPort1 = new Port(oConnection, oSectionPart, "BeginCap");
                oPort1.Origin = new Position(dX, dY, dZ);
                oPort1.SetOrientation(new Vector(dXDirX1, dXDirY1, dXDirZ1), new Vector(dZDirX1, dZDirY1, dZDirZ1));
                m_oPhysicalAspect.Outputs["BeginCap"] = oPort1;

                #endregion

                #region "Set EndCap Ports"

                oCrossSectionServices.GetCardinalPointOffset(oCrossSection, (int)lCardinalPt1, out dCPOffsetx, out dCPOffsety);//  'End Cap

                dX = dCPOffsetx + dEndCapXOffset;
                dY = dCPOffsety + dEndCapYOffset;
                dZ = 0;

                if (dEndCapRotZ == 0)
                {
                    dXDirX2 = 1;
                    dXDirY2 = 0;
                    dXDirZ2 = 0;
                    dZDirX2 = 0;
                    dZDirY2 = 0;
                    dZDirZ2 = 1;
                }
                else
                {
                    dXDirX2 = Math.Cos(dEndCapRotZ);
                    dXDirY2 = Math.Sin(dEndCapRotZ);
                    dXDirZ2 = 0;
                    dZDirX2 = 0;
                    dZDirY2 = 0;
                    dZDirZ2 = 1;
                }

                //EndCap Port
                Port oPort2 = new Port(oConnection, oSectionPart, "EndCap");
                oPort2.Origin = new Position(dX, dY, dZ+dLength);
                oPort2.SetOrientation(new Vector(dXDirX2, dXDirY2, dXDirZ2), new Vector(dZDirX2, dZDirY2, dZDirZ2));
                m_oPhysicalAspect.Outputs["EndCap"] = oPort2;

                #endregion

                #region "Set BeginFace"
                oCrossSectionServices.GetCardinalPointOffset(oCrossSection,(int) lCardinalPt3,out dCPOffsetx,out dCPOffsety);//    'Begin Face

                dX1 = dCPOffsetx;
                dY1 = dCPOffsety;

                if (bIsCutback == CUTBACK_STEEL)
                    dBFZ1 = BeginFaceZCoOrdinateForCutback(oSectionPart, lBeginCutbackAnchorPoint, lCardinalPt3, dCutbackBeginAngle);
                else
                    dBFZ1 = BeginFaceZCoOrdinateForSnip(oSectionPart, lCardinalPt3, dSnipCBBeginAngle, lSnipBeginCBAnchorPt, lFacePortOrient, dBeginAlongFlangeOffset, dBeginAlongWebOffset);

                if (sSectionType == "HSSC" || sSectionType == "PIPE" || sSectionType == "CS")
                {
                    dBFZ1 = 0;
                    dXDirX3 = 1;
                    dXDirY3 = 0;
                    dXDirZ3 = 0;
                    dZDirX3 = 0;
                    dZDirY3 = 0;
                    dZDirZ3 = 1;
                }

                //BeginFace Port
                Port oPort3 = new Port(oConnection, oSectionPart, "BeginFace");
                oPort3.Origin = new Position(dX1, dY1, -dBeginOverLength + dBFZ1);
                oPort3.SetOrientation(new Vector(dXDirX3, dXDirY3, dXDirZ3), new Vector(dZDirX3, dZDirY3, dZDirZ3));
                m_oPhysicalAspect.Outputs["BeginFace"] = oPort3;

                #endregion

                #region "Set EndFace"
                oCrossSectionServices.GetCardinalPointOffset(oCrossSection,(int) lCardinalPt4,out dCPOffsetx,out dCPOffsety);//    'EndFace

                dX2 = dCPOffsetx;
                dY2 = dCPOffsety;

                if (bIsCutback == CUTBACK_STEEL)
                    dEFZ2 = EndFaceZCoordForCutback(oSectionPart, lEndCutbackAnchorPoint, lCardinalPt4, dCutbackEndAngle);
                else
                    dEFZ2 = EndFaceZCoordForSnip(oSectionPart, lCardinalPt4, dSnipCBEndAngle, lSnipEndCBAnchorPt, lFacePortOrient, dEndAlongFlangeOffset, dEndAlongWebOffset);


                if (sSectionType == "HSSC" || sSectionType == "PIPE" || sSectionType == "CS")
                {
                    dEFZ2 = 0;
                    dXDirX4 = 1;
                    dXDirY4 = 0;
                    dXDirZ4 = 0;
                    dZDirX4 = 0;
                    dZDirY4 = 0;
                    dZDirZ4 = 1;
                }

                //EndFace Port
                Port oPort4 = new Port(oConnection, oSectionPart, "EndFace");
                oPort4.Origin = new Position(dX2, dY2, dLength + dEndOverLength + dEFZ2);
                oPort4.SetOrientation(new Vector(dXDirX4, dXDirY4, dXDirZ4), new Vector(dZDirX4, dZDirY4, dZDirZ4));
                m_oPhysicalAspect.Outputs["EndFace"] = oPort4;
                #endregion

                #region "Neutral Port"
                //Neutral
                oCrossSectionServices.GetCardinalPointOffset(oCrossSection,(int) lCardinalPt5,out dCPOffsetx,out dCPOffsety);//    'Neutral

                dX = dCPOffsetx;
                dY = dCPOffsety;
                dZ = dLength / 2;

                Port oPort5 = new Port(oConnection, oSectionPart, "Neutral");
                oPort5.Origin = new Position(dX, dY, dZ);
                oPort5.SetOrientation(new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_oPhysicalAspect.Outputs["Neutral"] = oPort5;

                #endregion

                #region "BeginFlex Port"
                //BeginFlex
                oCrossSectionServices.GetCardinalPointOffset(oCrossSection, (int)lCardinalPt6, out dCPOffsetx, out dCPOffsety);

                dX6 = dCPOffsetx + dFlexPortXOffset;
                dY6 = dCPOffsety + dFlexPortYOffset;
                dRouteZ6 = dFlexPortZOffset;

                if (dFlexPortRotX == 0 && dFlexPortRotY == 0 && dFlexPortRotZ == 0)
                {
                    dXDirX6 = 1;
                    dXDirY6 = 0;
                    dXDirZ6 = 0;
                    dZDirX6 = 0;
                    dZDirY6 = 0;
                    dZDirZ6 = 1;
                }
                else
                {
                    Matrix4X4 transMatrix = new Matrix4X4();
                    Vector tempVactor = new Vector();
                    transMatrix.SetIdentity();


                    if (dFlexPortRotX != 0)
                    {
                        tempVactor.Set(1, 0, 0);
                        transMatrix.Rotate(dFlexPortRotX, tempVactor);
                    }

                    if (dFlexPortRotY != 0)
                    {
                        tempVactor.Set(0, 1, 0);
                        transMatrix.Rotate(dFlexPortRotY, tempVactor);
                    }

                    if (dFlexPortRotZ != 0)
                    {
                        tempVactor.Set(0, 0, 1);
                        transMatrix.Rotate(dFlexPortRotZ, tempVactor);
                    }
                    dXDirX6 = transMatrix.GetIndexValue(0);
                    dXDirY6 = transMatrix.GetIndexValue(1);
                    dXDirZ6 = transMatrix.GetIndexValue(2);
                    dZDirX6 = transMatrix.GetIndexValue(8);
                    dZDirY6 = transMatrix.GetIndexValue(9);
                    dZDirZ6 = transMatrix.GetIndexValue(10);
                }

                //BeginFlex

                Port oPort6 = new Port(oConnection, oSectionPart, "BeginFlex");
                oPort6.Origin = new Position(dX6, dY6, dRouteZ6);
                oPort6.SetOrientation(new Vector(dXDirX6, dXDirY6, dXDirZ6), new Vector(dZDirX6, dZDirY6, dZDirZ6));
                m_oPhysicalAspect.Outputs["BeginFlex"] = oPort6;

                #endregion

                #region "EndFlex Port"
                //EndFlex
                oCrossSectionServices.GetCardinalPointOffset(oCrossSection, (int)lCardinalPt7, out dCPOffsetx, out dCPOffsety);

                dX7 = dCPOffsetx + dEndFlexPortXOffset;
                dY7 = dCPOffsety + dEndFlexPortYOffset;
                dRouteZ7 = dLength + dEndFlexPortZOffset;

                if (dEndFlexPortRotX == 0 && dEndFlexPortRotY == 0 && dEndFlexPortRotZ == 0)
                {
                    dXDirX7 = 1;
                    dXDirY7 = 0;
                    dXDirZ7 = 0;
                    dZDirX7 = 0;
                    dZDirY7 = 0;
                    dZDirZ7 = 1;
                }
                else
                {
                    Matrix4X4 transMatrix1 = new Matrix4X4();
                    Vector tempVactor1 = new Vector();
                    transMatrix1.SetIdentity();


                    if (dEndFlexPortRotX != 0)
                    {
                        tempVactor1.Set(1, 0, 0);
                        transMatrix1.Rotate(dEndFlexPortRotX, tempVactor1);
                    }

                    if (dEndFlexPortRotY != 0)
                    {
                        tempVactor1.Set(0, 1, 0);
                        transMatrix1.Rotate(dEndFlexPortRotY, tempVactor1);
                    }

                    if (dEndFlexPortRotZ != 0)
                    {
                        tempVactor1.Set(0, 0, 1);
                        transMatrix1.Rotate(dEndFlexPortRotZ, tempVactor1);
                    }
                    dXDirX7 = transMatrix1.GetIndexValue(0);
                    dXDirY7 = transMatrix1.GetIndexValue(1);
                    dXDirZ7 = transMatrix1.GetIndexValue(2);
                    dZDirX7 = transMatrix1.GetIndexValue(8);
                    dZDirY7 = transMatrix1.GetIndexValue(9);
                    dZDirZ7 = transMatrix1.GetIndexValue(10);
                }

                //EndFlex
                Port oPort7 = new Port(oConnection, oSectionPart, "EndFlex");
                oPort7.Origin = new Position(dX7, dY7, dRouteZ7);
                oPort7.SetOrientation(new Vector(dXDirX7, dXDirY7, dXDirZ7), new Vector(dZDirX7, dZDirY7, dZDirZ7));
                m_oPhysicalAspect.Outputs["EndFlex"] = oPort7;

                #endregion

            }
            catch (Exception oEx)
            {
                throw oEx;
            }
        }
        #endregion

        private double BeginFaceZCoOrdinateForCutback(Part oPart, AnchorPoint lBeginCutbackAnchorPoint, long lCardinalPoint, double dBeginCBAngle)
        {
            double dBeginFaceZ = 0;
            try
            {
                AnchorPoint lBeginCBAnchorPoint = lBeginCutbackAnchorPoint;
                long lBeginFaceCardinalPt = lCardinalPoint;
                double dBeginCutbackAngle = dBeginCBAngle;
                double dWidth, dFlangeThickness, dWebThickness, dDepth;
                double dZ3, dZ5;

                GetSectionData(oPart, out dWidth, out dFlangeThickness, out dWebThickness, out dDepth);

                dZ3 = dWidth * Math.Tan(-dBeginCBAngle);    //for Anchor Points 4 and 6
                dZ5 = dDepth * Math.Tan(dBeginCBAngle);    //for Anchor Points 2 and 8

                //Begin Face and End Face Port Offsets
                if (lBeginCutbackAnchorPoint == AnchorPoint.MiddleLeft  || lBeginCutbackAnchorPoint == AnchorPoint.MiddleRight)
                    //if (lBeginCutbackAnchorPoint == 4 || lBeginCutbackAnchorPoint == 6)
                {
                    if (lBeginFaceCardinalPt == 1 || lBeginFaceCardinalPt == 4 || lBeginFaceCardinalPt == 7 || lBeginFaceCardinalPt == 12)
                    {
                        //if (lBeginCutbackAnchorPoint == 4)
                        if (lBeginCutbackAnchorPoint == AnchorPoint.MiddleLeft)
                            dBeginFaceZ = 0;
                        else if (lBeginCutbackAnchorPoint == AnchorPoint.MiddleRight)
                            //else if (lBeginCutbackAnchorPoint == 6)
                            dBeginFaceZ = -dZ3;
                    }
                    else if (lBeginFaceCardinalPt == 2 || lBeginFaceCardinalPt == 5 || lBeginFaceCardinalPt == 8)
                    {
                        if (lBeginCutbackAnchorPoint == AnchorPoint.MiddleLeft)
                            //if (lBeginCutbackAnchorPoint == 4)
                            dBeginFaceZ = -(dZ3 / 2);
                        else if (lBeginCutbackAnchorPoint == AnchorPoint.MiddleRight)
                            //else if (lBeginCutbackAnchorPoint == 6)
                            dBeginFaceZ = (dZ3 / 2);

                    }
                    else if (lBeginFaceCardinalPt == 11 || lBeginFaceCardinalPt == 10 || lBeginFaceCardinalPt == 14 || lBeginFaceCardinalPt == 15)
                    {
                        if (lBeginCutbackAnchorPoint == AnchorPoint.MiddleLeft)
                            //if (lBeginCutbackAnchorPoint == 4)
                            dBeginFaceZ = -dZ3 / 4;
                        else if (lBeginCutbackAnchorPoint == AnchorPoint.MiddleRight)
                            //else if (lBeginCutbackAnchorPoint == 6)
                            dBeginFaceZ = dZ3 / 4;
                    }
                    else if (lBeginFaceCardinalPt == 3 || lBeginFaceCardinalPt == 6 || lBeginFaceCardinalPt == 9 || lBeginFaceCardinalPt == 13)
                    {
                        if (lBeginCutbackAnchorPoint == AnchorPoint.MiddleLeft)
                            //if (lBeginCutbackAnchorPoint == 4)
                            dBeginFaceZ = -dZ3;
                        else if (lBeginCutbackAnchorPoint == AnchorPoint.MiddleRight)
                            //else if (lBeginCutbackAnchorPoint == 6)
                            dBeginFaceZ = 0;

                    }
                }
                else if (lBeginCutbackAnchorPoint == AnchorPoint.BottomCenter || lBeginCutbackAnchorPoint == AnchorPoint.TopCenter)
                //else if (lBeginCutbackAnchorPoint == 2 || lBeginCutbackAnchorPoint == 8)
                {
                    if (lBeginFaceCardinalPt == 1 || lBeginFaceCardinalPt == 2 || lBeginFaceCardinalPt == 3 || lBeginFaceCardinalPt == 11)
                    {
                        if (lBeginCutbackAnchorPoint == AnchorPoint.BottomCenter)
                            //if (lBeginCutbackAnchorPoint == 2)
                            dBeginFaceZ = 0;
                        else if (lBeginCutbackAnchorPoint == AnchorPoint.TopCenter)
                            //else if (lBeginCutbackAnchorPoint == 8)
                            dBeginFaceZ = dZ5;
                    }
                    else if (lBeginFaceCardinalPt == 4 || lBeginFaceCardinalPt == 5 || lBeginFaceCardinalPt == 6 || lBeginFaceCardinalPt == 15)
                    {
                        if (lBeginCutbackAnchorPoint == AnchorPoint.BottomCenter)
                            //if (lBeginCutbackAnchorPoint == 2)
                            dBeginFaceZ = -dZ5 / 2;
                        else if (lBeginCutbackAnchorPoint == AnchorPoint.TopCenter)
                            //else if (lBeginCutbackAnchorPoint == 8)
                            dBeginFaceZ = dZ5 / 2;
                    }
                    else if (lBeginFaceCardinalPt == 10 || lBeginFaceCardinalPt == 12 || lBeginFaceCardinalPt == 13)
                    {
                        if (lBeginCutbackAnchorPoint == AnchorPoint.BottomCenter)
                            //if (lBeginCutbackAnchorPoint == 2)
                            dBeginFaceZ = -dZ5 / 4;
                        else if (lBeginCutbackAnchorPoint == AnchorPoint.TopCenter)
                            //else if (lBeginCutbackAnchorPoint == 8)
                            dBeginFaceZ = dZ5 / 4;
                    }
                    else if (lBeginFaceCardinalPt == 7 || lBeginFaceCardinalPt == 8 || lBeginFaceCardinalPt == 9 || lBeginFaceCardinalPt == 14)
                    {
                        if (lBeginCutbackAnchorPoint == AnchorPoint.BottomCenter)
                            //if (lBeginCutbackAnchorPoint == 2)
                            dBeginFaceZ = -dZ5;
                        else if (lBeginCutbackAnchorPoint == AnchorPoint.TopCenter)
                            //else if (lBeginCutbackAnchorPoint == 8)
                            dBeginFaceZ = 0;

                    }
                }


                return dBeginFaceZ;

            }
            catch (Exception oEx)
            {
                throw oEx;
            }

        }

        private double BeginFaceZCoOrdinateForSnip(Part oPart, long CardinalPt, double BeginCBAngle, AnchorPoint BeginCBAnchorPt, long FacePortOrient, double BeginOffsetAlongFlange, double BeginOffsetAlongWeb)
        {
            double dBeginFaceZ = 0;
            try
            {

                double dWidth, dFlangeThickness, dWebThickness, dDepth;
                GetSectionData(oPart, out dWidth, out dFlangeThickness, out dWebThickness, out dDepth);

                AnchorPoint lBeginCBAnchorPt; 
                long lBeginFaceCardinalPt;
                double dCBBeginAngle, dBeginOffsetAlongFlange, dBeginOffsetAlongWeb, dBeginFaceZCoord = 0;

                lBeginCBAnchorPt = BeginCBAnchorPt;
                dCBBeginAngle = BeginCBAngle;
                lBeginFaceCardinalPt = CardinalPt;
                dBeginOffsetAlongFlange = BeginOffsetAlongFlange;
                dBeginOffsetAlongWeb = BeginOffsetAlongWeb;

                double dZ3, dZ5;

                dZ3 = dWidth * Math.Tan(-dCBBeginAngle);    //for Anchor Points 4 and 6
                dZ5 = dDepth * Math.Tan(dCBBeginAngle);    //for Anchor Points 2 and 8

                //Begin Face and End Face Port Offsets
                if (FacePortOrient == 1)      //for flange cutback
                {
                    if (lBeginFaceCardinalPt == 1 || lBeginFaceCardinalPt == 4 || lBeginFaceCardinalPt == 7 || lBeginFaceCardinalPt == 12)
                    {
                        if (lBeginCBAnchorPt == AnchorPoint.MiddleLeft)
                            //if (lBeginCBAnchorPt == 4)
                            dBeginFaceZCoord = 0;
                        else if (lBeginCBAnchorPt == AnchorPoint.MiddleRight)
                            //else if (lBeginCBAnchorPt == 6)
                            if (CompareDoubleGreaterthan(dBeginOffsetAlongFlange, LINEAR_TOLERANCE))
                                dBeginFaceZCoord = 0;
                            else
                                dBeginFaceZCoord = dZ3;

                    }
                    else if (lBeginFaceCardinalPt == 2 || lBeginFaceCardinalPt == 5 || lBeginFaceCardinalPt == 8)
                    {
                        if (CompareDoubleGreaterthan(dBeginOffsetAlongFlange, dWidth / 2 + LINEAR_TOLERANCE))
                            dBeginFaceZCoord = 0;
                        else
                            if (lBeginCBAnchorPt == AnchorPoint.MiddleLeft)
                                //if (lBeginCBAnchorPt == 4)
                                dBeginFaceZCoord = -(dWidth / 2 - dBeginOffsetAlongFlange) * Math.Tan(dCBBeginAngle);
                            else if (lBeginCBAnchorPt == AnchorPoint.MiddleRight)
                                //else if (lBeginCBAnchorPt == 6)
                                dBeginFaceZCoord = dWidth / 2 * Math.Tan(-dCBBeginAngle);

                    }
                    else if (lBeginFaceCardinalPt == 11 || lBeginFaceCardinalPt == 10 || lBeginFaceCardinalPt == 14 || lBeginFaceCardinalPt == 15)
                    {
                        if (CompareDoubleGreaterthan(dBeginOffsetAlongFlange, dWidth / 4 + LINEAR_TOLERANCE))
                            dBeginFaceZCoord = 0;
                        else
                            if (lBeginCBAnchorPt == AnchorPoint.MiddleLeft)
                                //if (lBeginCBAnchorPt == 4)
                                dBeginFaceZCoord = -(dWidth / 4 - dBeginOffsetAlongFlange) * Math.Tan(-dCBBeginAngle);
                            else if (lBeginCBAnchorPt == AnchorPoint.MiddleRight)
                                //else if (lBeginCBAnchorPt == 6)
                                dBeginFaceZCoord = dWidth / 4 * Math.Tan(-dCBBeginAngle);

                    }
                    else if (lBeginFaceCardinalPt == 3 || lBeginFaceCardinalPt == 6 || lBeginFaceCardinalPt == 9 || lBeginFaceCardinalPt == 13)
                    {
                        if (lBeginCBAnchorPt == AnchorPoint.MiddleLeft)
                            //if (lBeginCBAnchorPt == 4)
                            dBeginFaceZCoord = (dWidth - dBeginOffsetAlongFlange) * Math.Tan(-dCBBeginAngle);
                        else if (lBeginCBAnchorPt == AnchorPoint.MiddleRight)
                            //else if (lBeginCBAnchorPt == 6)
                            dBeginFaceZCoord = 0;

                    }
                }
                else if (FacePortOrient == 2)      //For web cutback
                {
                    if (lBeginFaceCardinalPt == 1 || lBeginFaceCardinalPt == 2 || lBeginFaceCardinalPt == 3 || lBeginFaceCardinalPt == 11)
                    {
                        if (lBeginCBAnchorPt == AnchorPoint.BottomCenter)
                            //if (lBeginCBAnchorPt == 2)
                            dBeginFaceZCoord = 0;
                        else if (lBeginCBAnchorPt == AnchorPoint.TopCenter)
                            //else if (lBeginCBAnchorPt == 8)
                            if (CompareDoubleGreaterthan(dBeginOffsetAlongWeb, LINEAR_TOLERANCE))
                                dBeginFaceZCoord = 0;
                            else
                                dBeginFaceZCoord = dZ5;

                    }
                    else if (lBeginFaceCardinalPt == 4 || lBeginFaceCardinalPt == 5 || lBeginFaceCardinalPt == 6 || lBeginFaceCardinalPt == 15)
                    {
                        if (CompareDoubleGreaterthan(dBeginOffsetAlongWeb, dDepth / 2 + LINEAR_TOLERANCE))
                            dBeginFaceZCoord = 0;
                        else
                            if (lBeginCBAnchorPt == AnchorPoint.BottomCenter)
                                //if (lBeginCBAnchorPt == 2)
                                dBeginFaceZCoord = -(dDepth / 2 - dBeginOffsetAlongWeb) * Math.Tan(dCBBeginAngle);
                            else if (lBeginCBAnchorPt == AnchorPoint.TopCenter)
                                //else if (lBeginCBAnchorPt == 8)
                                dBeginFaceZCoord = dDepth / 2 * Math.Tan(dCBBeginAngle);
                    }
                    else if (lBeginFaceCardinalPt == 10 || lBeginFaceCardinalPt == 12 || lBeginFaceCardinalPt == 13)
                    {
                        if (CompareDoubleGreaterthan(dBeginOffsetAlongWeb, dDepth / 2 + LINEAR_TOLERANCE))
                            dBeginFaceZCoord = 0;
                        else
                            if (lBeginCBAnchorPt == AnchorPoint.BottomCenter)
                                //if (lBeginCBAnchorPt == 2)
                                dBeginFaceZCoord = -(dDepth / 4 - dBeginOffsetAlongWeb) * Math.Tan(dCBBeginAngle);
                            else if (lBeginCBAnchorPt == AnchorPoint.TopCenter)
                                //else if (lBeginCBAnchorPt == 8)
                                dBeginFaceZCoord = dDepth / 4 * Math.Tan(dCBBeginAngle);

                    }
                    else if (lBeginFaceCardinalPt == 7 || lBeginFaceCardinalPt == 8 || lBeginFaceCardinalPt == 9 || lBeginFaceCardinalPt == 14)
                    {
                        if (lBeginCBAnchorPt == AnchorPoint.BottomCenter)
                            //if (lBeginCBAnchorPt == 2)
                            dBeginFaceZCoord = -(dDepth - dBeginOffsetAlongWeb) * Math.Tan(dCBBeginAngle);
                        else if (lBeginCBAnchorPt == AnchorPoint.TopCenter)
                            //else if (lBeginCBAnchorPt == 8)
                            dBeginFaceZCoord = 0;
                    }

                }

                dBeginFaceZ = dBeginFaceZCoord;
                return dBeginFaceZ;
            }
            catch (Exception oEx)
            {
                throw oEx;
            }

        }

        private void GetSectionData(Part oPart, out double dWidth, out double dFlangeThickness, out double dWebThickness, out double dDepth)
        {
            try
            {
                //Get the Cross Section from the Relationship between the HgrPart and the CrossSection
                RelationCollection oHgrRelation;
                CrossSection oCrossSection;
                CrossSectionServices oCrossSectionServices = new CrossSectionServices();

                oHgrRelation = oPart.GetRelationship("HgrCrossSection", "CrossSection");
                oCrossSection = (CrossSection)oHgrRelation.TargetObjects.First();

                //Get Section type and standard
                string sSectionType = oCrossSection.Name;
                string sSectionStd = (oCrossSection.CrossSectionClass.Name);
                dFlangeThickness = 0;
                dWebThickness = 0;
                //Get Required Properties From Cross Section
                dWidth = (double)((PropertyValueDouble)oCrossSection.GetPropertyValue("IStructCrossSectionDimensions", "Width")).PropValue;
                dDepth = (double)((PropertyValueDouble)oCrossSection.GetPropertyValue("IStructCrossSectionDimensions", "Depth")).PropValue;
                if (oCrossSection.SupportsInterface("IStructFlangedSectionDimensions"))
                {
                    dFlangeThickness = (double)((PropertyValueDouble)oCrossSection.GetPropertyValue("IStructFlangedSectionDimensions", "tf")).PropValue;
                    try
                    {
                        dWebThickness = (double)((PropertyValueDouble)oCrossSection.GetPropertyValue("IStructFlangedSectionDimensions", "tw")).PropValue;
                    }
                    catch
                    {
                        dWebThickness = dFlangeThickness;
                    }
                }

            }
            catch (Exception oEx)
            {
                throw oEx;
            }
        }

        private bool CompareDoubleGreaterthan(double LeftVariable, double RightVariable)
        {
            try
            {
                if (LeftVariable > (RightVariable - LINEAR_TOLERANCE))
                    return true;
                else
                    return false;

            }
            catch (Exception oEx)
            {
                throw oEx;
            }
        }

        private double EndFaceZCoordForCutback(Part oPart, AnchorPoint EndCBAnchorPt, long CardinalPt, double EndCBAngle)
        {
            try
            {

                AnchorPoint lEndCBAnchorPt;
                long lEndFaceCardinalPt;
                double dEndFaceZCoord = 0, dCBEndAngle;

                lEndCBAnchorPt = EndCBAnchorPt;
                dCBEndAngle = EndCBAngle;
                lEndFaceCardinalPt = CardinalPt;

                double dWidth, dFlangeThickness, dWebThickness, dDepth;
                double dZ4, dZ6;

                GetSectionData(oPart, out dWidth, out dFlangeThickness, out dWebThickness, out dDepth);

                dZ4 = dWidth * Math.Tan(-dCBEndAngle);   //for Anchor Points 4 and 6
                dZ6 = dDepth * Math.Tan(dCBEndAngle);     //for Anchor Points 2 and 8

                if (lEndCBAnchorPt == AnchorPoint.MiddleLeft || lEndCBAnchorPt == AnchorPoint.MiddleRight)
                    //if (lEndCBAnchorPt == 4 || lEndCBAnchorPt == 6)
                {
                    if (lEndFaceCardinalPt == 1 || lEndFaceCardinalPt == 4 || lEndFaceCardinalPt == 7 || lEndFaceCardinalPt == 12)
                    {
                        if (lEndCBAnchorPt == AnchorPoint.MiddleLeft)
                            //if (lEndCBAnchorPt == 4)
                            dEndFaceZCoord = 0;
                        else if (lEndCBAnchorPt == AnchorPoint.MiddleRight)
                            //else if (lEndCBAnchorPt == 6)
                            dEndFaceZCoord = -dZ4;
                    }
                    else if (lEndFaceCardinalPt == 2 || lEndFaceCardinalPt == 5 || lEndFaceCardinalPt == 8)
                    {
                        if (lEndCBAnchorPt == AnchorPoint.MiddleLeft)
                            //if (lEndCBAnchorPt == 4)
                            dEndFaceZCoord = -dZ4 / 2;
                        else if (lEndCBAnchorPt == AnchorPoint.MiddleRight)
                            //else if (lEndCBAnchorPt == 6)
                            dEndFaceZCoord = dZ4 / 2;
                    }
                    else if (lEndFaceCardinalPt == 11 || lEndFaceCardinalPt == 10 || lEndFaceCardinalPt == 14 || lEndFaceCardinalPt == 15)
                    {
                        if (lEndCBAnchorPt == AnchorPoint.MiddleLeft)
                            //if (lEndCBAnchorPt == 4)
                            dEndFaceZCoord = -dZ4 / 4;
                        else if (lEndCBAnchorPt == AnchorPoint.MiddleRight)
                            //else if (lEndCBAnchorPt == 6)
                            dEndFaceZCoord = dZ4 / 4;
                    }
                    else if (lEndFaceCardinalPt == 3 || lEndFaceCardinalPt == 6 || lEndFaceCardinalPt == 9 || lEndFaceCardinalPt == 13)
                    {
                        if (lEndCBAnchorPt == AnchorPoint.MiddleLeft)
                            //if (lEndCBAnchorPt == 4)
                            dEndFaceZCoord = -dZ4;
                        else if (lEndCBAnchorPt == AnchorPoint.MiddleRight)
                            //else if (lEndCBAnchorPt == 6)
                            dEndFaceZCoord = 0;
                    }

                }
                else if (lEndCBAnchorPt == AnchorPoint.BottomCenter || lEndCBAnchorPt == AnchorPoint.TopCenter)
                    //else if (lEndCBAnchorPt == 2 || lEndCBAnchorPt == 8)
                {
                    if (lEndFaceCardinalPt == 1 || lEndFaceCardinalPt == 2 || lEndFaceCardinalPt == 3 || lEndFaceCardinalPt == 11)
                    {
                        if (lEndCBAnchorPt == AnchorPoint.BottomCenter)
                            //if (lEndCBAnchorPt == 2)
                            dEndFaceZCoord = 0;
                        else if (lEndCBAnchorPt == AnchorPoint.TopCenter)
                            //else if (lEndCBAnchorPt == 8)
                            dEndFaceZCoord = dZ6;
                    }
                    else if (lEndFaceCardinalPt == 4 || lEndFaceCardinalPt == 5 || lEndFaceCardinalPt == 6 || lEndFaceCardinalPt == 15)
                    {
                        if (lEndCBAnchorPt == AnchorPoint.BottomCenter)
                            //if (lEndCBAnchorPt == 2)
                            dEndFaceZCoord = -dZ6 / 2;
                        else if (lEndCBAnchorPt == AnchorPoint.TopCenter)
                            //else if (lEndCBAnchorPt == 8)
                            dEndFaceZCoord = dZ6 / 2;
                    }
                    else if (lEndFaceCardinalPt == 10 || lEndFaceCardinalPt == 12 || lEndFaceCardinalPt == 13)
                    {
                        if (lEndCBAnchorPt == AnchorPoint.BottomCenter)
                            //if (lEndCBAnchorPt == 2)
                            dEndFaceZCoord = -dZ6 / 4;
                        else if (lEndCBAnchorPt == AnchorPoint.TopCenter)
                            //else if (lEndCBAnchorPt == 8)
                            dEndFaceZCoord = dZ6 / 4;
                    }
                    else if (lEndFaceCardinalPt == 7 || lEndFaceCardinalPt == 8 || lEndFaceCardinalPt == 9 || lEndFaceCardinalPt == 14)
                    {
                        if (lEndCBAnchorPt == AnchorPoint.BottomCenter)
                            //if (lEndCBAnchorPt == 2)
                            dEndFaceZCoord = -dZ6;
                        else if (lEndCBAnchorPt == AnchorPoint.TopCenter)
                            //else if (lEndCBAnchorPt == 8)
                            dEndFaceZCoord = 0;
                    }
                }
                return dEndFaceZCoord;

            }
            catch (Exception oEx)
            {
                throw oEx;
            }
        }

        private double EndFaceZCoordForSnip(Part oPart, long CardinalPt, double EndCBAngle, AnchorPoint EndCBAnchorPt, long FacePortOrient, double EndOffsetAlongFlange, double EndOffsetAlongWeb)
        {
            try
            {

                AnchorPoint lEndCBAnchorPt;
                long lEndFaceCardinalPt;
                double dEndFaceZCoord = 0, dEndOffsetAlongWeb, dEndOffsetAlongFlange, dCBEndAngle;

                lEndCBAnchorPt = EndCBAnchorPt;
                dCBEndAngle = EndCBAngle;
                lEndFaceCardinalPt = CardinalPt;
                dEndOffsetAlongWeb = EndOffsetAlongWeb;
                dEndOffsetAlongFlange = EndOffsetAlongFlange;

                double dWidth, dFlangeThickness, dWebThickness, dDepth;
                double dZ4, dZ6;

                GetSectionData(oPart, out dWidth, out dFlangeThickness, out dWebThickness, out dDepth);
                dZ4 = dWidth * Math.Tan(dCBEndAngle);     //for Anchor Points 4 and 6
                dZ6 = dDepth * Math.Tan(dCBEndAngle);     //for Anchor Points 2 and 8

                if (FacePortOrient == 1)      //for flange cutback
                {
                    if (lEndFaceCardinalPt == 1 || lEndFaceCardinalPt == 4 || lEndFaceCardinalPt == 7 || lEndFaceCardinalPt == 12)
                    {
                        if (lEndCBAnchorPt == AnchorPoint.MiddleLeft)
                            //if (lEndCBAnchorPt == 4)
                            dEndFaceZCoord = 0;
                        else if (lEndCBAnchorPt == AnchorPoint.MiddleRight)
                            //else if (lEndCBAnchorPt == 6)
                            if (CompareDoubleGreaterthan(dEndOffsetAlongFlange, LINEAR_TOLERANCE))
                                dEndFaceZCoord = 0;
                            else
                                dEndFaceZCoord = dZ4;
                    }
                    else if (lEndFaceCardinalPt == 2 || lEndFaceCardinalPt == 5 || lEndFaceCardinalPt == 8)
                    {
                        if (CompareDoubleGreaterthan(dEndOffsetAlongFlange, dWidth / 2 + LINEAR_TOLERANCE))
                            dEndFaceZCoord = 0;
                        else
                        {
                            if (lEndCBAnchorPt == AnchorPoint.MiddleLeft)
                                //if (lEndCBAnchorPt == 4)
                                dEndFaceZCoord = -(dWidth / 2 - dEndOffsetAlongFlange) * Math.Tan(dCBEndAngle);
                            else if (lEndCBAnchorPt == AnchorPoint.MiddleRight)
                                //else if (lEndCBAnchorPt == 6)
                                dEndFaceZCoord = dWidth / 2 * Math.Tan(dCBEndAngle);
                        }
                    }
                    else if (lEndFaceCardinalPt == 11 || lEndFaceCardinalPt == 10 || lEndFaceCardinalPt == 14 || lEndFaceCardinalPt == 15)
                    {
                        if (CompareDoubleGreaterthan(dEndOffsetAlongFlange, dWidth / 4 + LINEAR_TOLERANCE))
                            dEndFaceZCoord = 0;
                        else
                        {
                            if (lEndCBAnchorPt == AnchorPoint.MiddleLeft)
                                //if (lEndCBAnchorPt == 4)
                                dEndFaceZCoord = -(dWidth / 4 - dEndOffsetAlongFlange) * Math.Tan(dCBEndAngle);
                            else if (lEndCBAnchorPt == AnchorPoint.MiddleRight)
                                //else if (lEndCBAnchorPt == 6)
                                dEndFaceZCoord = dWidth / 4 * Math.Tan(dCBEndAngle);
                        }
                    }
                    else if (lEndFaceCardinalPt == 3 || lEndFaceCardinalPt == 6 || lEndFaceCardinalPt == 9 || lEndFaceCardinalPt == 13)
                    {
                        if (lEndCBAnchorPt == AnchorPoint.MiddleLeft)
                            //if (lEndCBAnchorPt == 4)
                            dEndFaceZCoord = -(dWidth - dEndOffsetAlongFlange) * Math.Tan(dCBEndAngle);
                        else if (lEndCBAnchorPt == AnchorPoint.MiddleRight)
                            //else if (lEndCBAnchorPt == 6)
                            dEndFaceZCoord = 0;
                    }
                }
                else if (FacePortOrient == 2)      //'for web cutback
                {
                    if (lEndFaceCardinalPt == 1 || lEndFaceCardinalPt == 2 || lEndFaceCardinalPt == 3 || lEndFaceCardinalPt == 11)
                    {
                        if (lEndCBAnchorPt == AnchorPoint.BottomCenter)
                            //if (lEndCBAnchorPt == 2)
                            dEndFaceZCoord = 0;
                        else if (lEndCBAnchorPt == AnchorPoint.TopCenter)
                        //else if (lEndCBAnchorPt == 8)
                        {
                            if (CompareDoubleGreaterthan(dEndOffsetAlongWeb, LINEAR_TOLERANCE))
                                dEndFaceZCoord = 0;
                            else
                                dEndFaceZCoord = dZ6;
                        }
                    }
                    else if (lEndFaceCardinalPt == 4 || lEndFaceCardinalPt == 5 || lEndFaceCardinalPt == 6 || lEndFaceCardinalPt == 15)
                    {
                        if (CompareDoubleGreaterthan(dEndOffsetAlongWeb, dDepth / 2 + LINEAR_TOLERANCE))
                            dEndFaceZCoord = 0;
                        else
                        {
                            if (lEndCBAnchorPt == AnchorPoint.BottomCenter)
                                //if (lEndCBAnchorPt == 2)
                                dEndFaceZCoord = -(dDepth / 2 - dEndOffsetAlongWeb) * Math.Tan(dCBEndAngle);
                            else if (lEndCBAnchorPt == AnchorPoint.TopCenter)
                                //else if (lEndCBAnchorPt == 8)
                                dEndFaceZCoord = dDepth / 2 * Math.Tan(dCBEndAngle);
                        }
                    }
                    else if (lEndFaceCardinalPt == 10 || lEndFaceCardinalPt == 12 || lEndFaceCardinalPt == 13)
                    {
                        if (CompareDoubleGreaterthan(dEndOffsetAlongWeb, dDepth / 2 + LINEAR_TOLERANCE))
                            dEndFaceZCoord = 0;
                        else
                        {
                            if (lEndCBAnchorPt == AnchorPoint.BottomCenter)
                                //if (lEndCBAnchorPt == 2)
                                dEndFaceZCoord = -(dDepth / 4 - dEndOffsetAlongWeb) * Math.Tan(dCBEndAngle);
                            else if (lEndCBAnchorPt == AnchorPoint.TopCenter)
                                //else if (lEndCBAnchorPt == 8)
                                dEndFaceZCoord = dDepth / 4 * Math.Tan(dCBEndAngle);
                        }
                    }
                    else if (lEndFaceCardinalPt == 7 || lEndFaceCardinalPt == 8 || lEndFaceCardinalPt == 9 || lEndFaceCardinalPt == 14)
                    {
                        if (lEndCBAnchorPt == AnchorPoint.BottomCenter)
                            //if (lEndCBAnchorPt == 2)
                            dEndFaceZCoord = -(dDepth - dEndOffsetAlongWeb) * Math.Tan(dCBEndAngle);
                        else if (lEndCBAnchorPt == AnchorPoint.TopCenter)
                            //else if (lEndCBAnchorPt == 8)
                            dEndFaceZCoord = 0;
                    }
                }

                return dEndFaceZCoord;
            }
            catch (Exception oEx)
            {
                throw oEx;
            }
        }
        
        #region ICustomHgrWeightCG Members

        void ICustomHgrWeightCG.WeightCG(SupportComponent supportComponent, ref double weight, ref double cogX, ref double cogY, ref double cogZ)
        {
            try
            {
                //Get the Cross Section from the Relationship between the HgrPart and the CrossSection
                RelationCollection oHgrRelation;
                CrossSection oCrossSection;
                CrossSectionServices oCrossSectionServices = new CrossSectionServices();
                Part oSectionPart = (Part)supportComponent.GetRelationship("madeFrom", "part").TargetObjects[0];

                oHgrRelation = oSectionPart.GetRelationship("HgrCrossSection", "CrossSection");
                oCrossSection = (CrossSection)oHgrRelation.TargetObjects.First();
                
                //Get Section type and standard
                string sSectionType = oCrossSection.Name;
                string sSectionStd = (oCrossSection.CrossSectionClass.Name);
                
                
                double dBeginOverLen = (double)((PropertyValueDouble)GetPropValue(supportComponent, "BeginOverLength")).PropValue;
                double dEndOverLen = (double)((PropertyValueDouble)GetPropValue(supportComponent, "EndOverLength")).PropValue;
                double dBeginCutbackAng = (double)((PropertyValueDouble)GetPropValue(supportComponent, "CutbackBeginAngle")).PropValue;
                double dEndCutbackAng = (double)((PropertyValueDouble)GetPropValue(supportComponent, "CutbackEndAngle")).PropValue;
                
                double dBeginAlongFlangeAngle = (double)((PropertyValueDouble)GetPropValue(supportComponent,"CutbackBeginAngle1")).PropValue;
                double dBeginAlongWebAngle = (double)((PropertyValueDouble)GetPropValue(supportComponent,"CutbackBeginAngle2")).PropValue;
                double dEndAlongFlangeAngle = (double)((PropertyValueDouble)GetPropValue(supportComponent,"CutbackEndAngle1")).PropValue;
                double dEndAlongWebAngle = (double)((PropertyValueDouble)GetPropValue(supportComponent, "CutbackEndAngle2")).PropValue;
                double dBeginAlongFlangeOffset = (double)((PropertyValueDouble)GetPropValue(supportComponent,"BeginOffsetAlongFlange")).PropValue;
                double dBeginAlongWebOffset = (double)((PropertyValueDouble)GetPropValue(supportComponent, "BeginOffsetAlongWeb")).PropValue;
                double dEndAlongFlangeOffset = (double)((PropertyValueDouble)GetPropValue(supportComponent,"EndOffsetAlongFlange")).PropValue;
                double dEndAlongWebOffset = (double)((PropertyValueDouble)GetPropValue(supportComponent,"EndOffsetAlongWeb")).PropValue;
                double lCutLength;
                string sMaterialGrade = (string)((PropertyValueString)GetPropValue(supportComponent, "MaterialGrade")).PropValue;
                string sMaterialType = (string)((PropertyValueString)GetPropValue(supportComponent, "MaterialType")).PropValue;
                
                double dLength;
                
                try
                {
                    dLength = (double)((PropertyValueDouble)GetPropValue(supportComponent, "Length")).PropValue;
                }
                catch 
                {
                    dLength = 0;
                }

                try
                {
                    lCutLength = (double)((PropertyValueDouble)GetPropValue(supportComponent, "VarLength")).PropValue;
                }
                catch 
                {
                    lCutLength = 0;
                }

                long lBeginAncPoint = (long)((PropertyValueCodelist)GetPropValue(supportComponent,"BeginCutbackAnchorPoint")).PropValue;
                long lEndAncPoint = (long)((PropertyValueCodelist)GetPropValue(supportComponent, "EndCutbackAnchorPoint")).PropValue;
                long lBeginAlongFlangeAnchrPnt = (long)((PropertyValueCodelist)GetPropValue(supportComponent, "BeginCutbackAnchorPoint1")).PropValue;
                long lBeginAlongWebAnchrPnt = (long)((PropertyValueCodelist)GetPropValue(supportComponent, "BeginCutbackAnchorPoint2")).PropValue;
                long lEndAlongFlangeAnchrPnt = (long)((PropertyValueCodelist)GetPropValue(supportComponent, "EndCutbackAnchorPoint1")).PropValue;
                long lEndAlongWebAnchrPnt = (long)((PropertyValueCodelist)GetPropValue(supportComponent, "EndCutbackAnchorPoint2")).PropValue;
                long bIsCutBack = (long)((PropertyValueCodelist)GetPropValue(supportComponent, "HgrBeamType")).PropValue;
                double dMaxLength=0;

                double dDensity=0.25;// = GetDensity();   //Ramya
                double dWidth,dblWebThk, dblFlangeThk, dblOuterDia, dInnerDia, dblArea, dblNominalThk, totLength=0, dCutLen=0;
                Position centerOfGravity = new Position();
                //cogX = CogX;
                //cogY = CogY;
                //cogZ = CogZ;
                //weight = Weight;

                GetSectionData(oSectionPart, out dWidth, out dblFlangeThk, out dblWebThk, out dblOuterDia);
                
                if (dLength <= 0)
                {
                    dLength = 0.5;
                    supportComponent.SetPropertyValue(dLength, "IJUAHgrOccLength", "Length");
                }
                else
                {
                    try
                    {
                        dMaxLength = (double)((PropertyValueDouble)GetPropValue(oSectionPart, "MaxLength")).PropValue;
                        if ((dLength + dBeginOverLen + dEndOverLen) > dMaxLength)
                        {
                            dLength = dMaxLength - dBeginOverLen - dEndOverLen;
                            supportComponent.SetPropertyValue(dLength, "IJUAHgrOccLength", "Length");
                        }
                    }
                    catch 
                    {
                        dMaxLength = 0;
                    }

                    if (bIsCutBack == CUTBACK_STEEL)
                    {
                        if (sSectionType == "HSSC" || sSectionType == "PIPE" || sSectionType == "CS")
                        {
                            
                            if (sSectionType == "CS")
                            {
                                dInnerDia = dblOuterDia;
                                dblArea = (Math.PI * 0.25) * (dblOuterDia * dblOuterDia);
                            }
                            else
                            {
                                dblNominalThk = (double)((PropertyValueDouble)oCrossSection.GetPropertyValue("IJUAHSS", "tnom")).PropValue;
                                dInnerDia = dblOuterDia - 2 * dblNominalThk;
                                dblArea = (Math.PI * 0.25) * (dblOuterDia * dblOuterDia - dInnerDia * dInnerDia);
                            }

                            oCrossSectionServices.GetCardinalPointOffset(oCrossSection, 10, out cogX, out cogY);
                            
                            totLength = (dBeginOverLen + dLength + dEndOverLen);
                            cogZ = totLength / 2;

                            weight = dblArea * totLength * dDensity;
                            dCutLen = dBeginOverLen + dLength + dEndOverLen;

                        }
                        else
                        {
                            oCutbackSteelInput = new CutbackSteelInputs();
                            oCutbackSteelInput.BeginAnchorPoint = (AnchorPoint)lBeginAncPoint;
                            oCutbackSteelInput.BeginOverLength = dBeginOverLen;
                            oCutbackSteelInput.CutbackBeginAngle = dBeginCutbackAng;
                            oCutbackSteelInput.CutbackEndAngle = dEndCutbackAng;
                            oCutbackSteelInput.Density = 0.25; // default value
                            oCutbackSteelInput.EndAnchorPoint = (AnchorPoint)lEndAncPoint;
                            oCutbackSteelInput.EndOverLength = dEndOverLen;
                            oCutbackSteelInput.Length = dLength;
                            oCutbackSteelInput.Part = oSectionPart;

                            CalculateWeightCG(oCutbackSteelInput, ref weight, ref totLength, ref centerOfGravity);
                            //oGenericHelper.ConnectionComponentCutbackWeightCG(oSectionPart, dBeginOverLen, dLength, dEndOverLen, (int)lBeginAncPoint, (int)lEndAncPoint, dBeginCutbackAng, dEndCutbackAng, out lCutLength, out dDensity, out weight, out cogX, out cogY, out cogZ);
                            cogX = centerOfGravity.X;
                            cogY = centerOfGravity.Y;
                            cogZ = centerOfGravity.Z;

                            if (HgrCompareDoubleService.cmpdbl(dLength , dMaxLength)== true)
                            {
              

                                dCutLen = totLength - dBeginOverLen - dEndOverLen;
                            }
                            else
                                dCutLen = totLength;
                        }

                    }
                    else
                    {
                        if (sSectionType == "L")
                        {
                            if (dBeginAlongFlangeAngle < 0 || dBeginAlongWebAngle > 0 || dEndAlongFlangeAngle > 0 || dEndAlongWebAngle < 0)
                            {
                                if (dBeginAlongFlangeAngle < 0)
                                {
                                    dBeginAlongFlangeAngle = Math.Abs(dBeginAlongFlangeAngle);
                                    supportComponent.SetPropertyValue(dBeginAlongFlangeAngle, "IJOAhsSnipedSteel", "CutbackBeginAngle1");
                                }
                                if (dBeginAlongWebAngle > 0)
                                {
                                    dBeginAlongWebAngle = -dBeginAlongWebAngle;
                                    supportComponent.SetPropertyValue(dBeginAlongWebAngle, "IJOAhsSnipedSteel", "CutbackBeginAngle2");
                                }
                                if (dEndAlongFlangeAngle > 0)
                                {
                                    dEndAlongFlangeAngle = -dEndAlongFlangeAngle;
                                    supportComponent.SetPropertyValue(dEndAlongFlangeAngle, "IJOAhsSnipedSteel", "CutbackEndAngle1");
                                }
                                if (dEndAlongWebAngle < 0)
                                {
                                    dEndAlongWebAngle = Math.Abs(dEndAlongWebAngle);
                                    supportComponent.SetPropertyValue(dEndAlongWebAngle, "IJOAhsSnipedSteel", "CutbackEndAngle2");
                                }
                            }
                            
                            dCutLen = dBeginOverLen + dLength + dEndOverLen;

                            if (((dBeginAlongWebOffset > dblFlangeThk && dEndAlongWebOffset > dblFlangeThk) && (dBeginAlongFlangeOffset > dblWebThk && dEndAlongFlangeOffset > dblWebThk)) || (((HgrCompareDoubleService.cmpdbl(dBeginAlongWebOffset, 0) == false || HgrCompareDoubleService.cmpdbl(dEndAlongWebOffset, 0) == false || HgrCompareDoubleService.cmpdbl(dBeginAlongFlangeOffset, 0) == false || HgrCompareDoubleService.cmpdbl(dEndAlongFlangeOffset, 0) == false) && ((HgrCompareDoubleService.cmpdbl(dBeginAlongFlangeAngle, 0) == true) && (HgrCompareDoubleService.cmpdbl(dBeginAlongWebAngle, 0) == true) && (HgrCompareDoubleService.cmpdbl(dEndAlongFlangeAngle, 0) == true) && (HgrCompareDoubleService.cmpdbl(dEndAlongWebAngle, 0) == true)))))
                            {
                                //oGenericHelper.ConnectionComponentSnipWeightCG(oSectionPart, dBeginOverLen, dLength, dEndOverLen, (int)lBeginAlongFlangeAnchrPnt, (int)lEndAlongFlangeAnchrPnt, (int)lBeginAlongWebAnchrPnt, (int)lEndAlongWebAnchrPnt, dBeginAlongFlangeOffset, dBeginAlongWebOffset, dEndAlongFlangeOffset, dEndAlongWebOffset, dBeginAlongFlangeAngle, dBeginAlongWebAngle, dEndAlongFlangeAngle, dEndAlongWebAngle, dDensity, out weight, out cogX, out cogY, out cogZ);
                            }
                            else
                            {
                                CalculateWeightAndCGForSnip(oSectionPart, out weight, out cogX, out cogY, out cogZ);
                            }
                        }
                    }
          
                    supportComponent.SetPropertyValue(dCutLen, "IJUAhsVarLength", "VarLength");
                  }
            }
            catch (Exception oEx)
            {
                Exception oNewEx = new Exception("Error in Weight CG of RichHgrBeam", oEx);
                throw oNewEx;
            }

        }

        void CalculateWeightAndCGForSnip(Part supportComponent, out double Weight,out double CogX, out double CogY,out double CogZ)
        {
            try
            {
            double WebLen, WebThk, dLENGTH, FlngLen, FlngThk;
            double BeginOffsetFlng, EndOffsetFlng, BeginOffsetWeb, EndOffsetWeb;
            double BeginOffsetFlng_OtherLen, EndOffsetFlng_OtherLen, BeginOffsetWeb_OtherLen, EndOffsetWeb_OtherLen;
            double dBeginAngleAlongFlng, dBeginAngleAlongWeb, dEndAngleAlongFlng, dEndAngleAlongWeb;
            long lFacePortOrient;
            double dblBeginLength,dblEndLength, lDensity,TotalLen;
            string sMaterialGrade, sMaterialType;
                        
            dBeginAngleAlongFlng = (double)((PropertyValueDouble)GetPropValue(supportComponent, "CutbackBeginAngle1")).PropValue;
            dBeginAngleAlongWeb = (double)((PropertyValueDouble)GetPropValue(supportComponent, "CutbackBeginAngle2")).PropValue;
            dEndAngleAlongFlng = (double)((PropertyValueDouble)GetPropValue(supportComponent, "CutbackEndAngle1")).PropValue;
            dEndAngleAlongWeb = (double)((PropertyValueDouble)GetPropValue(supportComponent, "CutbackEndAngle2")).PropValue;
            
            BeginOffsetFlng = (double)((PropertyValueDouble)GetPropValue(supportComponent, "BeginOffsetAlongFlange")).PropValue;
            EndOffsetFlng = (double)((PropertyValueDouble)GetPropValue(supportComponent,"BeginOffsetAlongWeb")).PropValue;
            BeginOffsetWeb = (double)((PropertyValueDouble)GetPropValue(supportComponent, "EndOffsetAlongFlange")).PropValue;
            EndOffsetWeb = (double)((PropertyValueDouble)GetPropValue(supportComponent, "EndOffsetAlongWeb")).PropValue;
            dLENGTH = (double)((PropertyValueDouble)GetPropValue(supportComponent, "Length")).PropValue;
            lFacePortOrient = (long)((PropertyValueDouble)GetPropValue(supportComponent, "FacePortOrient")).PropValue;
            
            dblBeginLength = (double)((PropertyValueDouble)GetPropValue(supportComponent, "BeginOverLength")).PropValue;
            dblEndLength = (double)((PropertyValueDouble)GetPropValue(supportComponent, "EndOverLength")).PropValue;
            sMaterialGrade = (string)((PropertyValueString)GetPropValue(supportComponent, "MaterialGrade")).PropValue ;
            sMaterialType = (string)((PropertyValueString)GetPropValue(supportComponent, "MaterialType")).PropValue ;
            
            lDensity = 0.25 ;//GetDensity(sMaterialType, sMaterialGrade)
            
            TotalLen = dblBeginLength + dLENGTH + dblEndLength;
          
            double cx=0, cy=0, cz=0;

            //Get the Cross Section from the Relationship between the HgrPart and the CrossSection
            RelationCollection oHgrRelation;
            CrossSection oCrossSection;
            CrossSectionServices oCrossSectionServices = new CrossSectionServices();

            oHgrRelation = supportComponent.GetRelationship("HgrCrossSection", "CrossSection");
            oCrossSection = (CrossSection)oHgrRelation.TargetObjects.First();

            oCrossSectionServices.GetCardinalPointOffset(oCrossSection, 1, out cx, out cy);
            
            cz = dLENGTH / 2;
            GetSectionData(supportComponent,out FlngLen,out FlngThk,out WebThk,out WebLen);
            
            BeginOffsetFlng_OtherLen = OtherLength(dBeginAngleAlongFlng, FlngLen, BeginOffsetFlng);
            EndOffsetFlng_OtherLen = OtherLength(dEndAngleAlongFlng, FlngLen, EndOffsetFlng);
            BeginOffsetWeb_OtherLen = OtherLength(dBeginAngleAlongWeb, WebLen, BeginOffsetWeb);
            EndOffsetWeb_OtherLen = OtherLength(dEndAngleAlongWeb, WebLen, EndOffsetWeb);
               /*     
                ''
                ''    Formula
                ''
                ''    CgX = [(WL * WT * L * WT/2)] + [(FL-WT) * FT * L * (WT + (FL-WT)/2))] _
                ''           - 1 / 2 * (FL - BOFL) * S2 * FT * (FL - 1 / 3 * (FL - BOFL)) _
                ''           - 1 / 2 * (WL - BOWL) * S3 * WT * (WT / 2) _
                ''           - 1 / 2 * (FL - EOFL) * S4 * FT * (FL - 1 / 3 * (FL - EOFL)) _
                ''           - 1 / 2 * (WL - EOFL) * S5 * WT * (WT / 2)   / Total Volume
                ''
                ''
                ''    CgY = [(WL * WT * L * WL/2)] + [(FL-WT) * FT * L * (FT)/2] _
                ''           - 1 / 2 * (FL - BOFL) * S2 * FT * (FT / 2) _
                ''           - 1 / 2 * (WL - BOWL) * S3 * WT * (WL / 2) _
                ''           - 1 / 2 * (FL - EOFL) * S4 * FT * (FT / 2) _
                ''           - 1 / 2 * (WL - EOFL) * S5 * WT * (WL / 2)  / Total Volume
                ''
                ''    CgZ = [(WL * WT * L * L/2)] + [(FL-WT) * FT * L * L/2] _
                ''           - 1 / 2 * (FL - BOFL) * S2 * FT * (FL - 1 / 3 * (FL - BOFL)) _
                ''           - 1 / 2 * (WL - BOWL) * S3 * WT * (WT / 2) _
                ''           - 1 / 2 * (FL - EOFL) * S4 * FT * (FL - 1 / 3 * (FL - EOFL)) _
                ''           - 1 / 2 * (WL - EOFL) * S5 * WT * (WT / 2)  / Total Volume
                ''
                ''  Density = UniitWeight / Area
                ''
                ''  TotalWeight = Volume of snip * Density
                ''
                ''
            */
            double WebVolume, FlngVolume;
            double BeginTriangleAlongFlangeVolume, EndTriangleAlongFlangeVolume, BeginTriangleAlongWebVolume, EndTriangleAlongWebVolume;
            
            WebVolume = WebLen * WebThk * TotalLen;
            FlngVolume = (FlngLen - WebThk) * FlngThk * TotalLen;
            
            BeginTriangleAlongFlangeVolume = 1.0 / 2.0 * (FlngLen - BeginOffsetFlng) * BeginOffsetFlng_OtherLen * FlngThk;
            BeginTriangleAlongWebVolume = 1.0 / 2.0 * (WebLen - BeginOffsetWeb) * BeginOffsetWeb_OtherLen * WebThk;
            EndTriangleAlongFlangeVolume = 1.0 / 2.0 * (FlngLen - EndOffsetFlng) * EndOffsetFlng_OtherLen * FlngThk;
            EndTriangleAlongWebVolume = 1.0 / 2.0 * (WebLen - EndOffsetWeb) * EndOffsetWeb_OtherLen * WebThk;
                
            CogX = ((WebVolume * WebThk / 2) + (FlngVolume * (WebThk + (FlngLen - WebThk) / 2)) - BeginTriangleAlongFlangeVolume * (FlngLen - 1.0 / 3.0 * (FlngLen - BeginOffsetFlng)) - EndTriangleAlongFlangeVolume * (FlngLen - 1.0 / 3.0 * (FlngLen - EndOffsetFlng)) - BeginTriangleAlongWebVolume * (WebThk / 2) - EndTriangleAlongWebVolume * WebThk / 2) / (WebVolume + FlngVolume);
            // Transform the values from the origin
            CogX = CogX + cx;
            
            CogY = ((WebVolume * WebLen / 2) + (FlngVolume * FlngThk / 2)- BeginTriangleAlongFlangeVolume * FlngThk / 2 - EndTriangleAlongFlangeVolume * FlngThk / 2- BeginTriangleAlongWebVolume * WebLen / 2 - EndTriangleAlongWebVolume * WebLen / 2) / ((WebVolume) + (FlngVolume));
            // Transform the values from the origin
            CogY = CogY + cy;
            
            CogZ = ((WebVolume * TotalLen / 2) + (FlngVolume * TotalLen / 2) - BeginTriangleAlongFlangeVolume * (FlngLen - 1.0 / 3.0 * (FlngLen - BeginOffsetFlng)) - EndTriangleAlongFlangeVolume * (FlngLen - 1.0 / 3.0 * (FlngLen - EndOffsetFlng)) - BeginTriangleAlongWebVolume * (WebLen - 1.0 / 3.0 * (WebLen - BeginOffsetWeb)) - EndTriangleAlongWebVolume * (WebLen - 1.0 / 3.0 * (WebLen - EndOffsetWeb))) / (WebVolume + FlngVolume);
               
            // Total Volume of the section
            Weight = (WebVolume + FlngVolume - BeginTriangleAlongFlangeVolume - EndTriangleAlongFlangeVolume - BeginTriangleAlongWebVolume - EndTriangleAlongWebVolume) * lDensity;
        }
        catch(Exception oEx)
        {
            throw oEx;
        }
    }
        

    double OtherLength(double dAngle,double dLENGTH ,double OffsetLength)
    {
        try
        {
            return (Math.Abs(Math.Tan(dAngle) * (dLENGTH - OffsetLength)));
        }
        catch (Exception oEx)
        {
            throw oEx;
        }
    }

    /*
    double GetDensity(string sMaterialType, string sMaterialGrade)
    {
        try
        {

                
        }
        catch (Exception oEx)
        {
 
        }
    }
         
     */
    #endregion

    PropertyValue GetPropValue(BusinessObject oSupportComponent, string sPropName)
        {
            PropertyValue oPropValue = null;
            try
            {
                //MessageBox.Show("In GetPropValue");
                ReadOnlyCollection<PropertyValue> oAllProperties = oSupportComponent.GetAllProperties();

                for (int i = 0; i < oAllProperties.Count(); i++)
                {
                    if (oAllProperties[i].PropertyInfo.Name == sPropName)
                    {
                        oPropValue =oAllProperties[i];
                        break;
                    }
                }
               //  MessageBox.Show(sPropName);
                if (oPropValue == null)
                {
            
                    throw new Exception("Invalid Property Name: " + sPropName);
                }
                else
                    return oPropValue;
                
            }
            catch (Exception oEx)
            {
                throw oEx;
            }
        }

        

        #region ICustomHgrBOMDescription Members

        public string BOMDescription(BusinessObject oSupportOrComponent)
        {
            string sBOMString="";
            try
            {
                //Get the Cross Section from the Relationship between the HgrPart and the CrossSection
                RelationCollection oHgrRelation;
                CrossSection oCrossSection;
                CrossSectionServices oCrossSectionServices = new CrossSectionServices();

                Part oSectionPart = (Part)oSupportOrComponent.GetRelationship("madeFrom", "part").TargetObjects[0];

                oHgrRelation = oSectionPart.GetRelationship("HgrCrossSection", "CrossSection");
                oCrossSection = (CrossSection)oHgrRelation.TargetObjects.First();

                //Get Section type and standard
                string sSectionType = oCrossSection.Name;
                string sSectionStd = (oCrossSection.CrossSectionClass.Name);

                double dblLength = (double)((PropertyValueDouble)GetPropValue(oSupportOrComponent, "VarLength")).PropValue;
                double dblCutLen = dblLength;
                string sLength="";

                Ingr.SP3D.Support.Middle.Support oSupport = (Ingr.SP3D.Support.Middle.Support)oSupportOrComponent.GetRelationship("SupportHasComponents", "Support").TargetObjects[0];
                GenericHelper oGenericHelper = new GenericHelper(oSupport);
                double dUnitValue, dPrecision=0;
                oGenericHelper.GetDataByRule("HgrStructuralBOMUnits", oSupport, out dUnitValue);

                if ((UnitName.DISTANCE_METER == (UnitName)dUnitValue) || ((UnitName)dUnitValue == UnitName.DISTANCE_MILLIMETER))
                {
                    oGenericHelper.GetDataByRule("HgrStructuralBOMDecimals", oSupport, out dPrecision);

                }
                if (UnitName.DISTANCE_INCH == (UnitName)dUnitValue)
                {
                    if (dPrecision > 0)
                       sLength = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, dblCutLen, UnitName.DISTANCE_INCH);
                    else
                        sLength = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, dblCutLen, UnitName.DISTANCE_FOOT, UnitName.DISTANCE_INCH, UnitName.UNIT_NOT_SET);
                }
                else
                    sLength = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, dblCutLen, (UnitName)dUnitValue);
             
                if (sSectionType == "HSSC" || sSectionType == "PIPE" || sSectionType == "CS")
                {
                    sBOMString = oSectionPart.PartDescription + ", Length: " + sLength;
                }
                else
                {
                    sBOMString = oSectionPart.PartDescription + ", Cut Length: " + sLength;
                }
                return sBOMString;
                //If m_InputConfigHlpr Is Nothing Then Exit Sub

            }
            catch (Exception oEx)
            {
                Exception oNewEx = new Exception("Error in BOM of RichHgrBeam", oEx);
                throw oNewEx;
            }
        }

        #endregion
    }
}


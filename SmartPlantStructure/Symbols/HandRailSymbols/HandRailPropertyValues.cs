//'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//
//Copyright 1992 - 2014 Intergraph Corporation, Intergraph Process, Power & Marine. All rights reserved.
//
//File
//  HandRailTypeA.vb
//
//Abstract
//	This is .NET HandRailTypeA symbol. This class subclasses from HandRailSymbolDefinition.
//

//History:
//      Feb 16, 2015    3XCalibur           DI-CP-267808  Implement content changes for support of drop of Handrail 
//
//'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Content.Structure;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.Structure.Middle.Services;
using System.Collections.ObjectModel;

namespace Ingr.SP3D.Content.Structure
{
    /// <summary>
    /// The object of this class is used to hold all the handrail properties.
    /// </summary>
    internal class HandRailPropertyValues
    {
        internal Sketch3D sketchPath;
        internal double height;
        internal double toprailSectionAngle;
        internal double topOfMidrailDimension;
        
        internal double midrailSpacing;
        internal double topOfToePlateDimension;
        internal double beginExtensionLength;
        internal double endExtensionLength;
        internal double segmentMaximumSpacing;
        internal double slopedSegmentMaximumSpacing;
        internal double postSectionAngle;
        internal double midrailSectionAngle;
        internal double toePlateSectionAngle;

        internal int horizontalOffsetType;
        internal int beginTreatmentType;
        internal int endTreatmentType;
        internal int numberOfMidrails;
        internal HandrailPostOrientation orientationValue;
        internal int toprailSectionCP;
        internal int midrailSectionCP;
        internal int postSectionCP;
        internal int toePlateSectionCP;

        internal string midrailSectionName;
        internal string midrailSectionReferenceStandard;
        internal string toePlateSectionName;
        internal string toePlateSectionReferenceStandard;
        internal string toprailSectionName;
        internal string toprailSectionReferenceStandard;
        internal string postSectionName;
        internal string postSectionReferenceStandard;

        internal bool isWithToePlate;
        internal bool isPostAtEveryTurn;
        internal int postConnectionType;
        internal SP3DConnection connection;

        internal Collection<ICurve> sketchPathCurves;
        internal ComplexString3d handrailOffsetPath;
        internal CrossSection topRailCrossSection;
        internal CrossSection midRailCrossSection;
        internal CrossSection toePlateCrossSection;
        internal CrossSection postCrossSection;
        internal SweepOptions sweepOptions;
        internal double horizontalOffsetDimension;
        internal double circularTreatmentOffset;
        internal double lastMidrailHeight;
        internal double toprailSectionHeight;
        internal double toprailSectionWidth;
        internal double toprailSectionMaximumHeight;
        internal string material;
        internal string grade;
        //Default constructor
        internal HandRailPropertyValues()
        {

        }
    }
}

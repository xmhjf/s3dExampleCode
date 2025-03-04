//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   30.12.2014     PVK      TR-CP-264951	Resolve P3 coverity issues found in November 2014 report
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
using System;
using System.Linq;
using Ingr.SP3D.Structure.Middle.Services;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;

namespace Ingr.SP3D.Content.Support.Symbols
{
    class CommonSmartCutBackSteel
    {
        #region "Calculation Of Length For BOM"

        public double CalculateLengthForBOM(Part part, double L, double beginOverLength, double endOverLength, CodelistItem beginAnchorPoint, CodelistItem endAnchorPoint, double cutbackBeginAngle, double cutbackEndAngle)
        {
            double actLength = 0;
            Double Z1 = 0;
            Double Z2 = 0;
            Double Z3 = 0;
            Double Z4 = 0;
            Double Z5 = 0;
            Double Z6 = 0;
            Double Z7 = 0;
            Double Z8 = 0;

            RelationCollection hgrRelation;
            CrossSection crossSection;
            CrossSectionServices crossSectionServices = new CrossSectionServices();
            hgrRelation = part.GetRelationship("HgrCrossSection", "CrossSection");
            crossSection = (CrossSection)hgrRelation.TargetObjects.First();
            double width, depth, flangeThickness = 0, webThk;
            width = crossSection.Width;
            depth = crossSection.Depth;
            if (crossSection.SupportsInterface("IStructFlangedSectionDimensions"))
            {
                flangeThickness = (double)((PropertyValueDouble)crossSection.GetPropertyValue("IStructFlangedSectionDimensions", "tf")).PropValue;
            }
            webThk = flangeThickness;

            Z1 = webThk * Math.Tan(cutbackBeginAngle);
            Z2 = webThk * Math.Tan(cutbackEndAngle);
            Z3 = width * Math.Tan(cutbackBeginAngle);
            Z4 = width * Math.Tan(cutbackEndAngle);
            Z5 = depth * Math.Tan(cutbackBeginAngle);
            Z6 = depth * Math.Tan(cutbackEndAngle);
            Z7 = flangeThickness * Math.Tan(cutbackBeginAngle);
            Z8 = flangeThickness * Math.Tan(cutbackEndAngle);

            beginOverLength = beginOverLength * Math.Cos(cutbackBeginAngle);
            endOverLength = endOverLength * Math.Cos(cutbackEndAngle);

            if (beginAnchorPoint.Value == 4 && endAnchorPoint.Value == 4)
            {
                if (HgrCompareDoubleService.cmpdbl(cutbackBeginAngle, 0) == true && HgrCompareDoubleService.cmpdbl(cutbackEndAngle, 0) == true)
                    actLength = L + beginOverLength + endOverLength;
                else if (cutbackBeginAngle > 0 && HgrCompareDoubleService.cmpdbl(cutbackEndAngle, 0) == true)
                    actLength = L + beginOverLength + endOverLength + Z3;
                else if (HgrCompareDoubleService.cmpdbl(cutbackBeginAngle, 0) == true && cutbackEndAngle > 0)
                    actLength = L + beginOverLength + endOverLength;
                else if (cutbackBeginAngle > 0 && cutbackEndAngle > 0)
                    actLength = L + beginOverLength + endOverLength + Z3;
                else if (HgrCompareDoubleService.cmpdbl(cutbackBeginAngle, 0) == true && cutbackEndAngle < 0)
                    actLength = L + beginOverLength + endOverLength - Z4;
                else if (cutbackBeginAngle < 0 && HgrCompareDoubleService.cmpdbl(cutbackEndAngle, 0) == true)
                    actLength = L + beginOverLength + endOverLength;
                else if (cutbackBeginAngle < 0 && cutbackEndAngle < 0)
                    actLength = L + beginOverLength + endOverLength - Z4;
                else if (cutbackBeginAngle < 0 && cutbackEndAngle > 0)
                    actLength = L + beginOverLength + endOverLength;
                else if (cutbackBeginAngle > 0 && cutbackEndAngle < 0)
                    actLength = L + beginOverLength + endOverLength + Z3 - Z4;
            }
            else if (beginAnchorPoint.Value == 4 && endAnchorPoint.Value == 6)
            {
                if (HgrCompareDoubleService.cmpdbl(cutbackBeginAngle, 0) == true && HgrCompareDoubleService.cmpdbl(cutbackEndAngle, 0) == true)
                    actLength = L + beginOverLength + endOverLength;
                else if (cutbackBeginAngle > 0 && HgrCompareDoubleService.cmpdbl(cutbackEndAngle, 0) == true)
                    actLength = L + beginOverLength + endOverLength + Z3;
                else if (HgrCompareDoubleService.cmpdbl(cutbackBeginAngle, 0) == true && cutbackEndAngle > 0)
                    actLength = L + beginOverLength + endOverLength + Z4;
                else if (cutbackBeginAngle > 0 && cutbackEndAngle > 0)
                    actLength = L + beginOverLength + endOverLength + Z3 + Z4;
                else if (HgrCompareDoubleService.cmpdbl(cutbackBeginAngle, 0) == true && cutbackEndAngle < 0)
                    actLength = L + beginOverLength + endOverLength;
                else if (cutbackBeginAngle < 0 && HgrCompareDoubleService.cmpdbl(cutbackEndAngle, 0) == true)
                    actLength = L + beginOverLength + endOverLength;
                else if (cutbackBeginAngle < 0 && cutbackEndAngle < 0)
                    actLength = L + beginOverLength + endOverLength;
                else if (cutbackBeginAngle < 0 && cutbackEndAngle > 0)
                    actLength = L + beginOverLength + endOverLength + Z4;
                else if (cutbackBeginAngle > 0 && cutbackEndAngle < 0)
                    actLength = L + beginOverLength + endOverLength + Z3;
            }
            else if (beginAnchorPoint.Value == 6 && endAnchorPoint.Value == 4)
            {
                if (HgrCompareDoubleService.cmpdbl(cutbackBeginAngle, 0) == true && HgrCompareDoubleService.cmpdbl(cutbackEndAngle, 0) == true)
                    actLength = L + beginOverLength + endOverLength;
                else if (cutbackBeginAngle > 0 && HgrCompareDoubleService.cmpdbl(cutbackEndAngle, 0) == true)
                    actLength = L + beginOverLength + endOverLength;
                else if (HgrCompareDoubleService.cmpdbl(cutbackBeginAngle, 0) == true && cutbackEndAngle > 0)
                    actLength = L + beginOverLength + endOverLength;
                else if (cutbackBeginAngle > 0 && cutbackEndAngle > 0)
                    actLength = L + beginOverLength + endOverLength;
                else if (HgrCompareDoubleService.cmpdbl(cutbackBeginAngle, 0) == true && cutbackEndAngle < 0)
                    actLength = L + beginOverLength + endOverLength - Z4;
                else if (cutbackBeginAngle < 0 && HgrCompareDoubleService.cmpdbl(cutbackEndAngle, 0) == true)
                    actLength = L + beginOverLength + endOverLength - Z3;
                else if (cutbackBeginAngle < 0 && cutbackEndAngle < 0)
                    actLength = L + beginOverLength + endOverLength - Z3 - Z4;
                else if (cutbackBeginAngle < 0 && cutbackEndAngle > 0)
                    actLength = L + beginOverLength + endOverLength - Z3;
                else if (cutbackBeginAngle > 0 && cutbackEndAngle < 0)
                    actLength = L + beginOverLength + endOverLength - Z4;
            }
            else if (beginAnchorPoint.Value == 6 && endAnchorPoint.Value == 6)
            {
                if (HgrCompareDoubleService.cmpdbl(cutbackBeginAngle, 0) == true && HgrCompareDoubleService.cmpdbl(cutbackEndAngle, 0) == true)
                    actLength = L + beginOverLength + endOverLength;
                else if (cutbackBeginAngle > 0 && HgrCompareDoubleService.cmpdbl(cutbackEndAngle, 0) == true)
                    actLength = L + beginOverLength + endOverLength;
                else if (HgrCompareDoubleService.cmpdbl(cutbackBeginAngle, 0) == true && cutbackEndAngle > 0)
                    actLength = L + beginOverLength + endOverLength + Z4;
                else if (cutbackBeginAngle > 0 && cutbackEndAngle > 0)
                    actLength = L + beginOverLength + endOverLength + Z4;
                else if (HgrCompareDoubleService.cmpdbl(cutbackBeginAngle, 0) == true && cutbackEndAngle < 0)
                    actLength = L + beginOverLength + endOverLength;
                else if (cutbackBeginAngle < 0 && HgrCompareDoubleService.cmpdbl(cutbackEndAngle, 0) == true)
                    actLength = L + beginOverLength + endOverLength - Z3;
                else if (cutbackBeginAngle < 0 && cutbackEndAngle < 0)
                    actLength = L + beginOverLength + endOverLength - Z3;
                else if (cutbackBeginAngle < 0 && cutbackEndAngle > 0)
                    actLength = L + beginOverLength + endOverLength - Z3 + Z4;
                else if (cutbackBeginAngle > 0 && cutbackEndAngle < 0)
                    actLength = L + beginOverLength + endOverLength;
            }
            else if (beginAnchorPoint.Value == 2 && endAnchorPoint.Value == 2)
            {
                if (HgrCompareDoubleService.cmpdbl(cutbackBeginAngle, 0) == true && HgrCompareDoubleService.cmpdbl(cutbackEndAngle, 0) == true)
                    actLength = L + beginOverLength + endOverLength;
                else if (cutbackBeginAngle > 0 && HgrCompareDoubleService.cmpdbl(cutbackEndAngle, 0) == true)
                    actLength = L + beginOverLength + endOverLength + Z5;
                else if (HgrCompareDoubleService.cmpdbl(cutbackBeginAngle, 0) == true && cutbackEndAngle > 0)
                    actLength = L + beginOverLength + endOverLength;
                else if (cutbackBeginAngle > 0 && cutbackEndAngle > 0)
                    actLength = L + beginOverLength + endOverLength + Z5;
                else if (HgrCompareDoubleService.cmpdbl(cutbackBeginAngle, 0) == true && cutbackEndAngle < 0)
                    actLength = L + beginOverLength + endOverLength - Z6;
                else if (cutbackBeginAngle < 0 && HgrCompareDoubleService.cmpdbl(cutbackEndAngle, 0) == true)
                    actLength = L + beginOverLength + endOverLength;
                else if (cutbackBeginAngle < 0 && cutbackEndAngle < 0)
                    actLength = L + beginOverLength + endOverLength - Z6;
                else if (cutbackBeginAngle < 0 && cutbackEndAngle > 0)
                    actLength = L + beginOverLength + endOverLength;
                else if (cutbackBeginAngle > 0 && cutbackEndAngle < 0)
                    actLength = L + beginOverLength + endOverLength + Z5 - Z6;
            }
            else if (beginAnchorPoint.Value == 2 && endAnchorPoint.Value == 8)
            {
                if (HgrCompareDoubleService.cmpdbl(cutbackBeginAngle, 0) == true && HgrCompareDoubleService.cmpdbl(cutbackEndAngle, 0) == true)
                    actLength = L + beginOverLength + endOverLength;
                else if (cutbackBeginAngle > 0 && HgrCompareDoubleService.cmpdbl(cutbackEndAngle, 0) == true)
                    actLength = L + beginOverLength + endOverLength + Z5;
                else if (HgrCompareDoubleService.cmpdbl(cutbackBeginAngle, 0) == true && cutbackEndAngle > 0)
                    actLength = L + beginOverLength + endOverLength + Z6;
                else if (cutbackBeginAngle > 0 && cutbackEndAngle > 0)
                    actLength = L + beginOverLength + endOverLength + Z5 + Z6;
                else if (HgrCompareDoubleService.cmpdbl(cutbackBeginAngle, 0) == true && cutbackEndAngle < 0)
                    actLength = L + beginOverLength + endOverLength;
                else if (cutbackBeginAngle < 0 && HgrCompareDoubleService.cmpdbl(cutbackEndAngle, 0) == true)
                    actLength = L + beginOverLength + endOverLength;
                else if (cutbackBeginAngle < 0 && cutbackEndAngle < 0)
                    actLength = L + beginOverLength + endOverLength;
                else if (cutbackBeginAngle < 0 && cutbackEndAngle > 0)
                    actLength = L + beginOverLength + endOverLength + Z6;
                else if (cutbackBeginAngle > 0 && cutbackEndAngle < 0)
                    actLength = L + beginOverLength + endOverLength + Z5;
            }
            else if (beginAnchorPoint.Value == 8 && endAnchorPoint.Value == 2)
            {
                if (HgrCompareDoubleService.cmpdbl(cutbackBeginAngle, 0) == true && HgrCompareDoubleService.cmpdbl(cutbackEndAngle, 0) == true)
                    actLength = L + beginOverLength + endOverLength;
                else if (cutbackBeginAngle > 0 && HgrCompareDoubleService.cmpdbl(cutbackEndAngle, 0) == true)
                    actLength = L + beginOverLength + endOverLength;
                else if (HgrCompareDoubleService.cmpdbl(cutbackBeginAngle, 0) == true && cutbackEndAngle > 0)
                    actLength = L + beginOverLength + endOverLength;
                else if (cutbackBeginAngle > 0 && cutbackEndAngle > 0)
                    actLength = L + beginOverLength + endOverLength;
                else if (HgrCompareDoubleService.cmpdbl(cutbackBeginAngle, 0) == true && cutbackEndAngle < 0)
                    actLength = L + beginOverLength + endOverLength - Z6;
                else if (cutbackBeginAngle < 0 && HgrCompareDoubleService.cmpdbl(cutbackEndAngle, 0) == true)
                    actLength = L + beginOverLength + endOverLength - Z5;
                else if (cutbackBeginAngle < 0 && cutbackEndAngle < 0)
                    actLength = L + beginOverLength + endOverLength - Z5 - Z6;
                else if (cutbackBeginAngle < 0 && cutbackEndAngle > 0)
                    actLength = L + beginOverLength + endOverLength - Z5;
                else if (cutbackBeginAngle > 0 && cutbackEndAngle < 0)
                    actLength = L + beginOverLength + endOverLength - Z6;
            }
            else if (beginAnchorPoint.Value == 8 && endAnchorPoint.Value == 8)
            {
                if (HgrCompareDoubleService.cmpdbl(cutbackBeginAngle, 0) == true && HgrCompareDoubleService.cmpdbl(cutbackEndAngle, 0) == true)
                    actLength = L + beginOverLength + endOverLength;
                else if (cutbackBeginAngle > 0 && HgrCompareDoubleService.cmpdbl(cutbackEndAngle, 0) == true)
                    actLength = L + beginOverLength + endOverLength;
                else if (HgrCompareDoubleService.cmpdbl(cutbackBeginAngle, 0) == true && cutbackEndAngle > 0)
                    actLength = L + beginOverLength + endOverLength + Z6;
                else if (cutbackBeginAngle > 0 && cutbackEndAngle > 0)
                    actLength = L + beginOverLength + endOverLength + Z6;
                else if (HgrCompareDoubleService.cmpdbl(cutbackBeginAngle, 0) == true && cutbackEndAngle < 0)
                    actLength = L + beginOverLength + endOverLength;
                else if (cutbackBeginAngle < 0 && HgrCompareDoubleService.cmpdbl(cutbackEndAngle, 0) == true)
                    actLength = L + beginOverLength + endOverLength - Z5;
                else if (cutbackBeginAngle < 0 && cutbackEndAngle < 0)
                    actLength = L + beginOverLength + endOverLength - Z5;
                else if (cutbackBeginAngle < 0 && cutbackEndAngle > 0)
                    actLength = L + beginOverLength + endOverLength - Z5 + Z6;
                else if (cutbackBeginAngle > 0 && cutbackEndAngle < 0)
                    actLength = L + beginOverLength + endOverLength;
            }
            return actLength;
        }
        #endregion

        #region "Calculation Of Length For WCG"

        public double CalculateLengthForWCG(CrossSection crossSection, double length, double beginOverLength, double endOverLength, CodelistItem beginCutbackAnchorPt, CodelistItem endCutbackAnchorPt, double beginCutbackAngle, double endCutbackAngle, double beginFaceZCoord, double endFaceZCoord)
        {
            double actLength = 0;
            double cpOffSetX = 0;
            double cpOffSetY = 0;

            CrossSectionServices crossSectionServices = new CrossSectionServices();
            crossSectionServices.GetCardinalPointOffset(crossSection, 5, out cpOffSetX, out cpOffSetY);
            Double vecX1 = 0;
            Double vecY1 = 0;
            Double vecZ1 = 0;
            Double vecX2 = 0;
            Double vecY2 = 0;
            Double vecZ2 = 0;

            vecX1 = cpOffSetX;
            vecY1 = cpOffSetY;
            vecZ1 = -beginOverLength + beginFaceZCoord;

            vecX2 = cpOffSetX;
            vecY2 = cpOffSetY;
            vecZ2 = length + endOverLength + endFaceZCoord;

            Vector vecForWeight = new Vector();
            vecForWeight.Set(vecX1 - vecX2, vecY1 - vecY2, vecZ1 - vecZ2);

            actLength = vecForWeight.Length;
            return actLength;
        }
        #endregion
    }
}

﻿//-----------------------------------------------------------------------------
//      Copyright (C) 2010-12 Intergraph Corporation.  All rights reserved.
//
//      Component:  SectionLMap class is designed to return a cross section
//                  map based on section type L.
//
//      Author:  3XCalibur
//
//      History:
//      November 03, 2010       WR                  Created
//
//   October 11,2012     Alliagtors      CR-207934/DM-CP-220895(for v2013) 'pointMap' collection is filled appropriately.
//-----------------------------------------------------------------------------

using System.Collections.Generic;
using Ingr.SP3D.Structure.Middle;

namespace Ingr.SP3D.Content.Structure.CrossSectionMappings
{
    /// <summary>
    /// The SectionLMap class is designed to return a cross section edge map based on L type shape cross sections.
    /// </summary>
    internal class SectionLMap : ICrossSectionMap
    {
        #region ICrossSectionMap Members

        /// <summary>
        /// Gets the cross section map.
        /// </summary>
        /// <param name="flipLeftAndRight">if set to <c>true</c> [is mirrored].</param>
        /// <param name="quadrant">The quadrant.</param>
        /// <returns>The cross section mapping of edges based on the quadrant and the flipLeftAndRight flag.</returns>
        /// <exception cref="S3DEndCutMappingException">Unable to map cross section.</exception>
        public Dictionary<SectionFaceType, SectionFaceType> GetCrossSectionMap(bool flipLeftAndRight, int quadrant, Dictionary<int, int> pointMap)
        {
            Dictionary<SectionFaceType, SectionFaceType> mappedEdges = new Dictionary<SectionFaceType, SectionFaceType>();

            //      WR = Web_Right    
            //      Top
            //      WL = Web_Left
            //      Bottom
            //      BFR = Bottom_Flange_Right
            //      BFRT = Bottom_Flange_Right_Top
            //
            //             Top
            //           ******
            //           *    * 
            //           *    *
            //           *    *  
            //           *    *
            //        WL *    *  WR
            //           *    *
            //           *    *  BFRT
            //           *    **********
            //           *             * BFR 
            //           ***************
            //              Bottom

            if (!flipLeftAndRight)
            {
                // Not mirrored case
                switch (quadrant)
                {
                    case 1:
                        mappedEdges.Add(SectionFaceType.Web_Right, SectionFaceType.Web_Right);
                        mappedEdges.Add(SectionFaceType.Top, SectionFaceType.Top);
                        mappedEdges.Add(SectionFaceType.Web_Left, SectionFaceType.Web_Left);
                        mappedEdges.Add(SectionFaceType.Bottom, SectionFaceType.Bottom);
                        mappedEdges.Add(SectionFaceType.Bottom_Flange_Right, SectionFaceType.Bottom_Flange_Right);
                        mappedEdges.Add(SectionFaceType.Bottom_Flange_Right_Top, SectionFaceType.Bottom_Flange_Right_Top);

                        for (int i = 1; i < 28; i++)
                        {
                            pointMap.Add(i, i);
                        }
                        
                        break;
                    
                    case 2:
                        mappedEdges.Add(SectionFaceType.Web_Right, SectionFaceType.Bottom);
                        mappedEdges.Add(SectionFaceType.Top, SectionFaceType.Bottom_Flange_Right);
                        mappedEdges.Add(SectionFaceType.Web_Left, SectionFaceType.Bottom_Flange_Right_Top);
                        mappedEdges.Add(SectionFaceType.Bottom_Flange_Left_Top, SectionFaceType.Web_Right);
                        mappedEdges.Add(SectionFaceType.Bottom_Flange_Left, SectionFaceType.Top);
                        mappedEdges.Add(SectionFaceType.Bottom, SectionFaceType.Web_Left);

                        pointMap.Add(1, -1);
                        pointMap.Add(2, 5);
                        pointMap.Add(3, 8);
                        pointMap.Add(4, 13);
                        pointMap.Add(5, 14);
                        pointMap.Add(6, 20);
                        pointMap.Add(7, -1);
                        pointMap.Add(8, 21);
                        pointMap.Add(9, 21);
                        pointMap.Add(10, 21);
                        pointMap.Add(11, -1);
                        pointMap.Add(12, 21);
                        pointMap.Add(13, 22);
                        pointMap.Add(14, 23);
                        pointMap.Add(15, 23);
                        pointMap.Add(16, 23);
                        pointMap.Add(17, 23);
                        pointMap.Add(18, 23);
                        pointMap.Add(19, -1);
                        pointMap.Add(20, 24);
                        pointMap.Add(21, 24);
                        pointMap.Add(22, 2);
                        pointMap.Add(23, 2);
                        pointMap.Add(24, 2);
                        pointMap.Add(25, -1);
                        pointMap.Add(26, -1);
                        pointMap.Add(27, 27);
                        
                        break;

                    case 3:
                        mappedEdges.Add(SectionFaceType.Web_Right, SectionFaceType.Web_Left);
                        mappedEdges.Add(SectionFaceType.Top, SectionFaceType.Bottom);
                        mappedEdges.Add(SectionFaceType.Top_Flange_Left, SectionFaceType.Bottom_Flange_Right);
                        mappedEdges.Add(SectionFaceType.Top_Flange_Left_Bottom, SectionFaceType.Bottom_Flange_Right_Top);
                        mappedEdges.Add(SectionFaceType.Web_Left, SectionFaceType.Web_Right);
                        mappedEdges.Add(SectionFaceType.Bottom, SectionFaceType.Top);
                        
                        pointMap.Add(1, 13);
                        pointMap.Add(2, 14);
                        pointMap.Add(3, -1);
                        pointMap.Add(4, 14);
                        pointMap.Add(5, 14);
                        pointMap.Add(6, 14);
                        pointMap.Add(7, 19);
                        pointMap.Add(8, 20);
                        pointMap.Add(9, 21);
                        pointMap.Add(10, 22);
                        pointMap.Add(11, 23);
                        pointMap.Add(12, 24);
                        pointMap.Add(13, 1);
                        pointMap.Add(14, 2);
                        pointMap.Add(15, 2);
                        pointMap.Add(16, 2);
                        pointMap.Add(17, 5);
                        pointMap.Add(18, 5);
                        pointMap.Add(19, 7);
                        pointMap.Add(20, 8);
                        pointMap.Add(21, 8);
                        pointMap.Add(22, 8);
                        pointMap.Add(23, 8);
                        pointMap.Add(24, 8);
                        pointMap.Add(25, 26);
                        pointMap.Add(26, 25);
                        pointMap.Add(27, 27);
                        
                        break;

                    case 4:
                        mappedEdges.Add(SectionFaceType.Web_Right, SectionFaceType.Bottom_Flange_Right_Top);
                        mappedEdges.Add(SectionFaceType.Top_Flange_Right_Bottom, SectionFaceType.Web_Right);
                        mappedEdges.Add(SectionFaceType.Top_Flange_Right, SectionFaceType.Top);
                        mappedEdges.Add(SectionFaceType.Top, SectionFaceType.Web_Left);
                        mappedEdges.Add(SectionFaceType.Web_Left, SectionFaceType.Bottom);
                        mappedEdges.Add(SectionFaceType.Bottom, SectionFaceType.Bottom_Flange_Right);
                        
                        pointMap.Add(1, 22);
                        pointMap.Add(2, 23);
                        pointMap.Add(3, 23);
                        pointMap.Add(4, 23);
                        pointMap.Add(5, 23);
                        pointMap.Add(6, 23);
                        pointMap.Add(7, -1);
                        pointMap.Add(8, 2);
                        pointMap.Add(9, 2);
                        pointMap.Add(10, 2);
                        pointMap.Add(11, 2);
                        pointMap.Add(12, 2);
                        pointMap.Add(13, -1);
                        pointMap.Add(14, 5);
                        pointMap.Add(15, 8);
                        pointMap.Add(16, 13);
                        pointMap.Add(17, 14);
                        pointMap.Add(18, 20);
                        pointMap.Add(19, -1);
                        pointMap.Add(20, 21);
                        pointMap.Add(21, 21);
                        pointMap.Add(22, 21);
                        pointMap.Add(23, -1);
                        pointMap.Add(24, 21);
                        pointMap.Add(25, -1);
                        pointMap.Add(26, -1);
                        pointMap.Add(27, 27); 
                        
                        break;

                    default:
                        throw new EdgeMappingException(quadrant);
                }
            }
            else
            {
                switch (quadrant)
                {
                    case 1:
                        mappedEdges.Add(SectionFaceType.Web_Right, SectionFaceType.Web_Left);
                        mappedEdges.Add(SectionFaceType.Top, SectionFaceType.Top);
                        mappedEdges.Add(SectionFaceType.Web_Left, SectionFaceType.Web_Right);
                        mappedEdges.Add(SectionFaceType.Bottom_Flange_Left_Top, SectionFaceType.Bottom_Flange_Right_Top);
                        mappedEdges.Add(SectionFaceType.Bottom_Flange_Left, SectionFaceType.Bottom_Flange_Right);
                        mappedEdges.Add(SectionFaceType.Bottom, SectionFaceType.Bottom);

                        pointMap.Add(1, 1);
                        pointMap.Add(2, 24);
                        pointMap.Add(3, 23);
                        pointMap.Add(4, 22);
                        pointMap.Add(5, 21);
                        pointMap.Add(6, 20);
                        pointMap.Add(7, 19);
                        pointMap.Add(8, 18);
                        pointMap.Add(9, 17);
                        pointMap.Add(10, 16);
                        pointMap.Add(11, -1);
                        pointMap.Add(12, 14);
                        pointMap.Add(13, 13);
                        pointMap.Add(14, 12);
                        pointMap.Add(15, 11);
                        pointMap.Add(16, 10);
                        pointMap.Add(17, 9);
                        pointMap.Add(18, 8);
                        pointMap.Add(19, 7);
                        pointMap.Add(20, 6);
                        pointMap.Add(21, 5);
                        pointMap.Add(22, 4);
                        pointMap.Add(23, 3);
                        pointMap.Add(24, 2);
                        pointMap.Add(25, 26);
                        pointMap.Add(26, 25);
                        pointMap.Add(27, 27);

                        break;
                    
                    case 2:
                        mappedEdges.Add(SectionFaceType.Web_Right, SectionFaceType.Bottom);
                        mappedEdges.Add(SectionFaceType.Top, SectionFaceType.Web_Left);
                        mappedEdges.Add(SectionFaceType.Top_Flange_Left, SectionFaceType.Top);
                        mappedEdges.Add(SectionFaceType.Top_Flange_Left_Bottom, SectionFaceType.Web_Right);
                        mappedEdges.Add(SectionFaceType.Web_Left, SectionFaceType.Bottom_Flange_Right_Top);
                        mappedEdges.Add(SectionFaceType.Bottom, SectionFaceType.Bottom_Flange_Right);
                        
                        pointMap.Add(1, 22);
                        pointMap.Add(2, 21);
                        pointMap.Add(3, -1);
                        pointMap.Add(4, 21);
                        pointMap.Add(5, 21);
                        pointMap.Add(6, 21);
                        pointMap.Add(7, -1);
                        pointMap.Add(8, 20);
                        pointMap.Add(9, 14);
                        pointMap.Add(10, 13);
                        pointMap.Add(11, 8);
                        pointMap.Add(12, 5);
                        pointMap.Add(13, -1);
                        pointMap.Add(14, 2);
                        pointMap.Add(15, 2);
                        pointMap.Add(16, 2);
                        pointMap.Add(17, 2);
                        pointMap.Add(18, 2);
                        pointMap.Add(19, -1);
                        pointMap.Add(20, 23);
                        pointMap.Add(21, 23);
                        pointMap.Add(22, 23);
                        pointMap.Add(23, 23);
                        pointMap.Add(24, 23);
                        pointMap.Add(25, -1);
                        pointMap.Add(26, -1);
                        pointMap.Add(27, 27); 
                        
                        break;

                    case 3:
                        mappedEdges.Add(SectionFaceType.Web_Right, SectionFaceType.Web_Right);
                        mappedEdges.Add(SectionFaceType.Top_Flange_Right_Bottom, SectionFaceType.Bottom_Flange_Right_Top);
                        mappedEdges.Add(SectionFaceType.Top_Flange_Right, SectionFaceType.Bottom_Flange_Right);
                        mappedEdges.Add(SectionFaceType.Top, SectionFaceType.Bottom);
                        mappedEdges.Add(SectionFaceType.Web_Left, SectionFaceType.Web_Left);
                        mappedEdges.Add(SectionFaceType.Bottom, SectionFaceType.Top);
                        
                        pointMap.Add(1, 13);
                        pointMap.Add(2, 8);
                        pointMap.Add(3, 8);
                        pointMap.Add(4, 8);
                        pointMap.Add(5, 8);
                        pointMap.Add(6, 8);
                        pointMap.Add(7, 7);
                        pointMap.Add(8, 5);
                        pointMap.Add(9, 5);
                        pointMap.Add(10, 2);
                        pointMap.Add(11, 2);
                        pointMap.Add(12, 2);
                        pointMap.Add(13, 1);
                        pointMap.Add(14, 24);
                        pointMap.Add(15, 23);
                        pointMap.Add(16, 22);
                        pointMap.Add(17, 21);
                        pointMap.Add(18, 20);
                        pointMap.Add(19, 19);
                        pointMap.Add(20, 14);
                        pointMap.Add(21, 14);
                        pointMap.Add(22, 14);
                        pointMap.Add(23, -1);
                        pointMap.Add(24, 14);
                        pointMap.Add(25, 25);
                        pointMap.Add(26, 26);
                        pointMap.Add(27, 27); 
                        
                        break;

                    case 4:
                        mappedEdges.Add(SectionFaceType.Web_Right, SectionFaceType.Bottom_Flange_Right_Top);
                        mappedEdges.Add(SectionFaceType.Top, SectionFaceType.Bottom_Flange_Right);
                        mappedEdges.Add(SectionFaceType.Web_Left, SectionFaceType.Bottom);
                        mappedEdges.Add(SectionFaceType.Bottom, SectionFaceType.Web_Left);
                        mappedEdges.Add(SectionFaceType.Bottom_Flange_Right, SectionFaceType.Top);
                        mappedEdges.Add(SectionFaceType.Bottom_Flange_Right_Top, SectionFaceType.Web_Right);
                        
                        pointMap.Add(1, -1);
                        pointMap.Add(2, 2);
                        pointMap.Add(3, 2);
                        pointMap.Add(4, 2);
                        pointMap.Add(5, 24);
                        pointMap.Add(6, 24);
                        pointMap.Add(7, -1);
                        pointMap.Add(8, 23);
                        pointMap.Add(9, 23);
                        pointMap.Add(10, 23);
                        pointMap.Add(11, 23);
                        pointMap.Add(12, 23);
                        pointMap.Add(13, 22);
                        pointMap.Add(14, 21);
                        pointMap.Add(15, -1);
                        pointMap.Add(16, 21);
                        pointMap.Add(17, 21);
                        pointMap.Add(18, 21);
                        pointMap.Add(19, -1);
                        pointMap.Add(20, 20);
                        pointMap.Add(21, 14);
                        pointMap.Add(22, 13);
                        pointMap.Add(23, 8);
                        pointMap.Add(24, 5);
                        pointMap.Add(25, -1);
                        pointMap.Add(26, -1);
                        pointMap.Add(27, -1);
                        
                        break;

                    default:
                        throw new EdgeMappingException(quadrant);
                }
            }

            return mappedEdges;
        }

        #endregion
    }
}

//-----------------------------------------------------------------------------
//      Copyright (C) 2010-12 Intergraph Corporation.  All rights reserved.
//
//      Component:  SectionBUTL3Map class is designed to return a cross section
//                  map based on section type BUTL3.
//
//      Author:  Alligators
//
//      History:
//      November 09, 2011       RM                  Created
//
//   October 11,2012     Alliagtors      CR-207934/DM-CP-220895(for v2013) 'pointMap' collection is filled appropriately.
//-----------------------------------------------------------------------------

using System.Collections.Generic;
using Ingr.SP3D.Structure.Middle;

namespace Ingr.SP3D.Content.Structure.CrossSectionMappings
{
    /// <summary>
    /// The SectionBUTL3Map class is designed to return a cross section edge map based on BUTL3 type shape cross sections.
    /// </summary>
    internal class SectionBUTL3Map : ICrossSectionMap
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

            //      T = Top
            //      WRT = Web_Right_Top
            //      TFRT = Top_Flange_Right_Top
            //      TFRB = Top_Flange_Right_Bottom
            //      TFR = Top_Flange_Right
            //      WR = Web_Right
            //      WL = Web_Left         
            //      B = Bottom
            //
            //                       T
            //                     ******    
            //                     *    * WRT
            //                     *    * 
            //                     *    *  TFRT
            //                     *    ***********
            //                     *              * TFR
            //                     *    ***********
            //                 WL  *    *  TFRB
            //                     *    *
            //                     *    *  WR
            //                     *    *
            //                     *    *  
            //                     *    *
            //                     *    *  
            //                     ******
            //                       B

            if (!flipLeftAndRight)
            {
                // Not mirrored case
                switch (quadrant)
                {
                    case 1:
                        mappedEdges.Add(SectionFaceType.Top, SectionFaceType.Top);
                        mappedEdges.Add(SectionFaceType.Web_Right_Top, SectionFaceType.Web_Right_Top);
                        mappedEdges.Add(SectionFaceType.Top_Flange_Right_Top, SectionFaceType.Top_Flange_Right_Top);
                        mappedEdges.Add(SectionFaceType.Top_Flange_Right, SectionFaceType.Top_Flange_Right);
                        mappedEdges.Add(SectionFaceType.Top_Flange_Right_Bottom, SectionFaceType.Top_Flange_Right_Bottom);
                        mappedEdges.Add(SectionFaceType.Web_Right, SectionFaceType.Web_Right);
                        mappedEdges.Add(SectionFaceType.Bottom, SectionFaceType.Bottom);
                        mappedEdges.Add(SectionFaceType.Web_Left, SectionFaceType.Web_Left);
                        
                        for (int i = 1; i < 28; i++)
                        {
                            pointMap.Add(i, i);
                        }

                        pointMap.Add(50, 50);

                        break;
                    
                    case 2:
                        mappedEdges.Add(SectionFaceType.Bottom_Flange_Left, SectionFaceType.Top);
                        mappedEdges.Add(SectionFaceType.Bottom_Flange_Left_Top, SectionFaceType.Web_Right_Top);
                        mappedEdges.Add(SectionFaceType.Web_Left, SectionFaceType.Top_Flange_Right_Top);
                        mappedEdges.Add(SectionFaceType.Top, SectionFaceType.Top_Flange_Right);
                        mappedEdges.Add(SectionFaceType.Web_Right, SectionFaceType.Top_Flange_Right_Bottom);
                        mappedEdges.Add(SectionFaceType.Bottom_Flange_Right_Top, SectionFaceType.Web_Right);
                        mappedEdges.Add(SectionFaceType.Bottom_Flange_Right, SectionFaceType.Bottom);
                        mappedEdges.Add(SectionFaceType.Bottom, SectionFaceType.Web_Left);
                        
                        pointMap.Add(1, -1);
                        pointMap.Add(2, -1);
                        pointMap.Add(3, 8);
                        pointMap.Add(4, 13);
                        pointMap.Add(5, 14);
                        pointMap.Add(6, 50);
                        pointMap.Add(7, -1);
                        pointMap.Add(8, 15);
                        pointMap.Add(9, 15);
                        pointMap.Add(10, 15);
                        pointMap.Add(11, 15);
                        pointMap.Add(12, 15);
                        pointMap.Add(13, 16);
                        pointMap.Add(14, 17);
                        pointMap.Add(15, 17);
                        pointMap.Add(16, 17);
                        pointMap.Add(17, 17);
                        pointMap.Add(18, 17);
                        pointMap.Add(19, -1);
                        pointMap.Add(20, 18);
                        pointMap.Add(21, 20);
                        pointMap.Add(22, 1);
                        pointMap.Add(23, 2);
                        pointMap.Add(24, -1);
                        pointMap.Add(25, -1);
                        pointMap.Add(26, -1);
                        pointMap.Add(27, 27);

                        break;
                    
                    case 3:
                        mappedEdges.Add(SectionFaceType.Bottom, SectionFaceType.Top);
                        mappedEdges.Add(SectionFaceType.Web_Left_Bottom, SectionFaceType.Web_Right_Top);
                        mappedEdges.Add(SectionFaceType.Bottom_Flange_Left_Bottom, SectionFaceType.Top_Flange_Right_Top);
                        mappedEdges.Add(SectionFaceType.Bottom_Flange_Left, SectionFaceType.Top_Flange_Right);
                        mappedEdges.Add(SectionFaceType.Bottom_Flange_Left_Top, SectionFaceType.Top_Flange_Right_Bottom);
                        mappedEdges.Add(SectionFaceType.Web_Left, SectionFaceType.Web_Right);
                        mappedEdges.Add(SectionFaceType.Top, SectionFaceType.Bottom);
                        mappedEdges.Add(SectionFaceType.Web_Right, SectionFaceType.Web_Left);
                        
                        pointMap.Add(1, 13);
                        pointMap.Add(2, 14);
                        pointMap.Add(3, 15);
                        pointMap.Add(4, 16);
                        pointMap.Add(5, 17);
                        pointMap.Add(6, 18);
                        pointMap.Add(7, 19);
                        pointMap.Add(8, 20);
                        pointMap.Add(9, 20);
                        pointMap.Add(10, 20);
                        pointMap.Add(11, 23);
                        pointMap.Add(12, 20);
                        pointMap.Add(13,1 );
                        pointMap.Add(14, 2);
                        pointMap.Add(15, 2);
                        pointMap.Add(16, 2);
                        pointMap.Add(17, 2);
                        pointMap.Add(18, 2);
                        pointMap.Add(19, 7);
                        pointMap.Add(20, 8);
                        pointMap.Add(21, 8);
                        pointMap.Add(22, 8);
                        pointMap.Add(23, 8);
                        pointMap.Add(24, 8);
                        pointMap.Add(25, 26);
                        pointMap.Add(26, 25);
                        pointMap.Add(27, 27);
                        pointMap.Add(53, 50);
                        break;
                    
                    case 4:
                        mappedEdges.Add(SectionFaceType.Top_Flange_Right, SectionFaceType.Top);
                        mappedEdges.Add(SectionFaceType.Top_Flange_Right_Bottom, SectionFaceType.Web_Right_Top);
                        mappedEdges.Add(SectionFaceType.Web_Right, SectionFaceType.Top_Flange_Right_Top);
                        mappedEdges.Add(SectionFaceType.Bottom, SectionFaceType.Top_Flange_Right);
                        mappedEdges.Add(SectionFaceType.Web_Left, SectionFaceType.Top_Flange_Right_Bottom);
                        mappedEdges.Add(SectionFaceType.Top_Flange_Left_Bottom, SectionFaceType.Web_Right);
                        mappedEdges.Add(SectionFaceType.Top_Flange_Left, SectionFaceType.Bottom);
                        mappedEdges.Add(SectionFaceType.Top, SectionFaceType.Web_Left);
                        
                        pointMap.Add(1, 16);
                        pointMap.Add(2, 17);
                        pointMap.Add(3, 17);
                        pointMap.Add(4, 17);
                        pointMap.Add(5, 17);
                        pointMap.Add(6, 17);
                        pointMap.Add(7, -1);
                        pointMap.Add(8, 18);
                        pointMap.Add(9, 20);
                        pointMap.Add(10, 1);
                        pointMap.Add(11, 2);
                        pointMap.Add(12, -1);
                        pointMap.Add(13, -1);
                        pointMap.Add(14, -1);
                        pointMap.Add(15, 8);
                        pointMap.Add(16, 13);
                        pointMap.Add(17, 14);
                        pointMap.Add(18, 50);
                        pointMap.Add(19, -1);
                        pointMap.Add(20, 15);
                        pointMap.Add(21, 15);
                        pointMap.Add(22, 15);
                        pointMap.Add(23, 15);
                        pointMap.Add(24, 15);
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
                        mappedEdges.Add(SectionFaceType.Top, SectionFaceType.Top);
                        mappedEdges.Add(SectionFaceType.Web_Left_Top, SectionFaceType.Web_Right_Top);
                        mappedEdges.Add(SectionFaceType.Top_Flange_Left_Top, SectionFaceType.Top_Flange_Right_Top);
                        mappedEdges.Add(SectionFaceType.Top_Flange_Left, SectionFaceType.Top_Flange_Right);
                        mappedEdges.Add(SectionFaceType.Top_Flange_Left_Bottom, SectionFaceType.Top_Flange_Right_Bottom);
                        mappedEdges.Add(SectionFaceType.Web_Left, SectionFaceType.Web_Right);
                        mappedEdges.Add(SectionFaceType.Bottom, SectionFaceType.Bottom);
                        mappedEdges.Add(SectionFaceType.Web_Right, SectionFaceType.Web_Left);
                        
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
                        pointMap.Add(11, 15);
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
                        pointMap.Add(52, 50);

                        break;
                    
                    case 2:
                        mappedEdges.Add(SectionFaceType.Top_Flange_Left, SectionFaceType.Top);
                        mappedEdges.Add(SectionFaceType.Top_Flange_Left_Bottom, SectionFaceType.Web_Right_Top);
                        mappedEdges.Add(SectionFaceType.Web_Left, SectionFaceType.Top_Flange_Right_Top);
                        mappedEdges.Add(SectionFaceType.Bottom, SectionFaceType.Top_Flange_Right);
                        mappedEdges.Add(SectionFaceType.Web_Right, SectionFaceType.Top_Flange_Right_Bottom);
                        mappedEdges.Add(SectionFaceType.Top_Flange_Right_Bottom, SectionFaceType.Web_Right);
                        mappedEdges.Add(SectionFaceType.Top_Flange_Right, SectionFaceType.Bottom);
                        mappedEdges.Add(SectionFaceType.Top, SectionFaceType.Web_Left);
                        
                        pointMap.Add(1, 16);
                        pointMap.Add(2, 15);
                        pointMap.Add(3, 15);
                        pointMap.Add(4, 15);
                        pointMap.Add(5, 15);
                        pointMap.Add(6, 15);
                        pointMap.Add(7, -1);
                        pointMap.Add(8, 50);
                        pointMap.Add(9, 14);
                        pointMap.Add(10, 13);
                        pointMap.Add(11, 8);
                        pointMap.Add(12, -1);
                        pointMap.Add(13, -1);
                        pointMap.Add(14, -1);
                        pointMap.Add(15, 2);
                        pointMap.Add(16, 1);
                        pointMap.Add(17, 20);
                        pointMap.Add(18, 18);
                        pointMap.Add(19, -1);
                        pointMap.Add(20, 17);
                        pointMap.Add(21, 17);
                        pointMap.Add(22, 17);
                        pointMap.Add(23, 17);
                        pointMap.Add(24, 17);
                        pointMap.Add(25, -1);
                        pointMap.Add(26, -1);
                        pointMap.Add(27, 27);

                        break;
                    
                    case 3:
                        mappedEdges.Add(SectionFaceType.Bottom, SectionFaceType.Top);
                        mappedEdges.Add(SectionFaceType.Web_Right_Bottom, SectionFaceType.Web_Right_Top);
                        mappedEdges.Add(SectionFaceType.Bottom_Flange_Right_Bottom, SectionFaceType.Top_Flange_Right_Top);
                        mappedEdges.Add(SectionFaceType.Bottom_Flange_Right, SectionFaceType.Top_Flange_Right);
                        mappedEdges.Add(SectionFaceType.Bottom_Flange_Right_Top, SectionFaceType.Top_Flange_Right_Bottom);
                        mappedEdges.Add(SectionFaceType.Web_Right, SectionFaceType.Web_Right);
                        mappedEdges.Add(SectionFaceType.Top, SectionFaceType.Bottom);
                        mappedEdges.Add(SectionFaceType.Web_Left, SectionFaceType.Web_Left);
                        
                        pointMap.Add(1, 13);
                        pointMap.Add(2, 8);
                        pointMap.Add(3, 8);
                        pointMap.Add(4, 8);
                        pointMap.Add(5, 8);
                        pointMap.Add(6, 8);
                        pointMap.Add(7, 7);
                        pointMap.Add(8, 2);
                        pointMap.Add(9, 2);
                        pointMap.Add(10, 2);
                        pointMap.Add(11, 2);
                        pointMap.Add(12, 2);
                        pointMap.Add(13, 1);
                        pointMap.Add(14, 20);
                        pointMap.Add(15, 23);
                        pointMap.Add(16, 20);
                        pointMap.Add(17, 20);
                        pointMap.Add(18, 20);
                        pointMap.Add(19, 19);
                        pointMap.Add(20, 18);
                        pointMap.Add(21, 17);
                        pointMap.Add(22, 16);
                        pointMap.Add(23, 15);
                        pointMap.Add(24, 14);
                        pointMap.Add(25, 25);
                        pointMap.Add(26, 26);
                        pointMap.Add(27, 27);
                        pointMap.Add(51, 50);

                        break;

                    case 4:
                        mappedEdges.Add(SectionFaceType.Bottom_Flange_Right, SectionFaceType.Top);
                        mappedEdges.Add(SectionFaceType.Bottom_Flange_Right_Top, SectionFaceType.Web_Right_Top);
                        mappedEdges.Add(SectionFaceType.Web_Right, SectionFaceType.Top_Flange_Right_Top);
                        mappedEdges.Add(SectionFaceType.Top, SectionFaceType.Top_Flange_Right);
                        mappedEdges.Add(SectionFaceType.Web_Left, SectionFaceType.Top_Flange_Right_Bottom);
                        mappedEdges.Add(SectionFaceType.Bottom_Flange_Left_Top, SectionFaceType.Web_Right);
                        mappedEdges.Add(SectionFaceType.Bottom_Flange_Left, SectionFaceType.Bottom);
                        mappedEdges.Add(SectionFaceType.Bottom, SectionFaceType.Web_Left);
                    
                        pointMap.Add(1, -1);
                        pointMap.Add(2, -1);
                        pointMap.Add(3, 2);
                        pointMap.Add(4, 1);
                        pointMap.Add(5, 2);
                        pointMap.Add(6, 18);
                        pointMap.Add(7, -1);
                        pointMap.Add(8, 17);
                        pointMap.Add(9, 17);
                        pointMap.Add(10, 17);
                        pointMap.Add(11, 17);
                        pointMap.Add(12, 17);
                        pointMap.Add(13, 16);
                        pointMap.Add(14, 15);
                        pointMap.Add(15, 15);
                        pointMap.Add(16, 15);
                        pointMap.Add(17, 15);
                        pointMap.Add(18, 15);
                        pointMap.Add(19, -1);
                        pointMap.Add(20, 50);
                        pointMap.Add(21, 14);
                        pointMap.Add(22, 13);
                        pointMap.Add(23, 8);
                        pointMap.Add(24, -1);
                        pointMap.Add(25, -1);
                        pointMap.Add(26, -1);
                        pointMap.Add(27, 27);

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

//-----------------------------------------------------------------------------
//      Copyright (C) 2010-12 Intergraph Corporation.  All rights reserved.
//
//      Component:  Section2LMap class is designed to return a cross section
//                  map based on section type 2L.  
//
//      Author:  3XCalibur
//
//      History:
//      November 03, 2010       WR                  Created
//
//      October 11,2012     Alliagtors      CR-207934/DM-CP-220895(for v2013) 
//                        'pointMap' collection is filled for non-flipLeftAndRight and quadrant = 1 case.
//-----------------------------------------------------------------------------

using System.Collections.Generic;
using Ingr.SP3D.Structure.Middle;

namespace Ingr.SP3D.Content.Structure.CrossSectionMappings
{
    /// <summary>
    /// The Section2LMap class is designed to return a cross section edge map based on 2L type shape cross sections.
    /// </summary>
    internal class Section2LMap : ICrossSectionMap
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
            //      BFLT = Bottom_Flange_Left_Top
            //      BFL = Bottom_Flange_Left
            //      Bottom
            //      BFR = Bottom_Flange_Right
            //      BFRT = Bottom_Flange_Right_Top
            //
            //             Top
            //           *******
            //           *     *   
            //        WL *     * WR
            //           *     *
            //      BFLT *     * BFRT 
            //     *******     *******
            // BFL *                 * BFR
            //     *******************
            //           Bottom

            if (!flipLeftAndRight)
            {
                // Not mirrored case
                switch (quadrant)
                {
                    case 1:
                        mappedEdges.Add(SectionFaceType.Web_Right, SectionFaceType.Web_Right);
                        mappedEdges.Add(SectionFaceType.Top, SectionFaceType.Top);
                        mappedEdges.Add(SectionFaceType.Web_Left, SectionFaceType.Web_Left);
                        mappedEdges.Add(SectionFaceType.Bottom_Flange_Left_Top, SectionFaceType.Bottom_Flange_Left_Top);
                        mappedEdges.Add(SectionFaceType.Bottom_Flange_Left, SectionFaceType.Bottom_Flange_Left);
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
                        mappedEdges.Add(SectionFaceType.Right_Web_Top, SectionFaceType.Bottom_Flange_Right);
                        mappedEdges.Add(SectionFaceType.Inner_Web_Right_Top, SectionFaceType.Bottom_Flange_Right_Top);
                        mappedEdges.Add(SectionFaceType.Top, SectionFaceType.Web_Right);
                        mappedEdges.Add(SectionFaceType.Web_Left, SectionFaceType.Top);
                        mappedEdges.Add(SectionFaceType.Bottom, SectionFaceType.Web_Left);
                        mappedEdges.Add(SectionFaceType.Inner_Web_Right_Bottom, SectionFaceType.Bottom_Flange_Left_Top);
                        mappedEdges.Add(SectionFaceType.Right_Web_Bottom, SectionFaceType.Bottom_Flange_Left);
                        
                        //pointMap.Add(1, );
                        //pointMap.Add(2, );
                        //pointMap.Add(3, );
                        //pointMap.Add(4, );
                        //pointMap.Add(5, );
                        //pointMap.Add(6, );
                        //pointMap.Add(7, );
                        //pointMap.Add(8, );
                        //pointMap.Add(9, );
                        //pointMap.Add(10, );
                        //pointMap.Add(11, );
                        //pointMap.Add(12, );
                        //pointMap.Add(13, );
                        //pointMap.Add(14, );
                        //pointMap.Add(15, );
                        //pointMap.Add(16, );
                        //pointMap.Add(17, );
                        //pointMap.Add(18, );
                        //pointMap.Add(19, );
                        //pointMap.Add(20, );
                        //pointMap.Add(21, );
                        //pointMap.Add(22, );
                        //pointMap.Add(23, );
                        //pointMap.Add(24, );
                        //pointMap.Add(25, );
                        //pointMap.Add(26, );
                        //pointMap.Add(27, 27);

                        break;
                    
                    case 3:
                        mappedEdges.Add(SectionFaceType.Web_Right, SectionFaceType.Web_Left);
                        mappedEdges.Add(SectionFaceType.Top_Flange_Right_Bottom, SectionFaceType.Bottom_Flange_Left_Top);
                        mappedEdges.Add(SectionFaceType.Top_Flange_Right, SectionFaceType.Bottom_Flange_Left);
                        mappedEdges.Add(SectionFaceType.Top, SectionFaceType.Bottom);
                        mappedEdges.Add(SectionFaceType.Top_Flange_Left, SectionFaceType.Bottom_Flange_Right);
                        mappedEdges.Add(SectionFaceType.Top_Flange_Left_Bottom, SectionFaceType.Bottom_Flange_Right_Top);
                        mappedEdges.Add(SectionFaceType.Web_Left, SectionFaceType.Web_Right);
                        mappedEdges.Add(SectionFaceType.Bottom, SectionFaceType.Top);
                        
                        //pointMap.Add(1, );
                        //pointMap.Add(2, );
                        //pointMap.Add(3, );
                        //pointMap.Add(4, );
                        //pointMap.Add(5, );
                        //pointMap.Add(6, );
                        //pointMap.Add(7, );
                        //pointMap.Add(8, );
                        //pointMap.Add(9, );
                        //pointMap.Add(10, );
                        //pointMap.Add(11, );
                        //pointMap.Add(12, );
                        //pointMap.Add(13, );
                        //pointMap.Add(14, );
                        //pointMap.Add(15, );
                        //pointMap.Add(16, );
                        //pointMap.Add(17, );
                        //pointMap.Add(18, );
                        //pointMap.Add(19, );
                        //pointMap.Add(20, );
                        //pointMap.Add(21, );
                        //pointMap.Add(22, );
                        //pointMap.Add(23, );
                        //pointMap.Add(24, );
                        //pointMap.Add(25, );
                        //pointMap.Add(26, );
                        //pointMap.Add(27, 27);

                        break;
                    
                    case 4:
                        mappedEdges.Add(SectionFaceType.Web_Right, SectionFaceType.Bottom_Flange_Left_Top);
                        mappedEdges.Add(SectionFaceType.Top, SectionFaceType.Bottom_Flange_Left);
                        mappedEdges.Add(SectionFaceType.Web_Left, SectionFaceType.Bottom);
                        mappedEdges.Add(SectionFaceType.Bottom, SectionFaceType.Bottom_Flange_Right);
                        mappedEdges.Add(SectionFaceType.Web_Right_Bottom, SectionFaceType.Bottom_Flange_Right_Top);
                        mappedEdges.Add(SectionFaceType.Bottom_Flange_Right_Bottom, SectionFaceType.Web_Right);
                        mappedEdges.Add(SectionFaceType.Bottom_Flange_Right, SectionFaceType.Top);
                        mappedEdges.Add(SectionFaceType.Bottom_Flange_Right_Top, SectionFaceType.Web_Left);
                        
                        //pointMap.Add(1, );
                        //pointMap.Add(2, );
                        //pointMap.Add(3, );
                        //pointMap.Add(4, );
                        //pointMap.Add(5, );
                        //pointMap.Add(6, );
                        //pointMap.Add(7, );
                        //pointMap.Add(8, );
                        //pointMap.Add(9, );
                        //pointMap.Add(10, );
                        //pointMap.Add(11, );
                        //pointMap.Add(12, );
                        //pointMap.Add(13, );
                        //pointMap.Add(14, );
                        //pointMap.Add(15, );
                        //pointMap.Add(16, );
                        //pointMap.Add(17, );
                        //pointMap.Add(18, );
                        //pointMap.Add(19, );
                        //pointMap.Add(20, );
                        //pointMap.Add(21, );
                        //pointMap.Add(22, );
                        //pointMap.Add(23, );
                        //pointMap.Add(24, );
                        //pointMap.Add(25, );
                        //pointMap.Add(26, );
                        //pointMap.Add(27, 27);

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
                        mappedEdges.Add(SectionFaceType.Bottom_Flange_Right, SectionFaceType.Bottom_Flange_Left);
                        mappedEdges.Add(SectionFaceType.Bottom_Flange_Right_Top, SectionFaceType.Bottom_Flange_Left_Top);
                        
                        //pointMap.Add(1, 1);
                        //pointMap.Add(2, 24);
                        //pointMap.Add(3, 23);
                        //pointMap.Add(4, 22);
                        //pointMap.Add(5, 21);
                        //pointMap.Add(6, 20);
                        //pointMap.Add(7, 19);
                        //pointMap.Add(8, 18);
                        //pointMap.Add(9, 17);
                        //pointMap.Add(10, 16);
                        //pointMap.Add(11, 15);
                        //pointMap.Add(12, 14);
                        //pointMap.Add(13, 13);
                        //pointMap.Add(14, 12);
                        //pointMap.Add(15, 11);
                        //pointMap.Add(16, 10);
                        //pointMap.Add(17, 9);
                        //pointMap.Add(18, 8);
                        //pointMap.Add(19, 7);
                        //pointMap.Add(20, 6);
                        //pointMap.Add(21, 5);
                        //pointMap.Add(22, 4);
                        //pointMap.Add(23, 3);
                        //pointMap.Add(24, 2);
                        //pointMap.Add(25, 26);
                        //pointMap.Add(26, 25);
                        //pointMap.Add(27, 27);

                        break;
                   
                    case 2:
                        mappedEdges.Add(SectionFaceType.Web_Right, SectionFaceType.Bottom);
                        mappedEdges.Add(SectionFaceType.Right_Web_Top, SectionFaceType.Bottom_Flange_Left);
                        mappedEdges.Add(SectionFaceType.Inner_Web_Right_Top, SectionFaceType.Bottom_Flange_Left_Top);
                        mappedEdges.Add(SectionFaceType.Top, SectionFaceType.Web_Left);
                        mappedEdges.Add(SectionFaceType.Web_Left, SectionFaceType.Top);
                        mappedEdges.Add(SectionFaceType.Bottom, SectionFaceType.Web_Right);
                        mappedEdges.Add(SectionFaceType.Inner_Web_Right_Bottom, SectionFaceType.Bottom_Flange_Right_Top);
                        mappedEdges.Add(SectionFaceType.Right_Web_Bottom, SectionFaceType.Bottom_Flange_Right);
                        
                        //pointMap.Add(1, );
                        //pointMap.Add(2, );
                        //pointMap.Add(3, );
                        //pointMap.Add(4, );
                        //pointMap.Add(5, );
                        //pointMap.Add(6, );
                        //pointMap.Add(7, );
                        //pointMap.Add(8, );
                        //pointMap.Add(9, );
                        //pointMap.Add(10, );
                        //pointMap.Add(11, );
                        //pointMap.Add(12, );
                        //pointMap.Add(13, );
                        //pointMap.Add(14, );
                        //pointMap.Add(15, );
                        //pointMap.Add(16, );
                        //pointMap.Add(17, );
                        //pointMap.Add(18, );
                        //pointMap.Add(19, );
                        //pointMap.Add(20, );
                        //pointMap.Add(21, );
                        //pointMap.Add(22, );
                        //pointMap.Add(23, );
                        //pointMap.Add(24, );
                        //pointMap.Add(25, );
                        //pointMap.Add(26, );
                        //pointMap.Add(27, 27);

                        break;
                    
                    case 3:
                        mappedEdges.Add(SectionFaceType.Web_Right, SectionFaceType.Web_Right);
                        mappedEdges.Add(SectionFaceType.Top_Flange_Right_Bottom, SectionFaceType.Bottom_Flange_Right_Top);
                        mappedEdges.Add(SectionFaceType.Top_Flange_Right, SectionFaceType.Bottom_Flange_Right);
                        mappedEdges.Add(SectionFaceType.Top, SectionFaceType.Bottom);
                        mappedEdges.Add(SectionFaceType.Top_Flange_Left, SectionFaceType.Bottom_Flange_Left);
                        mappedEdges.Add(SectionFaceType.Top_Flange_Left_Bottom, SectionFaceType.Bottom_Flange_Left_Top);
                        mappedEdges.Add(SectionFaceType.Web_Left, SectionFaceType.Web_Left);
                        mappedEdges.Add(SectionFaceType.Bottom, SectionFaceType.Top);
                        
                        //pointMap.Add(1, );
                        //pointMap.Add(2, );
                        //pointMap.Add(3, );
                        //pointMap.Add(4, );
                        //pointMap.Add(5, );
                        //pointMap.Add(6, );
                        //pointMap.Add(7, );
                        //pointMap.Add(8, );
                        //pointMap.Add(9, );
                        //pointMap.Add(10, );
                        //pointMap.Add(11, );
                        //pointMap.Add(12, );
                        //pointMap.Add(13, );
                        //pointMap.Add(14, );
                        //pointMap.Add(15, );
                        //pointMap.Add(16, );
                        //pointMap.Add(17, );
                        //pointMap.Add(18, );
                        //pointMap.Add(19, );
                        //pointMap.Add(20, );
                        //pointMap.Add(21, );
                        //pointMap.Add(22, );
                        //pointMap.Add(23, );
                        //pointMap.Add(24, );
                        //pointMap.Add(25, );
                        //pointMap.Add(26, );
                        //pointMap.Add(27, 27);

                        break;
                    
                    case 4:
                        mappedEdges.Add(SectionFaceType.Web_Right, SectionFaceType.Bottom_Flange_Right_Top);
                        mappedEdges.Add(SectionFaceType.Top, SectionFaceType.Bottom_Flange_Right);
                        mappedEdges.Add(SectionFaceType.Web_Left, SectionFaceType.Bottom);
                        mappedEdges.Add(SectionFaceType.Bottom, SectionFaceType.Bottom_Flange_Left);
                        mappedEdges.Add(SectionFaceType.Web_Right_Bottom, SectionFaceType.Bottom_Flange_Left_Top);
                        mappedEdges.Add(SectionFaceType.Bottom_Flange_Right_Bottom, SectionFaceType.Web_Left);
                        mappedEdges.Add(SectionFaceType.Bottom_Flange_Right, SectionFaceType.Top);
                        mappedEdges.Add(SectionFaceType.Bottom_Flange_Right_Top, SectionFaceType.Web_Right);
                        
                        //pointMap.Add(1, );
                        //pointMap.Add(2, );
                        //pointMap.Add(3, );
                        //pointMap.Add(4, );
                        //pointMap.Add(5, );
                        //pointMap.Add(6, );
                        //pointMap.Add(7, );
                        //pointMap.Add(8, );
                        //pointMap.Add(9, );
                        //pointMap.Add(10, );
                        //pointMap.Add(11, );
                        //pointMap.Add(12, );
                        //pointMap.Add(13, );
                        //pointMap.Add(14, );
                        //pointMap.Add(15, );
                        //pointMap.Add(16, );
                        //pointMap.Add(17, );
                        //pointMap.Add(18, );
                        //pointMap.Add(19, );
                        //pointMap.Add(20, );
                        //pointMap.Add(21, );
                        //pointMap.Add(22, );
                        //pointMap.Add(23, );
                        //pointMap.Add(24, );
                        //pointMap.Add(25, );
                        //pointMap.Add(26, );
                        //pointMap.Add(27, 27);

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

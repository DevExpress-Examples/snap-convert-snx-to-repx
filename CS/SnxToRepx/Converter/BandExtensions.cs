using DevExpress.XtraReports.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SnxToRepx.Converter
{
    /// <summary>
    /// This class contains useful extensions for the BandCollection class
    /// </summary>
    static class BandExtensions
    {

        /// <summary>
        /// Creates a new band of the specified kind and adds it to the collection
        /// </summary>
        /// <param name="col">A BandCollection instance that should hold a new Band</param>
        /// <param name="kind">A band kind to create a new Band</param>
        /// <returns>The Band of the specified kind</returns>
        public static Band Create(this BandCollection col, BandKind kind)
        {
            Band band = null;
            //Check the BandKind value and create a new band
            switch (kind)
            {
                case BandKind.ReportHeader: 
                    band = new ReportHeaderBand();
                    break;
                case BandKind.Detail:
                    //Each report can contain only a single DetailBand.
                    //If it does not yet exist, create a new band
                    if (col[kind] == null)
                        band = new DetailBand();
                    else
                        //Otherwise, return the existing DetailBand
                        band = col[kind];
                    break;
                case BandKind.GroupHeader:
                    band = new GroupHeaderBand();
                    break;
                case BandKind.GroupFooter:
                    band = new GroupFooterBand();
                    break;
                case BandKind.ReportFooter:
                    band = new ReportFooterBand();
                    break;
                case BandKind.TopMargin:
                    band = new TopMarginBand();
                    break;
                case BandKind.BottomMargin:
                    band = new BottomMarginBand();
                    break;
                case BandKind.PageHeader:
                    band = new PageHeaderBand();
                    break;
                case BandKind.PageFooter:
                    band = new PageFooterBand();
                    break;
                default:
                    break;
            }

            //If a band was successfully created and it is not yet added to the collection,
            if (band != null && !col.Contains(band))
            {
                //add this band to the collection and set its height to 0
                //because it will automatically resize itself based on content
                col.Add(band);
                band.HeightF = 0;
            }
            //Return the created band
            return band;
        }
    }
}

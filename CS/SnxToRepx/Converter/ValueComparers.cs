using DevExpress.XtraRichEdit.API.Layout;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SnxToRepx.Converter
{
    /// <summary>
    /// This is a comparer to sort a collection of RangedLayoutElements
    /// </summary>
    class RangedElementComparer : IComparer<RangedLayoutElement>
    {
        /// <summary>
        /// Compares two RangedLayoutElements by their ranges
        /// </summary>
        /// <param name="x">The first element</param>
        /// <param name="y">The second element</param>
        /// <returns>The result of comparison</returns>
        public int Compare(RangedLayoutElement x, RangedLayoutElement y)
        {
            //Use the default integer comparer to compare start positions for elements
            return Comparer<int>.Default.Compare(x.Range.Start, y.Range.Start);
        }
    }
}

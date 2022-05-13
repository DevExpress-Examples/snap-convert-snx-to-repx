using DevExpress.Snap.Core.API;

namespace SnxToRepx.Converter
{
    /// <summary>
    /// This class generates an extended markup, which can be used in ExpressionBindings
    /// </summary>
    class ExpressionMarkupGenerator : MarkupGenerator
    {
        public ExpressionMarkupGenerator(ReportGenerator generator, SnapDocument document) : base(generator, document) { }

        /// <summary>
        /// Processes a field and returns a corresponding Expression part
        /// </summary>
        /// <param name="entity">A Snap entity to process</param>
        protected override void ProcessField(SnapSingleListItemEntity entity)
        {
            //IMPORTANT! This approach is required to properly process summaries. Background: the converter uses HTML Markup
            //available in reports to format content. Standard fields are passed as mail-merge fields. Since it is impossible to
            //specify summaries in the standard mail-merge report, we need to construct an Expression instead of standard markup.
            //It uses a slightly different mechanism. Specifically, static string content should be wrapped with '. Dynamic content (fields)
            //should be outside static strings, and we need to use + to concatenate different parts. For example:
            //Static markup: <p align=center>[UnitPrice]</p>
            //Expression: '<p align=center>' + [UnitPrice] + '</p>'

            //Add the closing apostrophe and the plus sign (see the example above)
            Buffer.Append("' + ");
            //Check if it is a SnapText instance. SnapText contains format strings, summary settings, etc.
            SnapText snapText = entity as SnapText;
            if (snapText == null)
                //If not, generate the standard field representation ( [fieldName] )
                base.ProcessField(entity);
            else
            {
                //If so, get SnapText related properties
                string dataFieldName = snapText.DataFieldName;
                string formatString = "{0:" + snapText.FormatString + "}";
                bool isParameter = snapText.IsParameter;

                //if it is a parameter, add the ? sign (used in XtraReports to determine whether a control is bound to a parameter)
                if (isParameter)
                    dataFieldName = string.Format("?{0}", dataFieldName);                    
                else
                {
                    //If not, check if it contains summary information
                    string summaryFunc = SummaryConverter.Convert(snapText.SummaryFunc);
                    if (summaryFunc != string.Empty)
                        //Add a corresponding expression function (sumSum, sumCount, etc.)
                        dataFieldName = string.Format("{0}([{1}])", summaryFunc, dataFieldName);
                    else
                        //If no summary added, use the standard field format
                        dataFieldName = string.Format("[{0}]", dataFieldName);
                }

                //Check if the SnapText has a format string
                if (string.IsNullOrEmpty(snapText.FormatString))
                    Buffer.Append(dataFieldName);
                else
                    //Use the FormatString expression function to format the resulting value
                    Buffer.AppendFormat("FormatString('{0}', {1})", formatString, dataFieldName);
            }

            //Add an opening apostrophe
            Buffer.Append(" + '");
        }

        //Some tags in markup use ' for strings (e.g., font name). To properly use it in expressions,
        //we need to return a double apostrophe.
        public override string FontValueWrapperSymbol => "''";
        //Returns the expression with markup
        public override string Text => string.Format("'{0}'" , base.Text);
    }
}

namespace WpfSpreadsheetEditorDemo
{
    /// <summary>
    /// Represents the chart type item.
    /// </summary>
    public class ChartTypeItem
    {

        /// <summary>
        /// Initializes a new instance of the <see cref="ChartTypeItem"/> class.
        /// </summary>
        /// <param name="value">The name of chart type.</param>
        public ChartTypeItem(string value)
        {
            _value = value;
        }



        string _value;
        /// <summary>
        /// Gets the name of chart type.
        /// </summary>
        public string Value
        {
            get
            {
                return _value;
            }
        }



        /// <summary>
        /// Returns a string that represents description of chart type item.
        /// </summary>
        public override string ToString()
        {
            switch (Value)
            {
                case "Column":
                    return "Column";

                case "Line":
                    return "Line";

                case "Pie":
                    return "Pie";

                case "Bar":
                    return "Bar";

                case "Area":
                    return "Area";

                default:
                    return Value;
            }
        }

    }
}

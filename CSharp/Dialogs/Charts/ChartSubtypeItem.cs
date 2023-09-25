namespace WpfSpreadsheetEditorDemo
{
    /// <summary>
    /// Represents the chart subtype item.
    /// </summary>
    public class ChartSubtypeItem
    {

        /// <summary>
        /// Initializes a new instance of the <see cref="ChartSubtypeItem"/> class.
        /// </summary>
        /// <param name="value">The name of chart subtype.</param>
        public ChartSubtypeItem(string value)
        {
            _value = value;
        }



        string _value;
        /// <summary>
        /// Gets the name of chart subtype.
        /// </summary>
        public string Value
        {
            get
            {
                return _value;
            }
        }



        /// <summary>
        /// Returns a string that represents description of chart subtype item.
        /// </summary>
        public override string ToString()
        {
            switch (Value)
            {
                case "Clustered_Column":
                    return "Clustered Column";

                case "Stacked_Column":
                    return "Stacked Column";

                case "100%_Stacked_Column":
                    return "100% Stacked Column";

                case "3D_Clustered_Column":
                    return "3D Clustered Column";

                case "3D_Stacked_Column":
                    return "3D Stacked Column";

                case "3D_100%_Stacked_Column":
                    return "3D 100% Stacked Column";

                case "Line":
                    return "Line";

                case "Stacked_Line":
                    return "Stacked Line";

                case "100%_Stacked_Line":
                    return "100% Stacked Line";

                case "Line_With_Markers":
                    return "Line With Markers";

                case "Stacked_Line_With_Markers":
                    return "Stacked Line With Markers";

                case "100%_Stacked_Line_With_Mar":
                    return "100% Stacked Line With Markers";

                case "Curved_Line":
                    return "Curved Line";

                case "Stacked_Curved_Line":
                    return "Stacked Curved Line";

                case "100%_Stacked_Curved_Line":
                    return "100% Stacked Curved Line";

                case "Curved_Line_With_Markers":
                    return "Curved Line With Markers";

                case "Stacked_CurvLine_With_Mark":
                    return "Stacked Curved Line With Markers";

                case "100%_Stacked_CurvLineMark":
                    return "100% Stacked Curved Line With Markers";

                case "Pie":
                    return "Pie";

                case "Pie_Explosion":
                    return "Pie (Explosion)";

                case "Doughnut":
                    return "Doughnut";

                case "Clustered_Bar":
                    return "Clustered Bar";

                case "Stacked_Bar":
                    return "Stacked Bar";

                case "100%_Stacked_Bar":
                    return "100% Stacked Bar";

                case "Area":
                    return "Area";

                case "Stacked_Area":
                    return "Stacked Area";

                case "100%_Stacked_Area":
                    return "100% Stacked Area";

                default:
                    return Value;
            }
        }

    }
}

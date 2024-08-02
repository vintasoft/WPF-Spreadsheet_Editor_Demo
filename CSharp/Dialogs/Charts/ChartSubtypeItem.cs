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

                case "3D_Column":
                    return "3-D Column";

                case "3D_Clustered_Column":
                    return "3-D Clustered Column";

                case "3D_Stacked_Column":
                    return "3-D Stacked Column";

                case "3D_100%_Stacked_Column":
                    return "3-D 100% Stacked Column";

                case "Line":
                    return "Line";

                case "Stacked_Line":
                    return "Stacked Line";

                case "100%_Stacked_Line":
                    return "100% Stacked Line";

                case "Line_With_Markers":
                    return "Line with Markers";

                case "Stacked_Line_With_Markers":
                    return "Stacked Line with Markers";

                case "100%_Stacked_Line_With_Mar":
                    return "100% Stacked Line with Markers";

                case "Curved_Line":
                    return "Curved Line";

                case "Stacked_Curved_Line":
                    return "Stacked Curved Line";

                case "100%_Stacked_Curved_Line":
                    return "100% Stacked Curved Line";

                case "Curved_Line_With_Markers":
                    return "Curved Line with Markers";

                case "Stacked_CurvLine_With_Mark":
                    return "Stacked Curved Line with Markers";

                case "100%_Stacked_CurvLineMark":
                    return "100% Stacked Curved Line with Markers";

                case "3D_Line":
                    return "3-D Line";

                case "Pie":
                    return "Pie";

                case "Pie_Explosion":
                    return "Pie (Explosion)";

                case "3D_Pie":
                    return "3-D Pie";

                case "3D_Pie_Explosion":
                    return "3-D Pie (Explosion)";

                case "Doughnut":
                    return "Doughnut";

                case "Clustered_Bar":
                    return "Clustered Bar";

                case "Stacked_Bar":
                    return "Stacked Bar";

                case "100%_Stacked_Bar":
                    return "100% Stacked Bar";

                case "3D_Clustered_Bar":
                    return "3-D Clustered Bar";

                case "3D_Stacked_Bar":
                    return "Stacked Bar";

                case "3D_100%_Stacked_Bar":
                    return "3-D 100% Stacked Bar";

                case "Area":
                    return "Area";

                case "Stacked_Area":
                    return "Stacked Area";

                case "100%_Stacked_Area":
                    return "100% Stacked Area";

                case "3D_Area":
                    return "3-D Area";

                case "3D_Stacked_Area":
                    return "Stacked Area";

                case "3D_100%_Stacked_Area":
                    return "3-D 100% Stacked Area";

                case "High_Low_Close":
                    return "High-Low-Close";

                case "Open_High_Low_Close":
                    return "Open-High-Low-Close";

                case "Scatter":
                    return "Scatter";

                case "Scatter_SmoothAndMarker":
                    return "Scatter with Smooth Lines and Markers";

                case "Scatter_Smooth_Lines":
                    return "Scatter with Smooth Lines";

                case "Scatter_StraightAndMark":
                    return "Scatter with Straight Lines and Markers";

                case "Scatter_Straight_Lines":
                    return "Scatter with Straight Lines";

                case "Bubble":
                    return "Bubble";

                case "3D_Bubble":
                    return "3-D Bubble";

                case "Radar":
                    return "Radar";

                case "Radar_with_Markers":
                    return "Radar with Markers";

                case "Filled_Radar":
                    return "Filled Radar";

                case "3D_Surface":
                    return "3-D Surface";

                case "3D_Surface_Wireframe":
                    return "Wireframe 3-D Surface";

                default:
                    return Value;
            }
        }

    }
}

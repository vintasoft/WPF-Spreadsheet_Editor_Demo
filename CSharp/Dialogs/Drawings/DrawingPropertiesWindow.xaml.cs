using System;
using System.Windows;

using Vintasoft.Imaging.Office.Spreadsheet.Document;
using Vintasoft.Imaging.Office.Spreadsheet.Document.Editors;
using Vintasoft.Imaging.Office.Spreadsheet.UI;

using WpfDemosCommonCode;

namespace WpfSpreadsheetEditorDemo
{
    /// <summary>
    /// A dialog that allows to edit the drawing properties.
    /// </summary>
    public partial class DrawingPropertiesWindow : Window
    {

        #region Fields

        /// <summary>
        /// The visual editor.
        /// </summary>
        SpreadsheetVisualEditor _visualEditor;

        /// <summary>
        /// The drawing.
        /// </summary>
        SheetDrawing _drawing;

        /// <summary>
        /// The chart series, which are selected in combobox.
        /// </summary>
        ChartDataSeries _selectedSeries;

        /// <summary>
        /// A value indicating whether chart properties are being initialized.
        /// </summary>
        bool _isChartPropertiesInitializing;

        /// <summary>
        /// A value indicating whether series properties are being initialized.
        /// </summary>
        bool _isSeriesPropertiesInitializing;

        #endregion



        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="DrawingPropertiesWindow"/> class.
        /// </summary>
        public DrawingPropertiesWindow()
        {
            InitializeComponent();

            // done separately to not trigger event during initialization
            markerSizeNumericUpDown.ValueChanged += markerProperties_Changed;
        } 

        /// <summary>
        /// Initializes a new instance of the <see cref="DrawingPropertiesWindow"/> class.
        /// </summary>
        /// <param name="visualEditor">The visual editor.</param>
        /// <param name="drawing">The drawing.</param>
        public DrawingPropertiesWindow(SpreadsheetVisualEditor visualEditor, SheetDrawing drawing)
            : this()
        {
            _visualEditor = visualEditor;
            _drawing = drawing;

            // drawing name
            nameTextBox.Text = drawing.Name;
            // drawing description
            descriptionTextBox.Text = drawing.Description;
            // drawing rotation
            rotationAngleNumericUpDown.Value = drawing.Rotation;
            rotationAngleNumericUpDown.IsEnabled = drawing.Type != DrawingType.Chart;

            // drawing location
            sheetDrawingLocationEditorControl.Worksheet = visualEditor.FocusedWorksheet;
            sheetDrawingLocationEditorControl.SheetDrawingLocation = drawing.Location;

            // chart
            if (drawing.ChartProperties == null)
            {
                drawingPropertiesTabControl.Items.Remove(chartTabPage);
            }
            else
            {
                _isChartPropertiesInitializing = true;

                // chart type
                chartTabPage.Header = string.Format("Chart ({0})", drawing.ChartProperties.ChartType);

                // cancel button is disabled because chart changes applies immediately
                cancelButton.IsEnabled = false;

                // init marker style comboBox
                foreach (ChartMarkerStyle markerStyle in Enum.GetValues(typeof(ChartMarkerStyle)))
                    markerTypeComboBox.Items.Add(markerStyle);

                // chart title
                titleTextBox.Text = drawing.ChartProperties.Title;

                // data range
                try
                {
                    dataRangeTextBox.Text = drawing.ChartProperties.GetCellReferencesSet().GetBounds().ToString();
                }
                catch
                {
                    dataRangeTextBox.Text = drawing.ChartProperties.GetCellReferencesSet().ToString(visualEditor.Document);
                }

                // init series comboBox
                if (drawing.ChartProperties.Series.Count > 0)
                {
                    for (int i = 0; i < drawing.ChartProperties.Series.Count; i++)
                        seriesComboBox.Items.Add(string.Format("Series {0}", i + 1));

                    seriesComboBox.SelectedIndex = 0;
                }

                // categories axis data range
                if (drawing.ChartProperties.CategoryAxis != null)
                    categoriesDataRangeTextBox.Text = drawing.ChartProperties.CategoryAxis.ToString(visualEditor.Document);
                else
                    categoriesDataRangeTextBox.Text = "";

                if (drawing.ChartProperties.ChartType == ChartType.Line ||
                    drawing.ChartProperties.ChartType == ChartType.Bar2D ||
                    drawing.ChartProperties.ChartType == ChartType.Bar3D ||
                    drawing.ChartProperties.ChartType == ChartType.Axial ||
                    drawing.ChartProperties.ChartType == ChartType.Scatter)
                {
                    categoryAxisTitleTextBox.Text = drawing.ChartProperties.CategoryAxisTitle;
                    valuesAxisTitleTextBox.Text = drawing.ChartProperties.ValuesAxisTitle;
                }
                else
                {
                    categoryAxisTitleTextBox.IsEnabled = false;
                    valuesAxisTitleTextBox.IsEnabled = false;
                }

                if (drawing.ChartProperties.ChartType == ChartType.Line)
                    smoothLineCheckBox.IsEnabled = true;
                else
                    smoothLineCheckBox.IsEnabled = false;

                _isChartPropertiesInitializing = false;
            }
        }

        #endregion



        #region Methods

        /// <summary>
        /// When overridden in a derived class, is invoked whenever application code or internal 
        /// processes call <see cref="System.Windows.FrameworkElement.ApplyTemplate" />.
        /// </summary>
        public override void OnApplyTemplate()
        {
            base.OnApplyTemplate();

            descriptionTextBox.Focus();
        }

        #region UI

        /// <summary>
        /// Handles the SelectionChanged event of seriesComboBox object.
        /// </summary>
        private void seriesComboBox_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            _isSeriesPropertiesInitializing = true;

            // get selected series
            _selectedSeries = _drawing.ChartProperties.Series[seriesComboBox.SelectedIndex];

            // set the selected series properties to the UI
            if (_selectedSeries.Name != null)
                nameRangeTextBox.Text = _selectedSeries.Name.ToString(_visualEditor.Document);
            else
                nameRangeTextBox.Text = "";

            if (_selectedSeries.Values != null)
                valuesRangeTextBox.Text = _selectedSeries.Values.ToString(_visualEditor.Document);
            else
                valuesRangeTextBox.Text = "";

            // clear data points combobox
            dataPointComboBox.Items.Clear();

            // init the combobox with data points
            // add 'All' item to the combobox with data points
            dataPointComboBox.Items.Add("All");

            // if series contains data points
            if (_selectedSeries.DataPoints != null && _selectedSeries.DataPoints.Length > 0)
            {
                // add data point items
                for (int i = 0; i < _selectedSeries.DataPoints.Length; i++)
                    dataPointComboBox.Items.Add(string.Format("Data Point {0}", i + 1));
            }

            smoothLineCheckBox.IsChecked = _selectedSeries.SmoothLine;

            _isSeriesPropertiesInitializing = false;
            
            dataPointComboBox.SelectedIndex = 0;
        }

        /// <summary>
        /// Handles the SelectionChanged event of dataPointComboBox object.
        /// </summary>
        private void dataPointComboBox_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if (_isSeriesPropertiesInitializing)
                return;

            _isSeriesPropertiesInitializing = true;

            // if 'All' is selected
            if (dataPointComboBox.SelectedItem.ToString() == "All")
            {
                // set the series appearance properties
                dataPointAppearanceEditor.ShapeAppearance = _selectedSeries.AppearanceProperties;
                dataPointAppearanceGroupBox.Header = "Series appearance";

                // if series marker is set
                if (_selectedSeries.Marker != null)
                {
                    markerAppearanceGroupBox.IsEnabled = true;
                    // set marker properties to UI
                    markerAppearanceEditor.ShapeAppearance = _selectedSeries.Marker.AppearanceProperties;
                    markerTypeComboBox.SelectedItem = _selectedSeries.Marker.Style;
                    markerSizeNumericUpDown.Value = _selectedSeries.Marker.Size;
                }
                else
                {
                    markerAppearanceGroupBox.IsEnabled = false;
                }
            }
            else
            {
                // get the selected data point
                ChartDataPoint selectedDataPoint = _selectedSeries.DataPoints[dataPointComboBox.SelectedIndex - 1];

                // set the data point appearance properties
                dataPointAppearanceEditor.ShapeAppearance = selectedDataPoint.AppearanceProperties;
                dataPointAppearanceGroupBox.Header = "Data point appearance";

                // if data point marker is set
                if (selectedDataPoint.Marker != null)
                {
                    markerAppearanceGroupBox.IsEnabled = true;
                    // set marker properties to UI
                    markerAppearanceEditor.ShapeAppearance = selectedDataPoint.Marker.AppearanceProperties;
                    markerTypeComboBox.SelectedItem = selectedDataPoint.Marker.Style;
                    markerSizeNumericUpDown.Value = selectedDataPoint.Marker.Size;
                }
                else
                {
                    markerAppearanceGroupBox.IsEnabled = false;
                }
            }

            _isSeriesPropertiesInitializing = false;
        }

        /// <summary>
        /// Handles the ShapeAppearanceChanged event of dataPointAppearanceEditor object.
        /// </summary>
        private void dataPointAppearanceEditor_ShapeAppearanceChanged(object sender, EventArgs e)
        {
            if (_isSeriesPropertiesInitializing)
                return;

            // create worksheet editor
            WorksheetEditor worksheetEditor = _visualEditor.Editor.StartEditing(_visualEditor.FocusedWorksheet);
            try
            {
                // create the chart series editor
                SheetDrawingEditor drawingEditor = worksheetEditor.CreateDrawingEditor(_drawing);
                ChartPropertiesEditor chartPropertiesEditor = drawingEditor.CreateChartPropertiesEditor();
                ChartDataSeriesEditor chartSeriesEditor = chartPropertiesEditor.CreateChartDataSeriesEditor(_selectedSeries);

                // data points
                ChartDataPoint[] points = _selectedSeries.DataPoints;

                // if 'All' item is selected
                if (dataPointComboBox.SelectedItem.ToString() == "All")
                {
                    // set the appearance properties to the series
                    chartSeriesEditor.SetAppearanceProperties(dataPointAppearanceEditor.ShapeAppearance);

                    // set the appearance properties to all points
                    if (points != null && points.Length > 0)
                    {
                        for (int i = 0; i < points.Length; i++)
                            points[i] = new ChartDataPoint(points[i].Marker, dataPointAppearanceEditor.ShapeAppearance);
                    }
                }
                else
                {
                    // set the appearance properties to the point
                    // get selected data point
                    int selectedIndex = dataPointComboBox.SelectedIndex - 1;
                    ChartDataPoint selectedPoint = points[selectedIndex];
                    // create data point with new appearance
                    ChartDataPoint newPoint = new ChartDataPoint(selectedPoint.Marker, dataPointAppearanceEditor.ShapeAppearance);
                    // set new point
                    points[selectedIndex] = newPoint;
                }

                // set data points
                chartSeriesEditor.SetDataPoints(points);
            }
            finally
            {
                _visualEditor.Editor.FinishEditing();
            }
        }

        /// <summary>
        /// Handles the Changed event of markerProperties object.
        /// </summary>
        private void markerProperties_Changed(object sender, EventArgs e)
        {
            if (_isSeriesPropertiesInitializing)
                return;

            // create worksheet editor
            WorksheetEditor worksheetEditor = _visualEditor.Editor.StartEditing(_visualEditor.FocusedWorksheet);
            try
            {
                // create the chart series editor
                SheetDrawingEditor drawingEditor = worksheetEditor.CreateDrawingEditor(_drawing);
                ChartPropertiesEditor chartPropertiesEditor = drawingEditor.CreateChartPropertiesEditor();
                ChartDataSeriesEditor chartSeriesEditor = chartPropertiesEditor.CreateChartDataSeriesEditor(_selectedSeries);

                // create new marker
                ChartMarker marker = new ChartMarker(
                    (ChartMarkerStyle)markerTypeComboBox.SelectedItem,
                    (double)markerSizeNumericUpDown.Value,
                    markerAppearanceEditor.ShapeAppearance);

                // if 'All' item is selected
                if (dataPointComboBox.SelectedItem.ToString() == "All")
                {
                    // set marker properties to series
                    chartSeriesEditor.SetMarker(marker);
                }
                else
                {
                    // set marker properties to the point
                    // get selected data point
                    ChartDataPoint[] points = _selectedSeries.DataPoints;
                    int selectedIndex = dataPointComboBox.SelectedIndex - 1;
                    ChartDataPoint selectedPoint = points[selectedIndex];
                    // create data point with new marker
                    ChartDataPoint newPoint = new ChartDataPoint(marker, selectedPoint.AppearanceProperties);
                    // set new point
                    points[selectedIndex] = newPoint;
                    // set data points
                    chartSeriesEditor.SetDataPoints(points);
                }
            }
            finally
            {
                _visualEditor.Editor.FinishEditing();
            }
        }

        /// <summary>
        /// Handles the CheckedChanged event of smoothLineCheckBox object.
        /// </summary>
        private void smoothLineCheckBox_CheckedChanged(object sender, RoutedEventArgs e)
        {
            if (_isChartPropertiesInitializing)
                return;

            // create worksheet editor
            WorksheetEditor worksheetEditor = _visualEditor.Editor.StartEditing(_visualEditor.FocusedWorksheet);
            try
            {
                // create drawing editor
                ChartDataSeriesEditor seriesEditor = worksheetEditor.CreateDrawingEditor(_drawing).CreateChartPropertiesEditor().CreateChartDataSeriesEditor(_selectedSeries);
                // set chart title
                seriesEditor.SetSmoothLine(smoothLineCheckBox.IsChecked.Value == true);
            }
            finally
            {
                _visualEditor.Editor.FinishEditing();
            }
        }

        /// <summary>
        /// Handles the Click event of okButton object.
        /// </summary>
        private void okButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                CheckDrawingName();

                // create worksheet editor
                WorksheetEditor worksheetEditor = _visualEditor.Editor.StartEditing(_visualEditor.FocusedWorksheet);
                try
                {
                    // create drawing editor
                    SheetDrawingEditor drawingEditor = worksheetEditor.CreateDrawingEditor(_drawing);

                    // set drawing name
                    drawingEditor.SetName(nameTextBox.Text.Trim());

                    // set drawing description
                    drawingEditor.SetDescription(descriptionTextBox.Text);

                    // set drawing description
                    if (rotationAngleNumericUpDown.IsEnabled)
                        drawingEditor.SetRotation(rotationAngleNumericUpDown.Value);

                    // if drawing location is changed
                    if (!Equals(_drawing.Location, sheetDrawingLocationEditorControl.SheetDrawingLocation))
                        // set the drawing location
                        drawingEditor.SetLocation(sheetDrawingLocationEditorControl.SheetDrawingLocation);
                }
                finally
                {
                    _visualEditor.Editor.FinishEditing();
                }

                if (_drawing.Type == DrawingType.Chart)
                {
                    _visualEditor.StartEditing("Set chart properties");
                    try
                    {
                        // set chart title
                        if (IsTitleChanged(_visualEditor.ChartTitle, titleTextBox.Text))
                            _visualEditor.ChartTitle = titleTextBox.Text;
                        // set category axis title
                        if (categoryAxisTitleTextBox.IsEnabled && IsTitleChanged(_visualEditor.ChartCategoryAxisTitle, categoryAxisTitleTextBox.Text))
                            _visualEditor.ChartCategoryAxisTitle = categoryAxisTitleTextBox.Text;
                        // set values axis title
                        if (valuesAxisTitleTextBox.IsEnabled && IsTitleChanged(_visualEditor.ChartValuesAxisTitle, valuesAxisTitleTextBox.Text))
                            _visualEditor.ChartValuesAxisTitle = valuesAxisTitleTextBox.Text;
                    }
                    finally
                    {
                        _visualEditor.FinishEditing();
                    }
                }
            }
            catch (Exception ex)
            {
                DemosTools.ShowWarningMessage("Spreadsheet Editor Demo", ex.Message);
                return;
            }

            DialogResult = true;
        }

        #endregion

        /// <summary>
        /// Checks the drawing name.
        /// </summary>
        private void CheckDrawingName()
        {
            // get the drawing name
            string name = nameTextBox.Text;

            if (name != null)
                name = name.Trim();

            // if drawing name is empty
            if (string.IsNullOrEmpty(name))
                throw new InvalidOperationException("The drawing name cannot be empty.");
        }

        /// <summary>
        /// Returns a value indicating whether title value is changed.
        /// </summary>
        /// <param name="oldValue">Old title value.</param>
        /// <param name="newValue">New title value.</param>
        /// <returns>A value indicating whether title value is changed.</returns>
        private bool IsTitleChanged(string oldValue, string newValue)
        {
            // for title, null and empty string are equal values
            if (string.IsNullOrEmpty(oldValue) && string.IsNullOrEmpty(newValue))
                return false;

            return oldValue != newValue;
        }

        #endregion

    }
}

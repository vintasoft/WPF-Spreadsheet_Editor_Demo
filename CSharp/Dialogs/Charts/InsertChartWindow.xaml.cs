using System;
using System.Collections.Generic;
using System.IO;
using System.Windows;

using WpfDemosCommonCode;

using Vintasoft.Imaging.Office.Spreadsheet;
using Vintasoft.Imaging.Office.Spreadsheet.UI;

namespace WpfSpreadsheetEditorDemo
{
    /// <summary>
    /// "Insert chart" window.
    /// </summary>
    public partial class InsertChartWindow : Window
    {

        #region Fields

        /// <summary>
        /// The chart preview manager.
        /// </summary>
        ChartPreviewManager _chartPreview;

        /// <summary>
        /// Destination visual editor.
        /// </summary>
        SpreadsheetVisualEditor _destVisualEditor;

        #endregion



        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="InsertChartWindow"/> class.
        /// </summary>
        public InsertChartWindow()
        {
            InitializeComponent();

            chartTypeListBox.IsEnabled = false;
            chartSubtypeListBox.IsEnabled = false;
            okButton.IsEnabled = false;
        }

        #endregion



        #region Methods

        /// <summary>
        /// Sets the chart data source.
        /// </summary>
        /// <param name="destVisualEditor">The dest visual editor.</param>
        /// <param name="chartSourceResourceName">Name of the chart source resource.</param>
        public void SetChartDataSource(SpreadsheetVisualEditor destVisualEditor, string chartSourceResourceName)
        {
            _destVisualEditor = destVisualEditor;

            _chartPreview = new ChartPreviewManager(destVisualEditor);
            _chartPreview.InitializationException += ChartPreview_InitializationException;
            _chartPreview.SelectedChartChanged += ChartPreview_SelectedChartChanged;            
            _chartPreview.InitializationFinished += ChartPreview_InitializationFinished;
            _chartPreview.ChartPreview = chartViewer.VisualEditor;

            using (Stream chartSourceStream = WpfDemosCommonCode.DemosResourcesManager.GetResourceAsStream(chartSourceResourceName))
            {
                byte[] data = new byte[(int)chartSourceStream.Length];
                chartSourceStream.Read(data, 0, data.Length);
                Stream stream = new MemoryStream();
                stream.Write(data, 0, data.Length);
                stream.Position = 0;
                chartViewer.VisualEditor.Editor = new SpreadsheetEditor(new SpreadsheetEditorSource(stream, true));
            }
        }


        /// <summary>
        /// Raises the <see cref="E:System.Windows.Forms.Form.Closed" /> event.
        /// </summary>
        /// <param name="e">The <see cref="T:System.EventArgs" /> that contains the event data.</param>
        protected override void OnClosed(EventArgs e)
        {
            SpreadsheetEditor editor = chartViewer.VisualEditor.Editor;
            chartViewer.VisualEditor.CloseDocument();
            editor.Dispose();

            base.OnClosed(e);
        }


        /// <summary>
        /// Handles the SelectedChartChanged event of ChartPreview object.
        /// </summary>
        private void ChartPreview_SelectedChartChanged(object sender, EventArgs e)
        {
            okButton.IsEnabled = _chartPreview.IsChartSelected;
        }

        /// <summary>
        /// Handles the InitializationException event of ChartPreview object.
        /// </summary>
        private void ChartPreview_InitializationException(object sender, Vintasoft.Imaging.ExceptionEventArgs e)
        {
            DemosTools.ShowErrorMessage(e.Exception);
        }

        /// <summary>
        /// Handles the InitializationFinished event of ChartPreview object.
        /// </summary>
        private void ChartPreview_InitializationFinished(object sender, EventArgs e)
        {
            if (_chartPreview.ChartTypes.Count > 0)
            {
                foreach (string chartType in _chartPreview.ChartTypes.Keys)
                    chartTypeListBox.Items.Add(new ChartTypeItem(chartType));
                chartTypeListBox.SelectedIndex = 0;
                chartTypeListBox.IsEnabled = true;
                chartSubtypeListBox.IsEnabled = true;
            }
        }

        /// <summary>
        /// Handles the SelectionChanged event of chartTypeListBox object.
        /// </summary>
        private void chartTypeListBox_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            chartSubtypeListBox.Items.Clear();
            ChartTypeItem typeItem = (ChartTypeItem)chartTypeListBox.SelectedItem;
            Dictionary<string, int> chartSubtypes = _chartPreview.ChartTypes[typeItem.Value];
            foreach (string chartSubtype in chartSubtypes.Keys)
                chartSubtypeListBox.Items.Add(new ChartSubtypeItem(chartSubtype));
            chartSubtypeListBox.SelectedIndex = 0;
        }

        /// <summary>
        /// Handles the SelectionChanged event of chartSubtypeListBox object.
        /// </summary>
        private void chartSubtypeListBox_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if (chartSubtypeListBox.Items.Count > 0)
            {
                ChartTypeItem typeItem = (ChartTypeItem)chartTypeListBox.SelectedItem;
                ChartSubtypeItem subtypeItem = (ChartSubtypeItem)chartSubtypeListBox.SelectedItem;
                int worksheetIndex = _chartPreview.ChartTypes[typeItem.Value][subtypeItem.Value];
                chartViewer.VisualEditor.FocusedWorksheetIndex = worksheetIndex;
            }
        }

        /// <summary>
        /// Handles the Click event of okButton object.
        /// </summary>
        private void okButton_Click(object sender, EventArgs e)
        {
            _chartPreview.AddSelectedChart(_destVisualEditor);
            DialogResult = true;
        }

        #endregion

    }
}

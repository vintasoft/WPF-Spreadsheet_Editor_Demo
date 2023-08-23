using System;
using System.Globalization;
using System.Windows;

using Vintasoft.Imaging.Office.Spreadsheet.Document;
using Vintasoft.Imaging.Office.Spreadsheet.UI;

using WpfDemosCommonCode;

namespace WpfSpreadsheetEditorDemo
{
    /// <summary>
    /// A window that allows to view and change the worksheet format properties.
    /// </summary>
    public partial class WorksheetFormatWindow : Window
    {

        #region Fields

        /// <summary>
        /// Spreadsheet visual editor.
        /// </summary>
        SpreadsheetVisualEditor _visualEditor;

        #endregion



        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="WorksheetFormatWindow"/> class.
        /// </summary>
        public WorksheetFormatWindow()
        {
            InitializeComponent();
        }

        #endregion



        #region Properties

        /// <summary>
        /// Gets the current culture.
        /// </summary>
        public CultureInfo Culture
        {
            get
            {
                if (_visualEditor != null)
                {
                    try
                    {
                        return CultureInfo.GetCultureInfo(_visualEditor.DocumentCulture);
                    }
                    catch
                    {
                    }
                }
                return CultureInfo.CurrentCulture;
            }
        }

        #endregion



        #region Methods

        /// <summary>
        /// Shows this form with current document information.
        /// </summary>
        /// <param name="visualEditor">Spreadsheet visual editor.</param>
        public static bool? ShowDialog(SpreadsheetVisualEditor visualEditor)
        {
            WorksheetFormatWindow window = new WorksheetFormatWindow();
            window.WindowStartupLocation = WindowStartupLocation.CenterOwner;
            window.Owner = Application.Current.MainWindow;
            window.SetVisualEditor(visualEditor);
            return window.ShowDialog();
        }

        /// <summary>
        /// Sets the document information properties to this form UI.
        /// </summary>
        /// <param name="visualEditor">Spreadsheet visual editor.</param>
        private void SetVisualEditor(SpreadsheetVisualEditor visualEditor)
        {
            _visualEditor = visualEditor;
            UpdateUI();
        }

        /// <summary>
        /// Updates the user interface of this form.
        /// </summary>
        private void UpdateUI()
        {
            WorksheetFormat format = _visualEditor.FocusedWorksheet.Format;

            rowHeightTextBox.Text = Math.Round(format.RowHeight, 2).ToString(Culture);
            columnWidthTextBox.Text = Math.Round(format.ColumnWidth, 2).ToString(Culture);
            rowAutoHeightCheckBox.IsChecked = format.AutoHeight;
        }

        /// <summary>
        /// Handles the Click event of ButtonOk object.
        /// </summary>
        private void buttonOk_Click(object sender, RoutedEventArgs e)
        {
            double rowHeight;
            if (!double.TryParse(rowHeightTextBox.Text, NumberStyles.Float, Culture, out rowHeight))
            {
                DemosTools.ShowWarningMessage("Spreadsheet Editor Demo", "Row height must be an integer or decimal number.");
                return;
            }

            double columnWidth;
            if (!double.TryParse(columnWidthTextBox.Text, NumberStyles.Float, Culture, out columnWidth))
            {
                DemosTools.ShowWarningMessage("Spreadsheet Editor Demo", "Column width must be an integer or decimal number.");
                return;
            }

            bool isAutoHeight = rowAutoHeightCheckBox.IsChecked.Value;

            bool isRowsHiddenByDefault = _visualEditor.FocusedWorksheet.Format.RowsHiddenByDefault;

            WorksheetFormat format = new WorksheetFormat(columnWidth, rowHeight, isAutoHeight, isRowsHiddenByDefault);

            // if format is changed, set it to worksheet
            if (!Equals(format, _visualEditor.FocusedWorksheet.Format))
                _visualEditor.SetWorksheetFormat(format);

            DialogResult = true;
        }

        #endregion

    }
}

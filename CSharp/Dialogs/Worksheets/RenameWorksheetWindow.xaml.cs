using System;
using System.Windows;

using Vintasoft.Imaging.Office.Spreadsheet.UI;

using WpfDemosCommonCode;

namespace WpfSpreadsheetEditorDemo
{
    /// <summary>
    /// A dialog that allows to rename the worksheet.
    /// </summary>
    public partial class RenameWorksheetWindow : Window
    {

        #region Fields

        /// <summary>
        /// The spreadsheet visual editor.
        /// </summary>
        SpreadsheetVisualEditor _spreadsheetVisualEditor;

        #endregion



        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="RenameWorksheetWindow"/> class.
        /// </summary>
        /// <param name="spreadsheetVisualEditor">The spreadsheet visual editor.</param>
        public RenameWorksheetWindow(SpreadsheetVisualEditor spreadsheetVisualEditor)
        {
            InitializeComponent();

            _spreadsheetVisualEditor = spreadsheetVisualEditor;

            worksheetNameTextBox.Text = spreadsheetVisualEditor.FocusedWorksheet.Name;
        } 

        #endregion



        #region Properties

        bool _isWorksheetNameChanged = false;
        /// <summary>
        /// Gets a value indicating whether worksheet name is changed.
        /// </summary>
        public bool IsWorksheetNameChanged
        {
            get
            {
                return _isWorksheetNameChanged;
            }
        }



        #endregion



        #region Methods

        /// <summary>
        /// Handles the Click event of okButton object.
        /// </summary>
        private void okButton_Click(object sender, RoutedEventArgs e)
        {
            string worksheetName = worksheetNameTextBox.Text;
            if (worksheetName.Length > 40)
            {
                DemosTools.ShowWarningMessage("Spreadsheet Editor Demo", "Worksheet name cannot contains more than 40 symbols.");
                return;
            }

            try
            {
                _spreadsheetVisualEditor.RenameWorksheet(worksheetName);
                _isWorksheetNameChanged = true;
            }
            catch (Exception ex)
            {
                DemosTools.ShowErrorMessage(ex);
                return;
            }

            DialogResult = true;
        }

        #endregion

    }
}

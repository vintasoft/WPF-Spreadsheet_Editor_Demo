using System;
using System.Windows;

using Vintasoft.Imaging.Office.Spreadsheet.Document;
using Vintasoft.Imaging.Office.Spreadsheet.UI;
using WpfDemosCommonCode;

namespace WpfSpreadsheetEditorDemo
{
    /// <summary>
    /// A dialog that allows to change position of worksheets in workbook.
    /// </summary>
    public partial class MoveWorksheetWindow : Window
    {

        #region Fields

        /// <summary>
        /// The spreadsheet visual editor.
        /// </summary>
        SpreadsheetVisualEditor _spreadsheetVisualEditor;

        #endregion



        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="MoveWorksheetWindow"/> class.
        /// </summary>
        /// <param name="spreadsheetVisualEditor">The spreadsheet visual editor.</param>
        public MoveWorksheetWindow(SpreadsheetVisualEditor spreadsheetVisualEditor)
        {
            InitializeComponent();

            _spreadsheetVisualEditor = spreadsheetVisualEditor;

            foreach (Worksheet worksheet in _spreadsheetVisualEditor.Document.Worksheets)
            {
                worksheetNameListBox.Items.Add(worksheet.Name);
            }
        } 

        #endregion



        #region Properties

        bool _isWorksheetOrderChanged = false;
        /// <summary>
        /// Gets a value indicating whether worksheet order is changed.
        /// </summary>
        public bool IsWorksheetOrderChanged
        {
            get
            {
                return _isWorksheetOrderChanged;
            }
        }

        #endregion



        #region Methods
        
        /// <summary>
        /// "Move up button is clicked.
        /// </summary>
        private void moveUpButton_Click(object sender, RoutedEventArgs e)
        {
            int selectedWorksheetIndex = worksheetNameListBox.SelectedIndex;
            if (selectedWorksheetIndex == -1 || selectedWorksheetIndex == 0)
                return;

            string selectedWorksheetName = (string)worksheetNameListBox.Items[selectedWorksheetIndex];
            worksheetNameListBox.Items.RemoveAt(selectedWorksheetIndex);
            worksheetNameListBox.Items.Insert(selectedWorksheetIndex - 1, selectedWorksheetName);
            worksheetNameListBox.SelectedIndex = selectedWorksheetIndex - 1;
        }
        
        /// <summary>
        /// "Move down" button is clicked.
        /// </summary>
        private void moveDownButton_Click(object sender, RoutedEventArgs e)
        {
            int selectedWorksheetIndex = worksheetNameListBox.SelectedIndex;
            if (selectedWorksheetIndex == -1 || selectedWorksheetIndex == (worksheetNameListBox.Items.Count - 1))
                return;

            string selectedWorksheetName = (string)worksheetNameListBox.Items[selectedWorksheetIndex];
            worksheetNameListBox.Items.RemoveAt(selectedWorksheetIndex);
            worksheetNameListBox.Items.Insert(selectedWorksheetIndex + 1, selectedWorksheetName);
            worksheetNameListBox.SelectedIndex = selectedWorksheetIndex + 1;
        }
        
        /// <summary>
        /// "OK" button is clicked.
        /// </summary>
        private void okButton_Click(object sender, RoutedEventArgs e)
        {
            int[] newIndexes = new int[worksheetNameListBox.Items.Count];

            // for each worksheet
            for (int newWorksheetIndex = 0; newWorksheetIndex < worksheetNameListBox.Items.Count; newWorksheetIndex++)
            {
                // get worksheet name
                string worksheetName = (string)worksheetNameListBox.Items[newWorksheetIndex];
                // get worksheet index
                newIndexes[newWorksheetIndex] = _spreadsheetVisualEditor.Document.GetWorksheetIndex(worksheetName);
            }

            try
            {
                _isWorksheetOrderChanged = _spreadsheetVisualEditor.ChangeWorksheetIndexes(newIndexes);
            }
            catch (Exception ex)
            {
                DemosTools.ShowErrorMessage(ex);
            }

            DialogResult = true;
        }

        #endregion

    }
}

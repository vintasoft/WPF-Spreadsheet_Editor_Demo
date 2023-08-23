using System.Windows;
using System.Windows.Controls;

using Vintasoft.Imaging.Office.Spreadsheet.UI;

namespace WpfSpreadsheetEditorDemo
{
    /// <summary>
    /// A dialog that allows to copy content of worksheet cells with special settings.
    /// </summary>
    public partial class CellPasteSpecialWindow : Window
    {

        /// <summary>
        /// The spreadsheet visual editor.
        /// </summary>
        SpreadsheetVisualEditor _spreadsheetVisualEditor;
        
        /// <summary>
        /// Initializes a new instance of the <see cref="CellPasteSpecialWindow"/> class.
        /// </summary>
        /// <param name="spreadsheetVisualEditor">The spreadsheet visual editor.</param>
        public CellPasteSpecialWindow(SpreadsheetVisualEditor spreadsheetVisualEditor)
        {
            InitializeComponent();

            _spreadsheetVisualEditor = spreadsheetVisualEditor;
        }
        
        /// <summary>
        /// Checkbox is checked.
        /// </summary>
        private void CheckBox_CheckedChanged(object sender, RoutedEventArgs e)
        {
            if (copyStylesCheckBox.IsChecked != true &&
                copyValuesCheckBox.IsChecked != true && 
                copyFormulasCheckBox.IsChecked != true &&
                copyCommentsCheckBox.IsChecked != true && 
                copyHyperlinksCheckBox.IsChecked != true)
            {
                MessageBox.Show("All checkboxes cannot be disabled.", "Error");

                CheckBox checkBox = (CheckBox)sender;
                checkBox.IsChecked = true;
            }
        }

        /// <summary>
        /// "OK" button is clicked.
        /// </summary>
        private void okButton_Click(object sender, RoutedEventArgs e)
        {
            Vintasoft.Imaging.Office.Spreadsheet.SheetCellsCopyMode sheetCellsCopyMode = Vintasoft.Imaging.Office.Spreadsheet.SheetCellsCopyMode.CopyAll;
            if (copyStylesCheckBox.IsChecked != true)
                sheetCellsCopyMode ^= Vintasoft.Imaging.Office.Spreadsheet.SheetCellsCopyMode.CopyCellStyle;
            if (copyValuesCheckBox.IsChecked != true)
                sheetCellsCopyMode ^= Vintasoft.Imaging.Office.Spreadsheet.SheetCellsCopyMode.CopyCellValue;
            if (copyFormulasCheckBox.IsChecked != true)
                sheetCellsCopyMode ^= Vintasoft.Imaging.Office.Spreadsheet.SheetCellsCopyMode.CopyCellFormula;
            if (copyCommentsCheckBox.IsChecked != true)
                sheetCellsCopyMode ^= Vintasoft.Imaging.Office.Spreadsheet.SheetCellsCopyMode.CopyCellComment;
            if (copyHyperlinksCheckBox.IsChecked != true)
                sheetCellsCopyMode ^= Vintasoft.Imaging.Office.Spreadsheet.SheetCellsCopyMode.CopyHyperlinks;

            _spreadsheetVisualEditor.PasteCells(sheetCellsCopyMode);

            DialogResult = true;
        }

    }
}

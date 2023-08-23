using System.Windows;

using Vintasoft.Imaging.Office.Spreadsheet.Document;
using Vintasoft.Imaging.Office.Spreadsheet.UI;

namespace WpfSpreadsheetEditorDemo
{
    /// <summary>
    /// A dialog that allows to view and manage the defined names.
    /// </summary>
    public partial class DefinedNameManagerWindow : Window
    {

        #region Fields

        /// <summary>
        /// The global scope.
        /// </summary>
        public const string GLOBAL_SCOPE = "Workbook";

        /// <summary>
        /// The spreadsheet visual editor.
        /// </summary>
        SpreadsheetVisualEditor _spreadsheetVisualEditor;

        #endregion



        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="DefinedNameManagerWindow"/> class.
        /// </summary>
        /// <param name="spreadsheetVisualEditor">The spreadsheet visual editor.</param>
        public DefinedNameManagerWindow(SpreadsheetVisualEditor spreadsheetVisualEditor)
        {
            InitializeComponent();

            _spreadsheetVisualEditor = spreadsheetVisualEditor;

            InitDefinedNameListView();
        }

        #endregion



        #region Methods

        /// <summary>
        /// "New..." button is clicked.
        /// </summary>
        private void newButton_Click(object sender, RoutedEventArgs e)
        {
            // create dialog that allows to add new defined name
            EditDefinedNameWindow dlg = new EditDefinedNameWindow(_spreadsheetVisualEditor);
            dlg.WindowStartupLocation = WindowStartupLocation.CenterOwner;
            dlg.Owner = this;
            // show the dialog
            if (dlg.ShowDialog() == true)
            {
                InitDefinedNameListView();
            }
        }

        /// <summary>
        /// "Edit..." button is clicked.
        /// </summary>
        private void editButton_Click(object sender, RoutedEventArgs e)
        {
            if (definedNameListView.SelectedItems.Count == 0)
                return;

            // get the defined name that should be edited
            DefinedName definedName = (DefinedName)definedNameListView.SelectedItems[0];

            // create dialog that allows to edit the defined name
            EditDefinedNameWindow dlg = new EditDefinedNameWindow(_spreadsheetVisualEditor, definedName);
            dlg.WindowStartupLocation = WindowStartupLocation.CenterOwner;
            dlg.Owner = this;
            // show the dialog
            if (dlg.ShowDialog() == true)
            {
                InitDefinedNameListView();
            }
        }

        /// <summary>
        /// "Delete" button is clicked.
        /// </summary>
        private void deleteButton_Click(object sender, RoutedEventArgs e)
        {
            if (definedNameListView.SelectedItems.Count == 0)
                return;

            // get the defined name that should be deleted
            DefinedName definedName = (DefinedName)definedNameListView.SelectedItems[0];
            // get the index of defined name in defined name list
            int definedNameIndex = _spreadsheetVisualEditor.Document.DefinedNames.IndexOf(definedName);
            // delete defined name by index
            _spreadsheetVisualEditor.Editor.RemoveDefinedNameAt(definedNameIndex);

            // delete item that represents defined name in the list view
            definedNameListView.Items.Remove(definedName);
        }

        /// <summary>
        /// Selected defined name is changed.
        /// </summary>
        private void definedNameListView_SelectedIndexChanged(object sender, System.EventArgs e)
        {
            bool isEditButtonEnabled = definedNameListView.SelectedItems.Count > 0;
            editButton.IsEnabled = isEditButtonEnabled;
            deleteButton.IsEnabled = isEditButtonEnabled;
        }

        /// <summary>
        /// Initializes the list view with defined names.
        /// </summary>
        private void InitDefinedNameListView()
        {
            definedNameListView.Items.Clear();
            foreach (DefinedName definedName in _spreadsheetVisualEditor.Document.DefinedNames)
                definedNameListView.Items.Add(definedName);
        }

        #endregion

    }
}

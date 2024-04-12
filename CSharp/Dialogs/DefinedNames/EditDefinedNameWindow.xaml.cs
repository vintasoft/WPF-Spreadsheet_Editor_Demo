using System;
using System.Windows;

using Vintasoft.Imaging.Office.Spreadsheet.Document;
using Vintasoft.Imaging.Office.Spreadsheet.UI;

using WpfDemosCommonCode;

namespace WpfSpreadsheetEditorDemo
{
    /// <summary>
    /// A dialog that allows to add or edit the defined name.
    /// </summary>
    public partial class EditDefinedNameWindow : Window
    {

        #region Fields

        /// <summary>
        /// Global scope name.
        /// </summary>
        const string GLOBAL_SCOPE = "Workbook";

        /// <summary>
        /// Spreadsheet visual editor.
        /// </summary>
        SpreadsheetVisualEditor _spreadsheetVisualEditor;

        /// <summary>
        /// Defined name to edit.
        /// </summary>
        DefinedName _definedName;

        #endregion



        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="EditDefinedNameWindow"/> class.
        /// </summary>
        public EditDefinedNameWindow()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="EditDefinedNameWindow"/> class.
        /// </summary>
        /// <param name="spreadsheetVisualEditor">The spreadsheet visual editor.</param>
        /// <param name="value">Value of new defined name.</param>
        public EditDefinedNameWindow(SpreadsheetVisualEditor spreadsheetVisualEditor, string value)
            : this()
        {
            _spreadsheetVisualEditor = spreadsheetVisualEditor;

            // initialize the scope combobox
            InitScopeComboBox();

            scopeComboBox.SelectedItem = GLOBAL_SCOPE;
            scopeComboBox.IsEnabled = true;

            if (string.IsNullOrEmpty(value))
                value = CreateRefersTo();
            if (!value.StartsWith("="))
                value = "=" + value;

            refersToTextBox.Text = value;

            this.Title = "Add Defined Name";
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="EditDefinedNameWindow"/> class.
        /// </summary>
        /// <param name="spreadsheetVisualEditor">The spreadsheet visual editor.</param>
        public EditDefinedNameWindow(SpreadsheetVisualEditor spreadsheetVisualEditor)
            : this(spreadsheetVisualEditor, string.Empty)
        {
        }


        /// <summary>
        /// Initializes a new instance of the <see cref="EditDefinedNameWindow"/> class.
        /// </summary>
        /// <param name="spreadsheetVisualEditor">The spreadsheet visual editor.</param>
        /// <param name="definedName">The defined name that should be edited.</param>
        public EditDefinedNameWindow(SpreadsheetVisualEditor spreadsheetVisualEditor, DefinedName definedName)
            : this()
        {
            if (spreadsheetVisualEditor == null)
                throw new ArgumentNullException("spreadsheetVisualEditor");
            if (definedName == null)
                throw new ArgumentNullException("definedName");

            _spreadsheetVisualEditor = spreadsheetVisualEditor;
            _definedName = definedName;

            // initialize the scope combobox
            InitScopeComboBox();

            if (string.IsNullOrEmpty(definedName.WorksheetName))
                scopeComboBox.SelectedItem = GLOBAL_SCOPE;
            else
                scopeComboBox.SelectedItem = definedName.WorksheetName;
            scopeComboBox.IsEnabled = false;

            nameTextBox.Text = definedName.Name;
            commentTextBox.Text = definedName.Comment;
            refersToTextBox.Text = "=" + spreadsheetVisualEditor.GetActualValueByDefinedName(definedName.Name);
        }

        #endregion



        #region Methods

        /// <summary>
        /// Handles the Click event of okButton object.
        /// </summary>
        private void okButton_Click(object sender, RoutedEventArgs e)
        {
            string name = nameTextBox.Text;
            if (string.IsNullOrEmpty(name))
            {
                DemosTools.ShowErrorMessage("Name is not defined.");
                nameTextBox.Focus();
                return;
            }

            string worksheetName = (string)scopeComboBox.SelectedItem;
            if (worksheetName == GLOBAL_SCOPE)
                worksheetName = null;

            try
            {
                // if existing defined name is editing
                if (_definedName != null)
                {
                    // set/change the defined name
                    _spreadsheetVisualEditor.SetDefinedName(_definedName, name, worksheetName, refersToTextBox.Text, commentTextBox.Text);
                }
                else
                {
                    // add the defined name
                    _spreadsheetVisualEditor.AddDefinedName(name, worksheetName, refersToTextBox.Text, commentTextBox.Text);
                }
            }
            catch (Exception ex)
            {
                DemosTools.ShowWarningMessage("Spreadsheet Editor Demo", ex.Message);
                return;
            }

            DialogResult = true;
        }

        /// <summary>
        /// Initializes combo box with scope (worksheet names).
        /// </summary>
        private void InitScopeComboBox()
        {
            // add "Workbook" (for ability to add global defined name) to the combobox
            scopeComboBox.Items.Add(GLOBAL_SCOPE);

            // for each worksheet
            foreach (Worksheet worksheet in _spreadsheetVisualEditor.Document.Worksheets)
            {
                // add the worksheet name to the combobox
                scopeComboBox.Items.Add(worksheet.Name);
            }
        }

        /// <summary>
        /// Returns a string that represents reference to the focused cell in focused worksheet.
        /// </summary>
        private string CreateRefersTo()
        {
            return _spreadsheetVisualEditor.GetFixedSelectedCells().ToString(_spreadsheetVisualEditor.Document.Defaults.FormattingProperties);
        }

        #endregion

    }
}

using System;
using System.Windows;

using Vintasoft.Imaging.Office.Spreadsheet.Document;
using Vintasoft.Imaging.Office.Spreadsheet.UI;

using WpfDemosCommonCode;

namespace WpfSpreadsheetEditorDemo
{
    /// <summary>
    /// A form that allows to add or edit the hyperlink.
    /// </summary>
    public partial class EditHyperlinkWindow : Window
    {

        #region Fields

        /// <summary>
        /// Spreadsheet visual editor.
        /// </summary>
        SpreadsheetVisualEditor _visualEditor;

        /// <summary>
        /// Indicates whether the dialog is opened in edit mode. True - "Edit hyperlink" menu is selected; False - "Add hyperlink" menu is selected.
        /// </summary>
        bool _isEditDialog;

        #endregion



        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="EditHyperlinkWindow"/> class.
        /// </summary>
        public EditHyperlinkWindow()
        {
            InitializeComponent();
        }

        #endregion



        #region Methods

        /// <summary>
        /// Shows this form with information about current document.
        /// </summary>
        /// <param name="visualEditor">Spreadsheet visual editor.</param>
        /// <param name="isEditDialog"><b>True</b> if existing hyperlink is being edited;
        /// <b>false</b> if new hyperlink is being added.</param>
        public static bool? ShowDialog(SpreadsheetVisualEditor visualEditor, bool isEditDialog)
        {
            EditHyperlinkWindow window = new EditHyperlinkWindow();
            window.WindowStartupLocation = WindowStartupLocation.CenterOwner;
            window.Owner = Application.Current.MainWindow;
            window.SetVisualEditor(visualEditor, isEditDialog);
            return window.ShowDialog();
        }


        /// <summary>
        /// Sets the document information properties to this form UI.
        /// </summary>
        /// <param name="visualEditor">Spreadsheet visual editor.</param>
        /// <param name="isEditDialog"><b>True</b> is existing hyperlink is being edited, 
        /// <b>false</b> if new hyperlink is being added.</param>
        private void SetVisualEditor(SpreadsheetVisualEditor visualEditor, bool isEditDialog)
        {
            _visualEditor = visualEditor;
            _isEditDialog = isEditDialog;
            UpdateUI();
        }

        /// <summary>
        /// Updates the user interface of this form.
        /// </summary>
        private void UpdateUI()
        {
            // add sheet names
            foreach (Worksheet worksheet in _visualEditor.Document.Worksheets)
                sheetComboBox.Items.Add(worksheet.Name);

            // add defined names
            foreach (DefinedName name in _visualEditor.Document.DefinedNames)
            {
                if (!string.IsNullOrEmpty(name.WorksheetName))
                    definedNamesListBox.Items.Add(string.Format("{0}!{1}", name.WorksheetName, name.Name));
                else
                    definedNamesListBox.Items.Add(name.Name);
            }

            // get hyperlink from focused cell
            Hyperlink cellHyperlink = _visualEditor.FocusedHyperlink;

            // if existing hyperlink is editing
            if (_isEditDialog)
            {
                if (cellHyperlink == null)
                    throw new InvalidOperationException();

                // if hyperlink contains cell reference
                if (cellHyperlink.Location != null)
                {
                    CellReferences reference = new CellReferences(cellHyperlink.Location.TopLeft, cellHyperlink.Location.BottomRight);
                    // set hyperlink location
                    cellReferenceTextBox.Text = reference.GetA1Name();

                    // if sheet exists
                    if (sheetComboBox.Items.Contains(cellHyperlink.Location.SheetName))
                        sheetComboBox.SelectedItem = cellHyperlink.Location.SheetName;
                    else
                        sheetComboBox.SelectedItem = _visualEditor.FocusedWorksheet.Name;

                    hyperlinkTabControl.SelectedItem = referenceTabPage;
                }
                // if hyperlink contains defined name
                else if (!string.IsNullOrEmpty(cellHyperlink.Name))
                {
                    // if defined name not found
                    if (!definedNamesListBox.Items.Contains(cellHyperlink.Name))
                        definedNamesListBox.Items.Insert(0, cellHyperlink.Name);

                    // select corresponding name
                    definedNamesListBox.SelectedItem = cellHyperlink.Name;

                    hyperlinkTabControl.SelectedItem = nameTabPage;
                }
                else
                {
                    // set hyperlink URI
                    addressTextBox.Text = cellHyperlink.Url;

                    hyperlinkTabControl.SelectedItem = addressTabPage;
                }
            }
            else
            {
                addressTextBox.Text = @"https://www.vintasoft.com";
                cellReferenceTextBox.Text = "A1";
                sheetComboBox.SelectedItem = _visualEditor.FocusedWorksheet.Name;
                this.Title = "Add Hyperlink";
            }
        }

        /// <summary>
        /// Handles the Click event of ButtonOk object.
        /// </summary>
        private void buttonOk_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // if "Existing file or web page" tab is selected
                if (hyperlinkTabControl.SelectedItem == addressTabPage)
                {
                    Exception ex = null;
                    if (TryCreateUri(addressTextBox.Text, out ex))
                    {
                        // create hyperlink to the URL
                        Hyperlink hyperlink = Hyperlink.CreateUrl(addressTextBox.Text);
                        // if existing hyperlink is editing
                        if (_isEditDialog)
                        {
                            // edit hyperlink of focused cell
                            _visualEditor.FocusedHyperlink = hyperlink;
                        }
                        // if new hyperlink is adding
                        else
                        {
                            _visualEditor.StartEditing("Add hyperlink");
                            try
                            {
                                // add hyperlink to the focused cell
                                _visualEditor.AddHyperlink(hyperlink);

                                if (_visualEditor.FocusedCell != null && string.IsNullOrEmpty(_visualEditor.FocusedCellValue))
                                    _visualEditor.FocusedCellValue = addressTextBox.Text;
                            }
                            finally
                            {
                                _visualEditor.FinishEditing();
                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show(ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                        return;
                    }
                }
                // if "Cell reference" tab is selected
                else if (hyperlinkTabControl.SelectedItem == referenceTabPage)
                {
                    try
                    {
                        // parse cell references
                        CellReferences cellReferences = CellReferences.Parse(cellReferenceTextBox.Text);

                        // create reference with sheet name
                        CellReferences fullReference = new CellReferences(sheetComboBox.SelectedItem.ToString(), cellReferences);
                        // create hyperlink
                        Hyperlink hyperlink = Hyperlink.CreateCellReferences(fullReference);
                        // if existing hyperlink is editing
                        if (_isEditDialog)
                        {
                            // edit hyperlink of focused cell
                            _visualEditor.FocusedHyperlink = hyperlink;
                        }
                        // if new hyperlink is adding
                        else
                        {
                            _visualEditor.StartEditing("Add hyperlink");
                            try
                            {
                                // add hyperlink to the focused cell
                                _visualEditor.AddHyperlink(hyperlink);

                                if (_visualEditor.FocusedCell != null && string.IsNullOrEmpty(_visualEditor.FocusedCellValue))
                                    _visualEditor.FocusedCellValue = fullReference.GetA1Name();
                            }
                            finally
                            {
                                _visualEditor.FinishEditing();
                            }
                        }
                    }
                    catch
                    {
                        DemosTools.ShowWarningMessage("Spreadsheet Editor Demo", "Cell reference is invalid.");
                        return;
                    }
                }
                // if "Defined names" tab is selected
                else
                {
                    // if defined name is selected
                    if (definedNamesListBox.SelectedItem != null)
                    {
                        // create hyperlink to the defined name
                        Hyperlink hyperlink = Hyperlink.CreateDefinedName(definedNamesListBox.SelectedItem.ToString());
                        // if existing hyperlink is editing
                        if (_isEditDialog)
                        {
                            // edit hyperlink of focused cell
                            _visualEditor.FocusedHyperlink = hyperlink;
                        }
                        // if new hyperlink is adding
                        else
                        {
                            _visualEditor.StartEditing("Add hyperlink");
                            try
                            {
                                // add hyperlink to the focused cell
                                _visualEditor.AddHyperlink(hyperlink);

                                if (_visualEditor.FocusedCell != null && string.IsNullOrEmpty(_visualEditor.FocusedCellValue))
                                    _visualEditor.FocusedCellValue = definedNamesListBox.SelectedItem.ToString();
                            }
                            finally
                            {
                                _visualEditor.FinishEditing();
                            }
                        }
                    }
                    else
                    {
                        DemosTools.ShowWarningMessage("Spreadsheet Editor Demo", "Defined name is not selected.");
                        return;
                    }
                }
            }
            catch (Exception ex)
            {
                DemosTools.ShowErrorMessage(ex);
            }
            DialogResult = true;
        }

        /// <summary>
        /// Tries to create an <see cref="Uri"/> instance from specified URI string.
        /// </summary>
        /// <param name="uriString">URI string.</param>
        /// <param name="exception">An exception that contains information about error.</param>
        /// <returns><b>True</b> if <see cref="Uri"/> is created successfully; otherwise, <b>false</b>.</returns>
        private bool TryCreateUri(string uriString, out Exception exception)
        {
            try
            {
                Uri uri = new Uri(uriString);
                exception = null;
                return true;
            }
            catch (UriFormatException ex)
            {
                exception = ex;
                return false;
            }
        }

        #endregion

    }
}

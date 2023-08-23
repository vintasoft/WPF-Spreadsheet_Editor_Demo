using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;

using Vintasoft.Imaging.Office.Spreadsheet;
using Vintasoft.Imaging.Office.Spreadsheet.Document;
using Vintasoft.Imaging.Office.Spreadsheet.Functions;

namespace WpfSpreadsheetEditorDemo
{
    /// <summary>
    /// A window that allows to view and select the function.
    /// </summary>
    public partial class SelectFunctionWindow : Window
    {

        #region Fields

        /// <summary>
        /// The document.
        /// </summary>
        SpreadsheetDocument _document;

        /// <summary>
        /// Dictionary that contains functions divided by categories: category name => function names.
        /// </summary>
        Dictionary<FunctionCategory, string[]> _categoryToFunctions;

        #endregion



        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="SelectFunctionWindow"/> class.
        /// </summary>
        public SelectFunctionWindow()
        {
            InitializeComponent();

            categoryComboBox.Items.Add(FunctionCategory.All);
            foreach (FunctionCategory category in SupportedFunctions.Categories)
                categoryComboBox.Items.Add(category);
            categoryComboBox.SelectedIndex = 0;
        }

        #endregion



        #region Methods

        /// <summary>
        /// Shows this dialog and returns selected function.
        /// </summary>
        /// <param name="document">The document.</param>
        public static string SelectFunction(SpreadsheetDocument document)
        {
            SelectFunctionWindow window = new SelectFunctionWindow();
            window._document = document;
            window.WindowStartupLocation = WindowStartupLocation.CenterOwner;
            window.Owner = Application.Current.MainWindow;
            if (window.ShowDialog() == true)
                return (string)window.functionsListBox.SelectedItem;
            return null;
        }

        /// <summary>
        /// When overridden in a derived class, is invoked whenever application code or
        /// internal processes call <see cref="System.Windows.FrameworkElement.ApplyTemplate" />.
        /// </summary>
        public override void OnApplyTemplate()
        {
            base.OnApplyTemplate();

            _categoryToFunctions = new Dictionary<FunctionCategory, string[]>();
            _categoryToFunctions.Add(FunctionCategory.All, _document.GetSupportedFunctions(FunctionCategory.All));
            foreach (FunctionCategory category in SupportedFunctions.Categories)
                _categoryToFunctions.Add(category, _document.GetSupportedFunctions(category));

            UpdateUI();
        }

        /// <summary>
        /// Updates the user interface of this form.
        /// </summary>
        private void UpdateUI()
        {
            if (_categoryToFunctions != null)
            {
                // clear the functions listbox
                functionsListBox.Items.Clear();

                // if search textbox has no text
                if (string.IsNullOrEmpty(searchTextBox.Text))
                {
                    // add all functions by selected category
                    foreach (string item in _categoryToFunctions[(FunctionCategory)categoryComboBox.SelectedItem])
                        functionsListBox.Items.Add(item);
                }
                else
                {
                    string text = searchTextBox.Text.ToUpperInvariant();

                    // add all functions which contains text from search textbox
                    foreach (string functionName in _categoryToFunctions[(FunctionCategory)categoryComboBox.SelectedItem])
                    {
                        if (functionName.ToUpperInvariant().Contains(text))
                            functionsListBox.Items.Add(functionName);
                    }
                }

                // if listbox has elements
                if (functionsListBox.Items.Count > 0)
                    functionsListBox.SelectedIndex = 0;
            }
        }

        /// <summary>
        /// Handles the TextChanged event of SearchTextBox object.
        /// </summary>
        private void searchTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            UpdateUI();
        }

        /// <summary>
        /// Handles the SelectionChanged event of CategoryComboBox object.
        /// </summary>
        private void categoryComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            UpdateUI();
        }

        /// <summary>
        /// Handles the Click event of ButtonOk object.
        /// </summary>
        private void buttonOk_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
        }

        #endregion

    }
}

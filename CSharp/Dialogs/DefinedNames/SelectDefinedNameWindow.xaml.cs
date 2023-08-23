using System.Windows;

using Vintasoft.Imaging.Office.Spreadsheet.Document;

namespace WpfSpreadsheetEditorDemo
{
    /// <summary>
    /// A dialog that allows to view a list of defined names and select the defined name.
    /// </summary>
    public partial class SelectDefinedNameWindow : Window
    {

        /// <summary>
        /// Initializes a new instance of the <see cref="SelectDefinedNameWindow"/> class.
        /// </summary>
        /// <param name="definedNames">The defined names.</param>
        public SelectDefinedNameWindow(DefinedName[] definedNames)
        {
            InitializeComponent();

            foreach (DefinedName definedName in definedNames)
                namesListBox.Items.Add(definedName.Name);
        }

        /// <summary>
        /// Gets the selected defined name.
        /// </summary>
        public string SelectedDefinedName
        {
            get
            {
                return (string)namesListBox.SelectedItem;
            }
        }

        /// <summary>
        /// Handles the Click event of OkButton object.
        /// </summary>
        private void okButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
        }

    }
}

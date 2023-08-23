using System.Windows;

using Vintasoft.Imaging.Office.Spreadsheet.Document;
using Vintasoft.Imaging.Office.Spreadsheet.UI;

namespace WpfSpreadsheetEditorDemo
{
    /// <summary>
    /// A window that allows to view and change the document information.
    /// </summary>
    public partial class DocumentInfoWindow : Window
    {

        #region Fields

        /// <summary>
        /// Spreadsheet visual editor.
        /// </summary>
        SpreadsheetVisualEditor _visualEditor;

        #endregion



        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="DocumentInfoWindow"/> class.
        /// </summary>
        public DocumentInfoWindow()
        {
            InitializeComponent();
        }

        #endregion



        #region Methods

        /// <summary>
        /// Shows this form with current document information.
        /// </summary>
        /// <param name="visualEditor">Spreadsheet visual editor.</param>
        public static bool? ShowDialog(SpreadsheetVisualEditor visualEditor)
        {
            DocumentInfoWindow window = new DocumentInfoWindow();
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
            DocumentInformation info = _visualEditor.DocumentInformation;
            titleTextBox.Text = info.Title;
            subjectTextBox.Text = info.Subject;
            creatorTextBox.Text = info.Creator;
            managerTextBox.Text = info.Manager;
            companyTextBox.Text = info.Company;
            categoryTextBox.Text = info.Category;
            keywordsTextBox.Text = info.Keywords;
            commentsTextBox.Text = info.Comments;
            hyperlinkBaseTextBox.Text = info.HyperlinkBase;
            createdTextBox.Text = info.CreatedDate;
            applicationTextBox.Text = info.Application;
            applicationVersionTextBox.Text = info.ApplicationVersion;
            modifiedTextBox.Text = info.ModifiedDate;
            lastModifiedByTextBox.Text = info.LastModifiedBy;
            printedTextBox.Text = info.PrintedDate;
        }

        /// <summary>
        /// Handles the Click event of OkButton object.
        /// </summary>
        private void okButton_Click(object sender, RoutedEventArgs e)
        {
            DocumentInformation info = new DocumentInformation(_visualEditor.Document.Information);

            info.Title = titleTextBox.Text;
            info.Subject = subjectTextBox.Text;
            info.Creator = creatorTextBox.Text;
            info.Manager = managerTextBox.Text;
            info.Company = companyTextBox.Text;
            info.Category = categoryTextBox.Text;
            info.Keywords = keywordsTextBox.Text;
            info.Comments = commentsTextBox.Text;
            info.HyperlinkBase = hyperlinkBaseTextBox.Text;

            // if information changed, set it to visual editor
            if (!Equals(info, _visualEditor.Document.Information))
                _visualEditor.DocumentInformation = info;

            DialogResult = true;
        }

        #endregion \

    }
}

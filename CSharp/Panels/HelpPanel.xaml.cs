using System.Text;
using System.Windows;
using System.Windows.Controls;

using WpfDemosCommonCode;

namespace WpfSpreadsheetEditorDemo
{
    /// <summary>
    /// Provides the "Help" panel.
    /// </summary>
    public partial class HelpPanel : UserControl
    {

        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="HelpPanel"/> class.
        /// </summary>
        public HelpPanel()
        {
            InitializeComponent();
        }

        #endregion



        #region Methods

        /// <summary>
        /// Shows the about dialog.
        /// </summary>
        public void ShowAboutDialog()
        {
            StringBuilder description = new StringBuilder();

            description.AppendLine("This project demonstrates the following SDK capabilities:");
            description.AppendLine();
            description.AppendLine("Create a new or open an existing XLSX document in spreadsheet editor control");
            description.AppendLine();
            description.AppendLine("Work with spreadsheet document");
            description.AppendLine("- Set culture of spreadsheet document");
            description.AppendLine("- Assign settings (author, etc) of spreadsheet document");
            description.AppendLine("- Edit style properties of spreadsheet document");
            description.AppendLine("- Add/delete defined names to/from spreadsheet document");
            description.AppendLine();
            description.AppendLine("Work with worksheets of spreadsheet document");
            description.AppendLine("- Get a list of worksheets");
            description.AppendLine("- Add/delete/rename a worksheet; copy/insert a worksheet; reorder worksheets");
            description.AppendLine();
            description.AppendLine("Work with worksheet of spreadsheet document");
            description.AppendLine("- Render a worksheet");
            description.AppendLine("- Change settings of worksheet view");
            description.AppendLine("- Navigate by cells using mouse and keyboard");
            description.AppendLine("- Insert or delete columns/rows");
            description.AppendLine("- Change size of columns/rows");
            description.AppendLine("- Show/hide columns/rows");
            description.AppendLine("- Search and replace text");
            description.AppendLine();
            description.AppendLine("Work with selected cells of worksheet");
            description.AppendLine("- Select cells using mouse and keyboard");
            description.AppendLine("- Insert, copy, paste and delete selected cells");
            description.AppendLine("- Change style properties (font, fill, borders, number format, text style, alignment, indent, etc) of selected cells");
            description.AppendLine("- Change size of selected cells");
            description.AppendLine("- Auto-fit column width or row height of selected cells");
            description.AppendLine("- Clear styles, content, hyperlinks of selected cells");
            description.AppendLine("- Merge and unmerge selected cells");
            description.AppendLine("- Show and hide selected cells");
            description.AppendLine("- Set the hyperlinks for selected cells");
            description.AppendLine();
            description.AppendLine("Work with cell of worksheet");
            description.AppendLine("- Display formatted and localized text of cell");
            description.AppendLine("- Calculate formula of cell");
            description.AppendLine("- Edit cell text directly in cell region");
            description.AppendLine("- Edit cell text in formula bar");
            description.AppendLine("- Highlight references while editing a cell formula");
            description.AppendLine();
            description.AppendLine("Work with Drawing (Charts, Pictures, Graphics)");
            description.AppendLine("- Render drawings on worksheet");
            description.AppendLine("- Update a chart if chart data has changed");
            description.AppendLine("- Select a drawing on worksheet");
            description.AppendLine("- Add drawing to a worksheet");
            description.AppendLine("- Delete drawing from worksheet");
            description.AppendLine();
            description.AppendLine("Work with comments");
            description.AppendLine("- Render comments on worksheet");
            description.AppendLine("- Add, edit, delete a comment");
            description.AppendLine();
            description.AppendLine();
            description.AppendLine("The project is available in C# and VB.NET for Visual Studio .NET.");


            WpfAboutBoxBaseWindow dlg = new WpfAboutBoxBaseWindow("vsoffice-dotnet");
            dlg.Description = description.ToString();
            dlg.Owner = Application.Current.MainWindow;
            dlg.ShowDialog();
        }

        /// <summary>
        /// Handles the Click event of AboutButton object.
        /// </summary>
        private void aboutButton_Click(object sender, RoutedEventArgs e)
        {
            ShowAboutDialog();
        }

        #endregion

    }
}

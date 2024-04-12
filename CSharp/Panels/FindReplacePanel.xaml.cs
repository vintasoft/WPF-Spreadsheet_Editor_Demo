using System.Windows;

using Vintasoft.Imaging;
using Vintasoft.Imaging.Office.Spreadsheet.Wpf.UI;
using Vintasoft.Imaging.Office.Spreadsheet.Wpf.UI.Controls;

namespace WpfSpreadsheetEditorDemo
{
    /// <summary>
    /// Provides a "Find and Replace" panel.
    /// </summary>
    public partial class FindReplacePanel : SpreadsheetVisualEditorPanel
    {

        #region Fields

        /// <summary>
        /// The window that allows to find and replace text.
        /// </summary>
        FindReplaceWindow _findReplaceWindow;

        #endregion



        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="FindReplacePanel"/> class.
        /// </summary>
        public FindReplacePanel()
        {
            InitializeComponent();
        }

        #endregion



        #region Methods  

        #region PUBLIC

        /// <summary>
        /// Shows the find dialog.
        /// </summary>
        public void ShowFindDialog()
        {
            ShowFindReplaceForm(false);
        }

        /// <summary>
        /// Shows the replace dialog.
        /// </summary>
        public void ShowReplaceDialog()
        {
            ShowFindReplaceForm(true);
        }

        #endregion


        #region PROTECTED

        /// <summary>
        /// Executes when the spreadsheet editor is changed.
        /// </summary>
        /// <param name="args">The <see cref="PropertyChangedEventArgs{SpreadsheetEditorControl}"/> instance containing the event data.</param>
        protected override void OnSpreadsheetEditorChanged(PropertyChangedEventArgs<WpfSpreadsheetEditorControl> args)
        {
            base.OnSpreadsheetEditorChanged(args);

            if (_findReplaceWindow != null)
            {
                if (args.NewValue != null)
                    _findReplaceWindow.VisualEditor = args.NewValue.VisualEditor;
                else
                    _findReplaceWindow.VisualEditor = null;
            }
        }

        #endregion


        #region PRIVATE

        /// <summary>
        /// Handles the Unloaded event of FindReplacePanel object.
        /// </summary>
        private void FindReplacePanel_Unloaded(object sender, RoutedEventArgs e)
        {
            if (_findReplaceWindow != null)
            {
                _findReplaceWindow.Close();
                _findReplaceWindow = null;
            }
        }

        /// <summary>
        /// Handles the Click event of findButton object.
        /// </summary>
        private void findButton_Click(object sender, RoutedEventArgs e)
        {
            ShowFindDialog();
        }

        /// <summary>
        /// Handles the Click event of replaceButton object.
        /// </summary>
        private void replaceButton_Click(object sender, RoutedEventArgs e)
        {
            ShowReplaceDialog();
        }

        /// <summary>
        /// Shows the find and replace form.
        /// </summary>
        /// <param name="replaceMode">Replace mode.</param>
        private void ShowFindReplaceForm(bool replaceMode)
        {
            bool needSetLocation = false;
            if (_findReplaceWindow == null)
            {
                _findReplaceWindow = new FindReplaceWindow();
                _findReplaceWindow.VisualEditor = VisualEditor;
                needSetLocation = true;
            }
            _findReplaceWindow.ReplaceMode = replaceMode;
            if (_findReplaceWindow.Visibility == Visibility.Visible)
            {
                _findReplaceWindow.Focus();
            }
            else
            {
                _findReplaceWindow.Show();
                if (needSetLocation)
                {
                    Window mainForm = Window.GetWindow(this);

                    Point location = new Point(mainForm.Left, mainForm.Top);
                    if (mainForm.WindowState == WindowState.Maximized)
                        location = new Point(0, 0);

                    _findReplaceWindow.Left = location.X + mainForm.ActualWidth - _findReplaceWindow.ActualWidth - 20;
                    _findReplaceWindow.Top = location.Y + mainForm.ActualHeight - _findReplaceWindow.ActualHeight - 150;
                }
            }
            _findReplaceWindow.Reset();
        }

        #endregion

        #endregion

    }
}

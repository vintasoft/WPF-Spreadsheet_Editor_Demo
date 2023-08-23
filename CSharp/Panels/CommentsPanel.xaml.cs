using System.Windows;

using Vintasoft.Imaging;
using Vintasoft.Imaging.Office.Spreadsheet.Document;
using Vintasoft.Imaging.Office.Spreadsheet.UI;
using Vintasoft.Imaging.Office.Spreadsheet.Wpf.UI;

namespace WpfSpreadsheetEditorDemo
{
    /// <summary>
    /// Provides the "Comments" panel.
    /// </summary>
    public partial class CommentsPanel : SpreadsheetVisualEditorPanel
    {

        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="CommentsPanel"/> class.
        /// </summary>
        public CommentsPanel()
        {
            InitializeComponent();
        }

        #endregion



        #region Methods

        #region PUBLIC

        /// <summary>
        /// Creates a new comment for focused cell.
        /// </summary>
        public void NewComment()
        {
            // get the focused cell
            CellReference focusedCell = VisualEditor.FocusedCell;

            // create dialog that allows to add the comment
            EditCommentWindow dlg = new EditCommentWindow(VisualEditor);
            dlg.WindowStartupLocation = WindowStartupLocation.CenterOwner;
            dlg.Owner = Application.Current.MainWindow;

            // show the dialog
            if (dlg.ShowDialog() == true)
            {
                CellComment newCellComment = new CellComment(focusedCell, dlg.Comment, true, dlg.CommentLocation);
                VisualEditor.SetCellComment(newCellComment);

                VisualEditor.FocusedComment = VisualEditor.FocusedCellComment;
            }
        }

        /// <summary>
        /// Edits the focused comment.
        /// </summary>
        internal void EditComment()
        {
            CellComment sourceCellComment = VisualEditor.FocusedComment ?? VisualEditor.FocusedCellComment;
            Comment sourceComment = sourceCellComment.Comment;
            SheetDrawingLocation sourceLocation = sourceCellComment.Location;

            // create dialog that allows to edit the comment
            EditCommentWindow dlg = new EditCommentWindow(VisualEditor, sourceComment, sourceLocation);
            dlg.WindowStartupLocation = WindowStartupLocation.CenterOwner;
            dlg.Owner = Application.Current.MainWindow;
            // show the dialog
            if (dlg.ShowDialog() == true)
            {
                VisualEditor.StartEditing("Edit comment");
                try
                {
                    VisualEditor.SetComment(dlg.Comment);
                    VisualEditor.SetCommentLocation(dlg.CommentLocation);
                }
                finally
                {
                    VisualEditor.FinishEditing();
                }
            }
        }

        #endregion


        #region PROTECTED

        /// <summary>
        /// Raises the <see cref="E:SpreadsheetEditorChanged" /> event.
        /// </summary>
        /// <param name="args">The <see cref="PropertyChangedEventArgs{SpreadsheetEditorControl}"/> instance containing the event data.</param>
        protected override void OnSpreadsheetEditorChanged(PropertyChangedEventArgs<WpfSpreadsheetEditorControl> args)
        {
            base.OnSpreadsheetEditorChanged(args);

            if (args.OldValue != null)
            {
                SpreadsheetVisualEditor visualEditor = args.OldValue.VisualEditor;
                visualEditor.FocusedWorksheetChanged -= VisualEditor_FocusedWorksheetChanged;
                visualEditor.FocusedCellChanged -= VisualEditor_FocusedCellChanged;
                visualEditor.FocusedCommentChanged -= VisualEditor_FocusedCommentChanged;
                args.OldValue.PreviewMouseDoubleClick -= SpreadsheetEditorControl_MouseDoubleClick;
            }

            if (args.NewValue != null)
            {
                SpreadsheetVisualEditor visualEditor = args.NewValue.VisualEditor;
                visualEditor.FocusedWorksheetChanged += VisualEditor_FocusedWorksheetChanged;
                visualEditor.FocusedCellChanged += VisualEditor_FocusedCellChanged;
                visualEditor.FocusedCommentChanged += VisualEditor_FocusedCommentChanged;
                args.NewValue.PreviewMouseDoubleClick += SpreadsheetEditorControl_MouseDoubleClick;
            }

            UpdateUI();
        }

        #endregion


        #region PRIVATE

        #region UI

        /// <summary>
        /// Handles the Click event of NewButton object.
        /// </summary>
        private void newButton_Click(object sender, RoutedEventArgs e)
        {
            NewComment();
        }

        /// <summary>
        /// Handles the Click event of EditButton object.
        /// </summary>
        private void editButton_Click(object sender, RoutedEventArgs e)
        {
            EditComment();
        }

        /// <summary>
        /// Handles the Click event of DeleteButton object.
        /// </summary>
        private void deleteButton_Click(object sender, RoutedEventArgs e)
        {
            VisualEditor.RemoveComments();
            UpdateUI();
        }

        /// <summary>
        /// Handles the Click event of PrevButton object.
        /// </summary>
        private void prevButton_Click(object sender, RoutedEventArgs e)
        {
            int currentIndex = 0;
            // get focused comment
            CellComment focusedComment = VisualEditor.FocusedComment;
            if (focusedComment == null)
            {
                // get focused cell comment
                focusedComment = VisualEditor.FocusedCellComment;
            }

            if (focusedComment != null)
            {
                // get index of comment
                currentIndex = VisualEditor.FocusedWorksheet.CellComments.IndexOf(focusedComment);
            }

            // get previous comment index
            if (currentIndex == 0)
                currentIndex = VisualEditor.FocusedWorksheet.CellComments.Count - 1;
            else
                currentIndex--;

            // focus previous comment
            VisualEditor.FocusedComment = VisualEditor.FocusedWorksheet.CellComments[currentIndex];
        }

        /// <summary>
        /// Handles the Click event of NextButton object.
        /// </summary>
        private void nextButton_Click(object sender, RoutedEventArgs e)
        {
            int currentIndex = 0;
            // get focused comment
            CellComment focusedComment = VisualEditor.FocusedComment;
            if (focusedComment == null)
            {
                // get focused cell comment
                focusedComment = VisualEditor.FocusedCellComment;
            }

            if (focusedComment != null)
            {
                // get index of comment
                currentIndex = VisualEditor.FocusedWorksheet.CellComments.IndexOf(focusedComment);
            }

            // get next comment index
            if (currentIndex == VisualEditor.FocusedWorksheet.CellComments.Count - 1)
                currentIndex = 0;
            else
                currentIndex++;

            // focus next comment
            VisualEditor.FocusedComment = VisualEditor.FocusedWorksheet.CellComments[currentIndex];
        }

        /// <summary>
        /// Handles the Click event of ShowHideButton object.
        /// </summary>
        private void showHideButton_Click(object sender, RoutedEventArgs e)
        {
            if (VisualEditor.FocusedCellComment != null)
            {
                VisualEditor.SetCommentIsVisible(!VisualEditor.FocusedCellComment.IsVisible);
                UpdateUI();
            }
        }

        /// <summary>
        /// Handles the Click event of ShowAllButton object.
        /// </summary>
        private void showAllButton_Click(object sender, RoutedEventArgs e)
        {
            if (VisualEditor.SelectionContainsSingleCell)
                VisualEditor.ShowAllComments();
            else
                VisualEditor.ShowComments();
            UpdateUI();
        }

        /// <summary>
        /// Handles the Click event of HideAllButton object.
        /// </summary>
        private void hideAllButton_Click(object sender, RoutedEventArgs e)
        {
            VisualEditor.FocusedComment = null;
            if (VisualEditor.SelectionContainsSingleCell)
                VisualEditor.HideAllComments();
            else
                VisualEditor.HideComments();
            UpdateUI();
        }

        private void VisualEditor_FocusedWorksheetChanged(object sender, PropertyChangedEventArgs<Worksheet> e)
        {
            UpdateUI();
        }
        private void VisualEditor_FocusedCommentChanged(object sender, PropertyChangedEventArgs<CellComment> e)
        {
            UpdateUI();
        }

        private void VisualEditor_FocusedCellChanged(object sender, PropertyChangedEventArgs<CellReference> e)
        {
            UpdateUI();
        }

        /// <summary>
        /// Handles the MouseDoubleClick event of SpreadsheetEditorControl object.
        /// </summary>
        private void SpreadsheetEditorControl_MouseDoubleClick(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (e.ChangedButton == System.Windows.Input.MouseButton.Left && VisualEditor.FocusedComment != null)
            {
                e.Handled = true;
                EditComment();
            }
        }

        /// <summary>
        /// Updates the User Interface.
        /// </summary>
        private void UpdateUI()
        {
            if (VisualEditor.FocusedWorksheet == null)
            {
                IsEnabled = false;
            }
            else
            {
                IsEnabled = true;
                bool hasFocusedComment = VisualEditor.FocusedComment != null;
                bool focusedCellHasComment = VisualEditor.FocusedCellComment != null;
                bool hasFocusedCell = VisualEditor.FocusedCell != null;
                bool sheetHasComments = VisualEditor.FocusedWorksheet.CellComments.Count > 0;
                newButton.IsEnabled = hasFocusedCell && !focusedCellHasComment;
                editButton.IsEnabled = hasFocusedComment || hasFocusedCell && focusedCellHasComment;
                deleteButton.IsEnabled = VisualEditor.SelectedCellsHasComments;
                prevButton.IsEnabled = sheetHasComments;
                nextButton.IsEnabled = sheetHasComments;
                showHideButton.IsEnabled = focusedCellHasComment || hasFocusedComment;
                showAllButton.IsEnabled = sheetHasComments;
                hideAllButton.IsEnabled = sheetHasComments;
            }
        }


        #endregion

        #endregion

        #endregion

    }
}

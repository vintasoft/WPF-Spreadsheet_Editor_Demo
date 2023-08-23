using System;
using System.IO;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

using Microsoft.Win32;

using Vintasoft.Imaging;
using Vintasoft.Imaging.Office.Spreadsheet.Document;
using Vintasoft.Imaging.Office.Spreadsheet.UI;

using WpfDemosCommonCode;
using WpfDemosCommonCode.Imaging;
using WpfDemosCommonCode.Imaging.Codecs;

namespace WpfSpreadsheetEditorDemo
{
    /// <summary>
    /// Main window.
    /// </summary>
    public partial class MainWindow : Window
    {

        #region Fields

        #region Context Menu

        /// <summary>
        /// The cells context menu.
        /// </summary>
        ContextMenu _cellsContextMenu;

        /// <summary>
        /// The drawing context menu.
        /// </summary>
        ContextMenu _drawingContextMenu;

        /// <summary>
        /// The comment context menu.
        /// </summary>
        ContextMenu _commentContextMenu;

        /// <summary>
        /// The "Set Image" drawing context menu item.
        /// </summary>
        MenuItem _drawingSetImageMenuItem;

        /// <summary>
        /// The "Remove Link" drawing context menu item.
        /// </summary>
        MenuItem _drawingRemoveLinkMenuItem;

        #endregion

        #endregion



        #region Classes

        /// <summary>
        /// SetStatus delegate.
        /// </summary>
        /// <param name="status">The status.</param>
        delegate void SetStatusDelegate(string status);

        #endregion



        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="MainWindow"/> class.
        /// </summary>
        public MainWindow()
        {
            // register the evaluation license for VintaSoft Imaging .NET SDK
            Vintasoft.Imaging.ImagingGlobalSettings.Register("REG_USER", "REG_EMAIL", "EXPIRATION_DATE", "REG_CODE");

            // set CustomFontProgramsController for all opened documents
            CustomFontProgramsController.SetDefaultFontProgramsController();

            InitializeComponent();

            VisualEditor.SynchronizationStarted += VisualEditor_SynchronizationStarted;
            VisualEditor.SynchronizationFinished += VisualEditor_SynchronizationFinished;
            VisualEditor.SynchronizationException += VisualEditor_SynchronizationException;
            VisualEditor.HoveredHyperlinkChanged += VisualEditor_HoveredHyperlinkChanged;
            VisualEditor.HoveredDrawingChanged += VisualEditor_HoveredDrawingChanged;
            VisualEditor.HoveredCellChanged += VisualEditor_HoveredCellChanged;
            VisualEditor.UriOpen += VisualEditor_UriOpen;
            VisualEditor.CellErrorClick += VisualEditor_CellErrorClick;
            VisualEditor.CellCommentClick += VisualEditor_CellCommentClick;
            VisualEditor.InvalidCellReferences += VisualEditor_InvalidCellReferences;
            VisualEditor.FocusedCellChanged += VisualEditor_FocusedCellChanged;
            VisualEditor.ContextMenuOpen += VisualEditor_ContextMenuOpen;
            VisualEditor.Editor = null;
            SetStatus("");

            // init spreadsheet editor context menus
            _cellsContextMenu = spreadsheetEditorControl1.ContextMenu;
            _commentContextMenu = FindResource("commentContextMenu") as ContextMenu;
            _drawingContextMenu = FindResource("drawingContextMenu") as ContextMenu;
            _drawingSetImageMenuItem = (MenuItem)_drawingContextMenu.Items[0];
            _drawingRemoveLinkMenuItem = (MenuItem)_drawingContextMenu.Items[3];

            UpdateUI();
        }

        #endregion



        #region Properties

        /// <summary>
        /// Gets the visual editor.
        /// </summary>
        public SpreadsheetVisualEditor VisualEditor
        {
            get
            {
                return spreadsheetEditorControl1.VisualEditor;
            }
        }

        #endregion



        #region Methods

        #region Common

        /// <summary>
        /// Handles the PreviewMouseDoubleClick event of SpreadsheetEditorControl1 object.
        /// </summary>
        private void spreadsheetEditorControl1_PreviewMouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left)
            {
                if (VisualEditor.FocusedDrawing != null)
                {
                    ShowFocusedDrawingPropertiesDialog();
                    e.Handled = true;
                }
            }
        }

        /// <summary>
        /// Handles the Closing event of Window object.
        /// </summary>
        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (!filePanel1.CheckChanges())
                e.Cancel = true;
        }

        private void VisualEditor_HoveredDrawingChanged(object sender, PropertyChangedEventArgs<SheetDrawing> e)
        {
            if (e.NewValue != null)
                SetStatus(e.NewValue.Name);
            else
                SetStatus("");
        }


        private void VisualEditor_HoveredCellChanged(object sender, PropertyChangedEventArgs<CellReference> e)
        {
            if (e.NewValue != null)
            {
                CellComment cellComment = VisualEditor.FocusedWorksheet.GetCellComment(e.NewValue);
                if (cellComment != null)
                {
                    string text = cellComment.Comment.Text;
                    text = text.Replace(Environment.NewLine, " ");
                    text = text.Replace("\n", " ");
                    SetStatus(cellComment.Comment.Author + ": " + text);
                }
                else if (e.OldValue != null)
                {
                    cellComment = VisualEditor.FocusedWorksheet.GetCellComment(e.OldValue);
                    if (cellComment != null)
                        SetStatus("");
                }
            }
        }

        private void VisualEditor_CellCommentClick(object sender, SheetCellMouseEventArgs e)
        {
            VisualEditor.SetCommentIsVisible(!VisualEditor.FocusedCellComment.IsVisible);
            e.Handled = true;
        }

        private void VisualEditor_UriOpen(object sender, UriEventArgs e)
        {
            try
            {
                Uri uri = new Uri(e.Uri);
                DemosTools.OpenBrowser(uri.AbsoluteUri);
            }
            catch
            {
                DemosTools.ShowWarningMessage("The address of this site is not valid: " + e.Uri);
            }
        }

        private void VisualEditor_HoveredHyperlinkChanged(object sender, PropertyChangedEventArgs<Vintasoft.Imaging.Office.Spreadsheet.Document.Hyperlink> e)
        {
            Vintasoft.Imaging.Office.Spreadsheet.Document.Hyperlink hyperlink = e.NewValue;
            if (hyperlink == null)
            {
                SetStatus("");
            }
            else
            {
                string hyperlinkDecription;
                if (hyperlink.Name != null)
                    hyperlink = spreadsheetEditorControl1.VisualEditor.GetHyperlinkByDefinedName(hyperlink.Name);
                if (!string.IsNullOrEmpty(hyperlink.Url))
                    hyperlinkDecription = hyperlink.Url;
                else if (hyperlink.Location != null)
                    hyperlinkDecription = hyperlink.Location.ToString();
                else
                    hyperlinkDecription = hyperlink.Name;
                hyperlinkDecription = "Link: " + hyperlinkDecription;
                if (VisualEditor.HoveredDrawing != null)
                    SetStatus(VisualEditor.HoveredDrawing.Name + ": " + hyperlinkDecription);
                else
                    SetStatus(hyperlinkDecription);
            }
        }

        private void VisualEditor_CellErrorClick(object sender, SheetCellMouseEventArgs e)
        {
            Worksheet worksheet = spreadsheetEditorControl1.VisualEditor.FocusedWorksheet;
            SheetCell cell = worksheet.FindCell(e.Cell);
            string errorMessage = GetErrorMessage(cell.ErrorType);
            DemosTools.ShowWarningMessage("Error: " + cell.Value, errorMessage);
            e.Handled = true;
        }

        private void VisualEditor_InvalidCellReferences(object sender, CellReferencesEventArgs e)
        {
            MessageBox.Show("Reference is not valid.", "Error", MessageBoxButton.OK, MessageBoxImage.Exclamation);
        }

        private void VisualEditor_FocusedCellChanged(object sender, PropertyChangedEventArgs<CellReference> e)
        {
            if (e.NewValue != null)
            {
                SheetCell cell = VisualEditor.FocusedWorksheet.FindCell(e.NewValue);
                if (cell != null && cell.ErrorType != CellErrorType.NoError)
                {
                    SetStatus(string.Format("Cell {0} with formula '={1}' calculation error: {2}", e.NewValue, cell.Formula, GetErrorMessage(cell.ErrorType)));
                    return;
                }
            }
            if (e.OldValue != null)
            {
                SheetCell cell = VisualEditor.FocusedWorksheet.FindCell(e.OldValue);
                if (cell != null && cell.ErrorType != CellErrorType.NoError)
                {
                    SetStatus("");
                }
            }
        }

        private void VisualEditor_ContextMenuOpen(object sender, Vintasoft.Imaging.UI.VintasoftControlMouseEventArgs e)
        {
            // if context menu for focused cell should be shown
            if (VisualEditor.FocusedComment != null)
            {
                spreadsheetEditorControl1.ContextMenu = _commentContextMenu;
            }
            // if context menu for focused drawing should be shown
            else if (VisualEditor.FocusedDrawing != null)
            {
                _drawingSetImageMenuItem.IsEnabled = VisualEditor.FocusedDrawing.Type == DrawingType.Picture;
                _drawingRemoveLinkMenuItem.IsEnabled = VisualEditor.FocusedDrawing.Hyperlink != null;
                spreadsheetEditorControl1.ContextMenu = _drawingContextMenu;
            }
            // if context menu for focused cells should be shown
            else
            {
                // selection contains whole columns
                bool isCoverColumns = VisualEditor.FocusedCells.IsCoverColumns;
                // selection contains whole rows
                bool isCoverRows = VisualEditor.FocusedCells.IsCoverRows;
                // context menu was opened on "select all" button (top left corner of the grid)
                bool isSelectAllHovered = VisualEditor.HoveredCell.RowIndex < 0 && VisualEditor.HoveredCell.ColumnIndex < 0;

                // all cells selected or selection does not have whole rows or columns
                bool addCellsMenuItems = isSelectAllHovered || !isCoverColumns && !isCoverRows;
                // menu was opened on column header or selection contains whole columns, but not whole rows
                bool addColumnsMenuItems = !addCellsMenuItems && (VisualEditor.HoveredCell.RowIndex < 0 || isCoverColumns && !isCoverRows);
                // none of the previous conditions were met and selection contains whole rows
                bool addRowsMenuItems = !addCellsMenuItems && !addColumnsMenuItems && isCoverRows;

                // add cells menu items
                SetVisibility(insertCellsMenuItem, addCellsMenuItems);
                SetVisibility(deleteCellsMenuItem, addCellsMenuItems);
                SetVisibility(defineNameMenuItem, addCellsMenuItems);
                SetVisibility(linkMenuItem, addCellsMenuItems);
                SetVisibility(removeLinkMenuItem, addCellsMenuItems);

                // focused cell comment
                bool cellHasComment = VisualEditor.FocusedCellComment != null;
                bool isSingleCellSelection = VisualEditor.SelectionContainsSingleCell;
                // add cells comment menu items
                SetVisibility(insertCommentMenuItem, addCellsMenuItems && !cellHasComment && isSingleCellSelection);
                SetVisibility(editCellCommentMenuItem, addCellsMenuItems && cellHasComment && isSingleCellSelection);
                SetVisibility(showHideCommentMenuItem, addCellsMenuItems && cellHasComment && isSingleCellSelection);
                SetVisibility(deleteCellCommentMenuItem, addCellsMenuItems && VisualEditor.SelectedCellsHasComments);
                SetVisibility(commentSectionSeparator, addCellsMenuItems && (isSingleCellSelection || VisualEditor.SelectedCellsHasComments));

                // add columns menu items
                SetVisibility(columnWidthMenuItem, addColumnsMenuItems);
                SetVisibility(insertColumnsMenuItem, addColumnsMenuItems);
                SetVisibility(deleteColumnsMenuItem, addColumnsMenuItems);
                SetVisibility(hideColumnsMenuItem, addColumnsMenuItems);
                SetVisibility(unhideColumnsMenuItem, addColumnsMenuItems);

                // add rows menu items
                SetVisibility(rowHeightMenuItem, addRowsMenuItems);
                SetVisibility(insertRowsMenuItem, addRowsMenuItems);
                SetVisibility(deleteRowsMenuItem, addRowsMenuItems);
                SetVisibility(hideRowsMenuItem, addRowsMenuItems);
                SetVisibility(unhideRowsMenuItem, addRowsMenuItems);

                spreadsheetEditorControl1.ContextMenu = _cellsContextMenu;
            }
        }

        /// <summary>
        /// Sets the control visibility value depending on specified value.
        /// </summary>
        /// <param name="control">The control to set visibility to.</param>
        /// <param name="isVisible">A value indicating whether the control should be visible.</param>
        private void SetVisibility(Control control, bool isVisible)
        {
            if (isVisible)
                control.Visibility = Visibility.Visible;
            else
                control.Visibility = Visibility.Collapsed;
        }

        private void VisualEditor_SynchronizationFinished(object sender, EventArgs e)
        {
            SetStatus("");
        }

        private void VisualEditor_SynchronizationStarted(object sender, EventArgs e)
        {
            if (VisualEditor.IsInitialized)
                SetStatus("Processing...");
            else
                SetStatus("Loading...");
        }

        private void VisualEditor_SynchronizationException(object sender, Vintasoft.Imaging.ExceptionEventArgs e)
        {
            DemosTools.ShowErrorMessage(e.Exception);
            filePanel1.CloseDocument(false);
        }

        private void SetStatus(string status)
        {
            if (Dispatcher.Thread != Thread.CurrentThread)
            {
                Dispatcher.Invoke(new SetStatusDelegate(SetStatus), status);
            }
            else
            {
                if (string.IsNullOrEmpty(status))
                    statusLabel.Content = "Ready";
                else
                    statusLabel.Content = status;
            }
        }

        private void filePanel1_FilenameChanged(object sender, EventArgs e)
        {
            UpdateUI();
        }

        /// <summary>
        /// Updates the User Interface.
        /// </summary>
        private void UpdateUI()
        {
            Title = "VintaSoft WPF Spreadsheet Editor Demo v" + ImagingGlobalSettings.ProductVersion;
            if (!string.IsNullOrEmpty(filePanel1.Filename))
                Title += " - " + Path.GetFileName(filePanel1.Filename);
        }


        /// <summary>
        /// Returns the error message.
        /// </summary>
        /// <param name="errorType">Type of the error.</param>
        private string GetErrorMessage(CellErrorType errorType)
        {
            switch (errorType)
            {
                case CellErrorType.Unknown:
                    return "Unknown error.";
                case CellErrorType.DivByZero:
                    return "Any number (including zero) or any error code is divided by zero.";
                case CellErrorType.External:
                    return "External error.";
                case CellErrorType.GettingData:
                    return "A cell reference cannot be evaluated because the value for the cell has not been retrieved or calculated.";
                case CellErrorType.Name:
                    return "Looks like a name is used but no such name has been defined.";
                case CellErrorType.NoError:
                    return "No error.";
                case CellErrorType.NotANumber:
                    return "A designated value is not available.";
                case CellErrorType.Null:
                    return "Two areas are required to intersect but do not.";
                case CellErrorType.Num:
                    return "An argument to a function has a compatible type but has a value that is outside the domain over which that function is defined.";
                case CellErrorType.Ref:
                    return "A cell reference cannot be evaluated.";
                case CellErrorType.Value:
                    return "An incompatible type argument is passed to a function, or an incompatible type operand is used with an operator.";
            }
            return "Unexpected error.";
        }

        /// <summary>
        /// Handles the PreviewKeyDown event of Window object.
        /// </summary>
        private void Window_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                bool isControlKeyPressed = e.KeyboardDevice.IsKeyDown(Key.LeftCtrl) || e.KeyboardDevice.IsKeyDown(Key.RightCtrl);
                bool isShiftKeyPressed = e.KeyboardDevice.IsKeyDown(Key.LeftShift) || e.KeyboardDevice.IsKeyDown(Key.RightShift);
               
                // Ctrl+O
                if (isControlKeyPressed && e.Key == Key.O)
                {
                    if (filePanel1.OpenDocument())
                        e.Handled = true;
                }

                // Ctrl+N
                if (isControlKeyPressed && e.Key == Key.N)
                {
                    if (filePanel1.NewDocument())
                        e.Handled = true;
                }

                // Ctrl+S
                if (isControlKeyPressed && e.Key == Key.S)
                {
                    if (filePanel1.SaveDocumentChanges())
                        e.Handled = true;
                }

                // Ctrl+Shift+S
                if (isControlKeyPressed && isShiftKeyPressed && e.Key == Key.S)
                {
                    if (filePanel1.SaveDocumentAs())
                        e.Handled = true;
                }

                // Ctrl+P
                if (isControlKeyPressed && e.Key == Key.P)
                {
                    if (filePanel1.PrintDocument())
                        e.Handled = true;
                }             

                // Ctrl+F
                if (isControlKeyPressed && e.Key == Key.F)
                {
                    if (VisualEditor.FocusedWorksheet != null)
                        findReplacePanel1.ShowFindDialog();
                }

                // Ctrl+H
                if (isControlKeyPressed && e.Key == Key.H)
                {
                    if (VisualEditor.FocusedWorksheet != null)
                        findReplacePanel1.ShowReplaceDialog();
                }

                // F1
                if (e.Key == Key.F1)
                {
                    helpPanel1.ShowAboutDialog();
                    e.Handled = true;
                }
            }
            catch (Exception ex)
            {
                DemosTools.ShowErrorMessage(ex);
                e.Handled = true;
            }
        }

        #endregion


        #region Comment context menu

        /// <summary>
        /// Handles the Click event of EditCommentMenuItem object.
        /// </summary>
        private void editCommentMenuItem_Click(object sender, RoutedEventArgs e)
        {
            commentsPanel.EditComment();
        }

        /// <summary>
        /// Handles the Click event of DeleteCommentMenuItem object.
        /// </summary>
        private void deleteCommentMenuItem_Click(object sender, RoutedEventArgs e)
        {
            VisualEditor.RemoveFocusedComment();
        }

        /// <summary>
        /// Handles the Click event of HideCommentMenuItem object.
        /// </summary>
        private void hideCommentMenuItem_Click(object sender, RoutedEventArgs e)
        {
            VisualEditor.SetCommentIsVisible(false);
            VisualEditor.FocusedComment = null;
        }

        #endregion


        #region Drawing context menu

        /// <summary>
        /// Handles the Click event of DrawingSetImageMenuItem object.
        /// </summary>
        private void drawingSetImageMenuItem_Click(object sender, RoutedEventArgs e)
        {
            // create dialog
            OpenFileDialog dialog = new OpenFileDialog();
            // specify that dialog should open folder with demo images
            DemosTools.SetTestFilesFolder(dialog);
            // set image filters
            CodecsFileFilters.SetFilters(dialog);

            // if image must be changed
            if (dialog.ShowDialog() == true)
                using (Stream stream = dialog.OpenFile())
                    VisualEditor.SetDrawingPicture(new ImageData(stream));
        }

        /// <summary>
        /// Handles the Click event of DeleteDrawingMenuItem object.
        /// </summary>
        private void deleteDrawingMenuItem_Click(object sender, RoutedEventArgs e)
        {
            VisualEditor.RemoveFocusedDrawing();
        }

        /// <summary>
        /// Handles the Click event of DrawingLinkMenuItem object.
        /// </summary>
        private void drawingLinkMenuItem_Click(object sender, RoutedEventArgs e)
        {
            EditHyperlinkWindow.ShowDialog(VisualEditor, VisualEditor.FocusedDrawing.Hyperlink != null);
        }

        /// <summary>
        /// Handles the Click event of DrawingRemoveLinkMenuItem object.
        /// </summary>
        private void drawingRemoveLinkMenuItem_Click(object sender, RoutedEventArgs e)
        {
            VisualEditor.RemoveHyperlinks();
        }

        /// <summary>
        /// Handles the Click event of DrawingPropertiesMenuItem object.
        /// </summary>
        private void drawingPropertiesMenuItem_Click(object sender, RoutedEventArgs e)
        {
            ShowFocusedDrawingPropertiesDialog();
        }

        private void ShowFocusedDrawingPropertiesDialog()
        {
            DrawingPropertiesWindow window = new DrawingPropertiesWindow(VisualEditor, VisualEditor.FocusedDrawing);
            window.WindowStartupLocation = WindowStartupLocation.CenterOwner;
            window.Owner = this;

            window.ShowDialog();
        }

        #endregion


        #region Cells context menu

        /// <summary>
        /// Handles the Click event of CopyMenuItem object.
        /// </summary>
        private void copyMenuItem_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                VisualEditor.Copy();
            }
            catch (Exception ex)
            {
                DemosTools.ShowErrorMessage(ex);
            }

        }

        /// <summary>
        /// Handles the Click event of CutMenuItem object.
        /// </summary>
        private void cutMenuItem_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                VisualEditor.Cut();
            }
            catch (Exception ex)
            {
                DemosTools.ShowErrorMessage(ex);
            }
        }

        /// <summary>
        /// Handles the Click event of PasteMenuItem object.
        /// </summary>
        private void pasteMenuItem_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                VisualEditor.PasteWithFill = Keyboard.IsKeyDown(Key.LeftCtrl) || Keyboard.IsKeyDown(Key.RightCtrl);
                VisualEditor.Paste();
            }
            catch (Exception ex)
            {
                DemosTools.ShowErrorMessage(ex);
            }
            finally
            {
                VisualEditor.PasteWithFill = false;
            }
        }

        /// <summary>
        /// Handles the Click event of InsertColumnsMenuItem object.
        /// </summary>
        private void insertColumnsMenuItem_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                VisualEditor.InsertEmptyColumns();
            }
            catch (Exception ex)
            {
                DemosTools.ShowErrorMessage(ex);
            }
        }

        /// <summary>
        /// Handles the Click event of InsertRowsMenuItem object.
        /// </summary>
        private void insertRowsMenuItem_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                VisualEditor.InsertEmptyRows();
            }
            catch (Exception ex)
            {
                DemosTools.ShowErrorMessage(ex);
            }
        }

        /// <summary>
        /// Handles the Click event of DeleteColumnsMenuItem object.
        /// </summary>
        private void deleteColumnsMenuItem_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                VisualEditor.RemoveColumns();
            }
            catch (Exception ex)
            {
                DemosTools.ShowErrorMessage(ex);
            }
        }

        /// <summary>
        /// Handles the Click event of DeleteRowsMenuItem object.
        /// </summary>
        private void deleteRowsMenuItem_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                VisualEditor.RemoveRows();
            }
            catch (Exception ex)
            {
                DemosTools.ShowErrorMessage(ex);
            }
        }

        /// <summary>
        /// Handles the Click event of ShiftCellsRightMenuItem object.
        /// </summary>
        private void shiftCellsRightMenuItem_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                VisualEditor.InsertCellsAndShiftRight();
            }
            catch (Exception ex)
            {
                DemosTools.ShowErrorMessage(ex);
            }
        }

        /// <summary>
        /// Handles the Click event of ShiftCellsDownMenuItem object.
        /// </summary>
        private void shiftCellsDownMenuItem_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                VisualEditor.InsertCellsAndShiftDown();
            }
            catch (Exception ex)
            {
                DemosTools.ShowErrorMessage(ex);
            }
        }

        /// <summary>
        /// Handles the Click event of InsertEntireRowMenuItem object.
        /// </summary>
        private void insertEntireRowMenuItem_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                VisualEditor.InsertEmptyRows();
            }
            catch (Exception ex)
            {
                DemosTools.ShowErrorMessage(ex);
            }
        }

        /// <summary>
        /// Handles the Click event of InsertEntireColumnMenuItem object.
        /// </summary>
        private void insertEntireColumnMenuItem_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                VisualEditor.InsertEmptyColumns();
            }
            catch (Exception ex)
            {
                DemosTools.ShowErrorMessage(ex);
            }
        }

        /// <summary>
        /// Handles the Click event of ShiftCellsLeftMenuItem object.
        /// </summary>
        private void shiftCellsLeftMenuItem_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                VisualEditor.RemoveCellsAndShiftLeft();
            }
            catch (Exception ex)
            {
                DemosTools.ShowErrorMessage(ex);
            }
        }

        /// <summary>
        /// Handles the Click event of ShiftCellsUpMenuItem object.
        /// </summary>
        private void shiftCellsUpMenuItem_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                VisualEditor.RemoveCellsAndShiftUp();
            }
            catch (Exception ex)
            {
                DemosTools.ShowErrorMessage(ex);
            }
        }

        /// <summary>
        /// Handles the Click event of DeleteEntireRowMenuItem object.
        /// </summary>
        private void deleteEntireRowMenuItem_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                VisualEditor.RemoveRows();
            }
            catch (Exception ex)
            {
                DemosTools.ShowErrorMessage(ex);
            }
        }

        /// <summary>
        /// Handles the Click event of DeleteEntireColumnMenuItem object.
        /// </summary>
        private void deleteEntireColumnMenuItem_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                VisualEditor.RemoveColumns();
            }
            catch (Exception ex)
            {
                DemosTools.ShowErrorMessage(ex);
            }
        }

        /// <summary>
        /// Handles the Click event of ClearContentsMenuItem object.
        /// </summary>
        private void clearContentsMenuItem_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                VisualEditor.ClearCellsContents();
            }
            catch (Exception ex)
            {
                DemosTools.ShowErrorMessage(ex);
            }
        }

        /// <summary>
        /// Handles the Click event of InsertCommentMenuItem object.
        /// </summary>
        private void insertCommentMenuItem_Click(object sender, RoutedEventArgs e)
        {
            commentsPanel.NewComment();
        }

        /// <summary>
        /// Handles the Click event of EditCellCommentMenuItem object.
        /// </summary>
        private void editCellCommentMenuItem_Click(object sender, RoutedEventArgs e)
        {
            commentsPanel.EditComment();
        }

        /// <summary>
        /// Handles the Click event of DeleteCellCommentMenuItem object.
        /// </summary>
        private void deleteCellCommentMenuItem_Click(object sender, RoutedEventArgs e)
        {
            VisualEditor.RemoveComments();
        }

        /// <summary>
        /// Handles the Click event of ShowHideCommentMenuItem object.
        /// </summary>
        private void showHideCommentMenuItem_Click(object sender, RoutedEventArgs e)
        {
            VisualEditor.SetCommentIsVisible(!VisualEditor.FocusedCellComment.IsVisible);
        }

        /// <summary>
        /// Handles the Click event of FormatCellsMenuItem object.
        /// </summary>
        private void formatCellsMenuItem_Click(object sender, RoutedEventArgs e)
        {
            CellsStyleWindow.ShowNumberFormatDialog(VisualEditor);
        }

        /// <summary>
        /// Handles the Click event of DefineNameMenuItem object.
        /// </summary>
        private void defineNameMenuItem_Click(object sender, RoutedEventArgs e)
        {
            // get value for defined name
            string value = VisualEditor.GetFixedSelectedCells().ToString(VisualEditor.Document.Defaults.FormattingProperties);

            // create dialog that allows to add new defined name
            EditDefinedNameWindow dlg = new EditDefinedNameWindow(VisualEditor, value);
            dlg.WindowStartupLocation = WindowStartupLocation.CenterOwner;
            dlg.Owner = this;
            // show the dialog
            dlg.ShowDialog();
        }

        /// <summary>
        /// Handles the Click event of LinkMenuItem object.
        /// </summary>
        private void linkMenuItem_Click(object sender, RoutedEventArgs e)
        {
            if (VisualEditor.FocusedHyperlink != null && VisualEditor.SelectionContainsSingleCell)
                EditHyperlinkWindow.ShowDialog(VisualEditor, true);
            else
                EditHyperlinkWindow.ShowDialog(VisualEditor, false);
        }

        /// <summary>
        /// Handles the Click event of RemoveLinkMenuItem object.
        /// </summary>
        private void removeLinkMenuItem_Click(object sender, RoutedEventArgs e)
        {
            VisualEditor.RemoveHyperlinks();
        }

        /// <summary>
        /// Handles the Click event of ColumnWidthMenuItem object.
        /// </summary>
        private void columnWidthMenuItem_Click(object sender, RoutedEventArgs e)
        {
            cellsEditorPanel1.SetColumnWidth();
        }

        /// <summary>
        /// Handles the Click event of RowHeightMenuItem object.
        /// </summary>
        private void rowHeightMenuItem_Click(object sender, RoutedEventArgs e)
        {
            cellsEditorPanel1.SetRowHeight();
        }

        /// <summary>
        /// Handles the Click event of HideColumnsMenuItem object.
        /// </summary>
        private void hideColumnsMenuItem_Click(object sender, RoutedEventArgs e)
        {
            VisualEditor.HideColumns();
        }

        /// <summary>
        /// Handles the Click event of HideRowsMenuItem object.
        /// </summary>
        private void hideRowsMenuItem_Click(object sender, RoutedEventArgs e)
        {
            VisualEditor.HideRows();
        }

        /// <summary>
        /// Handles the Click event of UnhideColumnsMenuItem object.
        /// </summary>
        private void unhideColumnsMenuItem_Click(object sender, RoutedEventArgs e)
        {
            VisualEditor.ShowColumns();
        }

        /// <summary>
        /// Handles the Click event of UnhideRowsMenuItem object.
        /// </summary>
        private void unhideRowsMenuItem_Click(object sender, RoutedEventArgs e)
        {
            VisualEditor.ShowRows();
        }

        #endregion

        #endregion

    }
}

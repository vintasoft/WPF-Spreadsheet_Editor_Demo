using System;
using System.Globalization;
using System.IO;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

using Microsoft.Win32;

using Vintasoft.Imaging;
using Vintasoft.Imaging.Codecs.Decoders;
using Vintasoft.Imaging.Office;
using Vintasoft.Imaging.Office.OpenXml;
using Vintasoft.Imaging.Office.Spreadsheet;
using Vintasoft.Imaging.Office.Spreadsheet.Document;
using Vintasoft.Imaging.Office.Spreadsheet.UI;
using Vintasoft.Imaging.Office.Spreadsheet.Wpf.UI.Controls;
using Vintasoft.Imaging.Office.Wpf.UI;
using Vintasoft.Imaging.Wpf;
using Vintasoft.Imaging.Wpf.Print;

using WpfDemosCommonCode;
using WpfDemosCommonCode.CustomControls;
using WpfDemosCommonCode.Imaging;
using WpfDemosCommonCode.Imaging.Codecs;

namespace WpfSpreadsheetEditorDemo
{
    /// <summary>
    /// Main window.
    /// </summary>
    public partial class MainWindow : Window
    {

        #region Constants

        /// <summary>
        /// The name of chart templates resource.
        /// </summary>
        const string ChartTemplatesResourceName = "ChartSource.xlsx";

        #endregion



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


        /// <summary> 
        /// The document converter.
        /// </summary>
        DocumentConverter _documentConverter;

        /// <summary> 
        /// The layout settings manager.
        /// </summary>
        ImageCollectionXlsxLayoutSettingsManager _layoutSettingsManager;

        /// <summary> 
        /// A value indicating whether the layout settings are initialized.
        /// </summary>
        bool _isLayoutSettingsInitialized = false;

        /// <summary> 
        /// The export file dialog.
        /// </summary>
        SaveFileDialog _exportFileDialog = new SaveFileDialog();

        /// <summary> 
        /// The open worksheet file dialog.
        /// </summary>
        OpenFileDialog _openWorksheetFileDialog = new OpenFileDialog();

        /// <summary> 
        /// The save worksheet file dialog.
        /// </summary>
        SaveFileDialog _saveWorksheetFileDialog = new SaveFileDialog();

        /// <summary> 
        /// The print manager.
        /// </summary>
        WpfImagePrintManager _imagePrintManager = new WpfImagePrintManager();

        /// <summary>
        /// The "Help" panel.
        /// </summary>
        HelpPanel _helpPanel;

        /// <summary>
        /// The "Find and Replace" panel.
        /// </summary>
        FindReplacePanel _findReplacePanel;

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

            DocumentEditorControl.SpreadsheetEditorControl.PreviewMouseDoubleClick += spreadsheetEditorControl1_PreviewMouseDoubleClick;

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
            VisualEditor.FocusedCellsChanged += VisualEditor_FocusedCellsChanged;
            VisualEditor.ContextMenuOpen += VisualEditor_ContextMenuOpen;
            VisualEditor.ChartTemplatesRequest += VisualEditor_ChartTemplatesRequest;
            VisualEditor.DocumentSavingStarted += VisualEditor_DocumentSavingStarted;
            VisualEditor.Editor = null;
            SetStatus("");

            // init spreadsheet editor context menus
            _cellsContextMenu = FindResource("spreadsheetEditorContextMenu") as ContextMenu;
            _drawingContextMenu = FindResource("drawingContextMenu") as ContextMenu;
            _drawingSetImageMenuItem = (MenuItem)_drawingContextMenu.Items[0];
            _drawingRemoveLinkMenuItem = (MenuItem)_drawingContextMenu.Items[3];

            DocumentEditorControl.SpreadsheetEditorControl.ContextMenu = _cellsContextMenu;


            _imagePrintManager.PrintScaleMode = Vintasoft.Imaging.Print.PrintScaleMode.BestFit;
            _documentConverter = new DocumentConverter();

            _layoutSettingsManager = new ImageCollectionXlsxLayoutSettingsManager(_documentConverter.Images);
            XlsxDocumentLayoutSettings layoutSettings = new XlsxDocumentLayoutSettings();
            layoutSettings.PageLayoutSettingsType = XlsxPageLayoutSettingsType.UseWorksheetWidth;
            _layoutSettingsManager.LayoutSettings = layoutSettings;


            _openWorksheetFileDialog.Filter = "XLSX files|*.xlsx|XLS files|*.xls|TSV files|*.tsv;*.tab|CSV files|*.csv|ODS Files|*.ods|All supported Workbooks|*.xlsx;*.xls;*.tsv;*.tab;*.csv;*.ods";
            _openWorksheetFileDialog.FilterIndex = 6;

            DemosTools.SetTestXlsxFolder(_openWorksheetFileDialog);

            _saveWorksheetFileDialog.Filter = "XLSX files|*.xlsx";

            CodecsFileFilters.SetFilters(_exportFileDialog, false);

            _exportFileDialog.Filter += "|TSV files|*.tsv|CSV files|*.csv";
            // set default filter index to PDF
            string[] filters = _exportFileDialog.Filter.Split('|');
            for (int i = 1; i < filters.Length; i++)
            {
                if (filters[i].ToUpperInvariant().Contains("PDF"))
                    _exportFileDialog.FilterIndex = i / 2 + 1;
            }


            AddCustomPanels();

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
                return DocumentEditorControl.VisualEditor;
            }
        }

        #endregion



        #region Methods

        #region UI

        /// <summary>
        /// Add custom panels to the <see cref="MainMenuPanel"/>.
        /// </summary>
        private void AddCustomPanels()
        {
            // create "Find and Replace" panel
            _findReplacePanel = new FindReplacePanel();
            // add "Find and Replace" panel to the main menu panel
            DocumentEditorControl.MainMenuPanel.AddTabItem("Edit", _findReplacePanel);


            // add "Help" tab to the main menu panel
            DocumentEditorControl.MainMenuPanel.AddTab("Help", "Help");

            // create "Help" panel
            _helpPanel = new HelpPanel();
            // add "Help" panel to the main menu panel
            DocumentEditorControl.MainMenuPanel.AddTabItem("Help", _helpPanel);
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

        /// <summary>
        /// Returns the control with specified name.
        /// </summary>
        /// <param name="currentControl">The control from which search must be started.</param> 
        /// <param name="controlName">The name of control to search.</param>
        private Control FindControl(DependencyObject currentControl, string controlName)
        {
            foreach (object child in LogicalTreeHelper.GetChildren(currentControl))
            {
                Control control = child as Control;
                if (control != null)
                {
                    Control foundedControl;
                    if (control.Name == controlName)
                        foundedControl = control;
                    else
                        foundedControl = FindControl(control, controlName);

                    if (foundedControl != null)
                        return foundedControl;
                }
            }

            return null;
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

        /// <summary>
        /// Updates the User Interface.
        /// </summary>
        private void UpdateUI()
        {
            Title = "VintaSoft WPF Spreadsheet Editor Demo v" + ImagingGlobalSettings.ProductVersion;
            if (!string.IsNullOrEmpty(DocumentEditorControl.MainMenuPanel.Filename))
                Title += " - " + Path.GetFileName(DocumentEditorControl.MainMenuPanel.Filename);
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
                    DocumentEditorControl.MainMenuPanel.OpenDocument();
                    e.Handled = true;
                }

                // Ctrl+N
                if (isControlKeyPressed && e.Key == Key.N)
                {
                    DocumentEditorControl.MainMenuPanel.NewDocument();
                    e.Handled = true;
                }

                // Ctrl+S
                if (isControlKeyPressed && e.Key == Key.S)
                {
                    DocumentEditorControl.MainMenuPanel.SaveDocumentChanges();
                    e.Handled = true;
                }

                // Ctrl+Shift+S
                if (isControlKeyPressed && isShiftKeyPressed && e.Key == Key.S)
                {
                    if (DocumentEditorControl.MainMenuPanel.SaveDocumentAs())
                        e.Handled = true;
                }

                // Ctrl+P
                if (isControlKeyPressed && e.Key == Key.P)
                {
                    if (PrintDocument())
                        e.Handled = true;
                }

                // Ctrl+F
                if (isControlKeyPressed && e.Key == Key.F)
                {
                    if (VisualEditor.FocusedWorksheet != null)
                        _findReplacePanel.ShowFindDialog();
                }

                // Ctrl+H
                if (isControlKeyPressed && e.Key == Key.H)
                {
                    if (VisualEditor.FocusedWorksheet != null)
                        _findReplacePanel.ShowReplaceDialog();
                }

                // F1
                if (e.Key == Key.F1)
                {
                    _helpPanel.ShowAboutDialog();
                    e.Handled = true;
                }
            }
            catch (Exception ex)
            {
                DemosTools.ShowErrorMessage(ex);
                e.Handled = true;
            }
        }

        /// <summary>
        /// Handles the Closing event of Window object.
        /// </summary>
        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (!DocumentEditorControl.MainMenuPanel.CheckChanges())
                e.Cancel = true;
        }

        #endregion


        #region Spreasheet editor control events

        /// <summary>
        /// Handles the PreviewMouseDoubleClick event of spreadsheetEditorControl1 object.
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

        #endregion


        #region Spreasheet visual editor events

        /// <summary>
        /// Handles the HoveredDrawingChanged event of VisualEditor object.
        /// </summary>
        private void VisualEditor_HoveredDrawingChanged(object sender, PropertyChangedEventArgs<SheetDrawing> e)
        {
            if (e.NewValue != null)
                SetStatus(e.NewValue.Name);
            else
                SetStatus("");
        }

        /// <summary>
        /// Handles the HoveredCellChanged event of VisualEditor object.
        /// </summary>
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

        /// <summary>
        /// Handles the CellCommentClick event of VisualEditor object.
        /// </summary>
        private void VisualEditor_CellCommentClick(object sender, SheetCellMouseEventArgs e)
        {
            VisualEditor.SetCommentIsVisible(!VisualEditor.FocusedCellComment.IsVisible);
            e.Handled = true;
        }

        /// <summary>
        /// Handles the UriOpen event of VisualEditor object.
        /// </summary>
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

        /// <summary>
        /// Handles the HoveredHyperlinkChanged event of VisualEditor object.
        /// </summary>
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
                    hyperlink = DocumentEditorControl.VisualEditor.GetHyperlinkByDefinedName(hyperlink.Name);
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

        /// <summary>
        /// Handles the CellErrorClick event of VisualEditor object.
        /// </summary>
        private void VisualEditor_CellErrorClick(object sender, SheetCellMouseEventArgs e)
        {
            Worksheet worksheet = DocumentEditorControl.VisualEditor.FocusedWorksheet;
            SheetCell cell = worksheet.FindCell(e.Cell);
            string errorMessage = GetErrorMessage(cell.ErrorType);
            DemosTools.ShowWarningMessage("Error: " + cell.Value, errorMessage);
            e.Handled = true;
        }

        /// <summary>
        /// Handles the InvalidCellReferences event of VisualEditor object.
        /// </summary>
        private void VisualEditor_InvalidCellReferences(object sender, CellReferencesEventArgs e)
        {
            MessageBox.Show("Reference is not valid.", "Error", MessageBoxButton.OK, MessageBoxImage.Exclamation);
        }

        /// <summary>
        /// Handles the FocusedCellChanged event of VisualEditor object.
        /// </summary>
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

        /// <summary>
        /// Handles the ContextMenuOpen event of VisualEditor object.
        /// </summary>
        private void VisualEditor_ContextMenuOpen(object sender, Vintasoft.Imaging.UI.VintasoftControlMouseEventArgs e)
        {
            // if context menu for focused cell should be shown
            if (VisualEditor.FocusedComment != null)
            {
                DocumentEditorControl.SpreadsheetEditorControl.ContextMenu = _commentContextMenu;
            }
            // if context menu for focused drawing should be shown
            else if (VisualEditor.FocusedDrawing != null)
            {
                _drawingSetImageMenuItem.IsEnabled = VisualEditor.FocusedDrawing.Type == DrawingType.Picture;
                _drawingRemoveLinkMenuItem.IsEnabled = VisualEditor.FocusedDrawing.Hyperlink != null;
                DocumentEditorControl.SpreadsheetEditorControl.ContextMenu = _drawingContextMenu;
            }
            // if context menu for focused cells should be shown
            else
            {
                ContextMenu contextMenu = _cellsContextMenu;

                // determine if selection contains whole columns
                bool isCoverColumns = VisualEditor.FocusedCells.IsCoverColumns;
                // determine if selection contains whole rows
                bool isCoverRows = VisualEditor.FocusedCells.IsCoverRows;
                // determine if context menu was opened for "select all" button (top left corner of the grid)
                bool isSelectAllHovered = VisualEditor.HoveredCell.RowIndex < 0 && VisualEditor.HoveredCell.ColumnIndex < 0;

                // determine if all cells selected or selection does not have whole rows or columns
                bool addCellsMenuItems = isSelectAllHovered || !isCoverColumns && !isCoverRows;
                // determine if menu was opened on column header or selection contains whole columns, but not whole rows
                bool addColumnsMenuItems = !addCellsMenuItems && (VisualEditor.HoveredCell.RowIndex < 0 || isCoverColumns && !isCoverRows);
                // determine if none of the previous conditions were met and selection contains whole rows
                bool addRowsMenuItems = !addCellsMenuItems && !addColumnsMenuItems && isCoverRows;

                // add cells menu items
                SetVisibility(FindControl(contextMenu, "insertCellsMenuItem"), addCellsMenuItems);
                SetVisibility(FindControl(contextMenu, "deleteCellsMenuItem"), addCellsMenuItems);
                SetVisibility(FindControl(contextMenu, "defineNameMenuItem"), addCellsMenuItems);
                SetVisibility(FindControl(contextMenu, "linkMenuItem"), addCellsMenuItems);
                SetVisibility(FindControl(contextMenu, "removeLinkMenuItem"), addCellsMenuItems);

                // focused cell comment
                bool cellHasComment = VisualEditor.FocusedCellComment != null;
                bool isSingleCellSelection = VisualEditor.SelectionContainsSingleCell;
                // add cells comment menu items
                SetVisibility(FindControl(contextMenu, "insertCommentMenuItem"), addCellsMenuItems && !cellHasComment && isSingleCellSelection);
                SetVisibility(FindControl(contextMenu, "editCellCommentMenuItem"), addCellsMenuItems && cellHasComment && isSingleCellSelection);
                SetVisibility(FindControl(contextMenu, "showHideCommentMenuItem"), addCellsMenuItems && cellHasComment && isSingleCellSelection);
                SetVisibility(FindControl(contextMenu, "deleteCellCommentMenuItem"), addCellsMenuItems && VisualEditor.SelectedCellsHasComments);
                SetVisibility(FindControl(contextMenu, "commentSectionSeparator"), addCellsMenuItems && (isSingleCellSelection || VisualEditor.SelectedCellsHasComments));

                // add columns menu items
                SetVisibility(FindControl(contextMenu, "columnWidthMenuItem"), addColumnsMenuItems);
                SetVisibility(FindControl(contextMenu, "insertColumnsMenuItem"), addColumnsMenuItems);
                SetVisibility(FindControl(contextMenu, "deleteColumnsMenuItem"), addColumnsMenuItems);
                SetVisibility(FindControl(contextMenu, "hideColumnsMenuItem"), addColumnsMenuItems);
                SetVisibility(FindControl(contextMenu, "unhideColumnsMenuItem"), addColumnsMenuItems);

                // add rows menu items
                SetVisibility(FindControl(contextMenu, "rowHeightMenuItem"), addRowsMenuItems);
                SetVisibility(FindControl(contextMenu, "insertRowsMenuItem"), addRowsMenuItems);
                SetVisibility(FindControl(contextMenu, "deleteRowsMenuItem"), addRowsMenuItems);
                SetVisibility(FindControl(contextMenu, "hideRowsMenuItem"), addRowsMenuItems);
                SetVisibility(FindControl(contextMenu, "unhideRowsMenuItem"), addRowsMenuItems);

                DocumentEditorControl.SpreadsheetEditorControl.ContextMenu = _cellsContextMenu;
            }
        }

        /// <summary>
        /// Handles the FocusedCellsChanged event of VisualEditor object.
        /// </summary>
        private void VisualEditor_FocusedCellsChanged(object sender, PropertyChangedEventArgs<CellReferences> e)
        {
            if (e.NewValue != null)
            {
                if (VisualEditor.IsFocusedCellsChanging)
                    SetStatus(e.NewValue.ToString());
                else if (Equals(e.NewValue, e.OldValue))
                    SetStatus("");
            }
            else
            {
                SetStatus("");
            }
        }

        /// <summary>
        /// Handles the SynchronizationFinished event of VisualEditor object.
        /// </summary>
        private void VisualEditor_SynchronizationFinished(object sender, EventArgs e)
        {
            SetStatus("");
        }

        /// <summary>
        /// Handles the SynchronizationStarted event of VisualEditor object.
        /// </summary>
        private void VisualEditor_SynchronizationStarted(object sender, EventArgs e)
        {
            if (VisualEditor.IsInitialized)
                SetStatus("Processing...");
            else
                SetStatus("Loading...");
        }

        /// <summary>
        /// Handles the SynchronizationException event of VisualEditor object.
        /// </summary>
        private void VisualEditor_SynchronizationException(object sender, Vintasoft.Imaging.ExceptionEventArgs e)
        {
            DemosTools.ShowErrorMessage(e.Exception);
            DocumentEditorControl.MainMenuPanel.CloseDocument(false);
        }

        /// <summary>
        /// Handles the ChartTemplatesRequest event of VisualEditor object.
        /// </summary>
        private void VisualEditor_ChartTemplatesRequest(object sender, StreamRequestEventArgs e)
        {
            e.Stream = DemosResourcesManager.GetResourceAsStream(ChartTemplatesResourceName);
        }

        /// <summary>
        /// Handles the DocumentSavingStarted event of VisualEditor object.
        /// </summary>
        private void VisualEditor_DocumentSavingStarted(object sender, EventArgs e)
        {
            SpreadsheetEditor editor = VisualEditor.Editor;
            DocumentInformation info = new DocumentInformation(editor.DocumentInformation);

            // set information about user, who last modified this document
            info.LastModifiedBy = Environment.UserName;

            // set the modified date
            info.ModifiedDate = DateTime.Now.ToString(CultureInfo.InvariantCulture);

            // set the document information
            editor.DocumentInformation = info;
        }

        #endregion


        #region Comment context menu

        /// <summary>
        /// Handles the Click event of editCommentMenuItem object.
        /// </summary>
        private void editCommentMenuItem_Click(object sender, RoutedEventArgs e)
        {
            EditComment();
        }

        /// <summary>
        /// Handles the Click event of deleteCommentMenuItem object.
        /// </summary>
        private void deleteCommentMenuItem_Click(object sender, RoutedEventArgs e)
        {
            VisualEditor.RemoveFocusedComment();
        }

        /// <summary>
        /// Handles the Click event of hideCommentMenuItem object.
        /// </summary>
        private void hideCommentMenuItem_Click(object sender, RoutedEventArgs e)
        {
            VisualEditor.SetCommentIsVisible(false);
            VisualEditor.FocusedComment = null;
        }

        #endregion


        #region Drawing context menu

        /// <summary>
        /// Handles the Click event of drawingSetImageMenuItem object.
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
            {
                using (Stream stream = dialog.OpenFile())
                    VisualEditor.SetDrawingPicture(new ImageData(stream));
            }
        }

        /// <summary>
        /// Handles the Click event of deleteDrawingMenuItem object.
        /// </summary>
        private void deleteDrawingMenuItem_Click(object sender, RoutedEventArgs e)
        {
            VisualEditor.RemoveFocusedDrawing();
        }

        /// <summary>
        /// Handles the Click event of drawingLinkMenuItem object.
        /// </summary>
        private void drawingLinkMenuItem_Click(object sender, RoutedEventArgs e)
        {
            EditHyperlinkWindow.ShowDialog(VisualEditor, VisualEditor.FocusedDrawing.Hyperlink != null);
        }

        /// <summary>
        /// Handles the Click event of drawingRemoveLinkMenuItem object.
        /// </summary>
        private void drawingRemoveLinkMenuItem_Click(object sender, RoutedEventArgs e)
        {
            VisualEditor.RemoveHyperlinks();
        }

        /// <summary>
        /// Handles the Click event of drawingPropertiesMenuItem object.
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
        /// Handles the Click event of copyMenuItem object.
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
        /// Handles the Click event of cutMenuItem object.
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
        /// Handles the Click event of pasteMenuItem object.
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
        /// Handles the Click event of insertColumnsMenuItem object.
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
        /// Handles the Click event of insertRowsMenuItem object.
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
        /// Handles the Click event of deleteColumnsMenuItem object.
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
        /// Handles the Click event of deleteRowsMenuItem object.
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
        /// Handles the Click event of shiftCellsRightMenuItem object.
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
        /// Handles the Click event of shiftCellsDownMenuItem object.
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
        /// Handles the Click event of insertEntireRowMenuItem object.
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
        /// Handles the Click event of insertEntireColumnMenuItem object.
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
        /// Handles the Click event of shiftCellsLeftMenuItem object.
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
        /// Handles the Click event of shiftCellsUpMenuItem object.
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
        /// Handles the Click event of deleteEntireRowMenuItem object.
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
        /// Handles the Click event of deleteEntireColumnMenuItem object.
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
        /// Handles the Click event of clearContentsMenuItem object.
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
        /// Handles the Click event of insertCommentMenuItem object.
        /// </summary>
        private void insertCommentMenuItem_Click(object sender, RoutedEventArgs e)
        {
            NewComment();
        }

        /// <summary>
        /// Handles the Click event of editCellCommentMenuItem object.
        /// </summary>
        private void editCellCommentMenuItem_Click(object sender, RoutedEventArgs e)
        {
            EditComment();
        }

        /// <summary>
        /// Handles the Click event of deleteCellCommentMenuItem object.
        /// </summary>
        private void deleteCellCommentMenuItem_Click(object sender, RoutedEventArgs e)
        {
            VisualEditor.RemoveComments();
        }

        /// <summary>
        /// Handles the Click event of showHideCommentMenuItem object.
        /// </summary>
        private void showHideCommentMenuItem_Click(object sender, RoutedEventArgs e)
        {
            VisualEditor.SetCommentIsVisible(!VisualEditor.FocusedCellComment.IsVisible);
        }

        /// <summary>
        /// Handles the Click event of formatCellsMenuItem object.
        /// </summary>
        private void formatCellsMenuItem_Click(object sender, RoutedEventArgs e)
        {
            CellsStyleWindow.ShowNumberFormatDialog(VisualEditor);
        }

        /// <summary>
        /// Handles the Click event of defineNameMenuItem object.
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
        /// Handles the Click event of linkMenuItem object.
        /// </summary>
        private void linkMenuItem_Click(object sender, RoutedEventArgs e)
        {
            if (VisualEditor.FocusedHyperlink != null && VisualEditor.SelectionContainsSingleCell)
                EditHyperlinkWindow.ShowDialog(VisualEditor, true);
            else
                EditHyperlinkWindow.ShowDialog(VisualEditor, false);
        }

        /// <summary>
        /// Handles the Click event of removeLinkMenuItem object.
        /// </summary>
        private void removeLinkMenuItem_Click(object sender, RoutedEventArgs e)
        {
            VisualEditor.RemoveHyperlinks();
        }

        /// <summary>
        /// Handles the Click event of columnWidthMenuItem object.
        /// </summary>
        private void columnWidthMenuItem_Click(object sender, RoutedEventArgs e)
        {
            SetColumnWidth();
        }

        /// <summary>
        /// Handles the Click event of rowHeightMenuItem object.
        /// </summary>
        private void rowHeightMenuItem_Click(object sender, RoutedEventArgs e)
        {
            SetRowHeight();
        }

        /// <summary>
        /// Handles the Click event of hideColumnsMenuItem object.
        /// </summary>
        private void hideColumnsMenuItem_Click(object sender, RoutedEventArgs e)
        {
            VisualEditor.HideColumns();
        }

        /// <summary>
        /// Handles the Click event of hideRowsMenuItem object.
        /// </summary>
        private void hideRowsMenuItem_Click(object sender, RoutedEventArgs e)
        {
            VisualEditor.HideRows();
        }

        /// <summary>
        /// Handles the Click event of unhideColumnsMenuItem object.
        /// </summary>
        private void unhideColumnsMenuItem_Click(object sender, RoutedEventArgs e)
        {
            VisualEditor.ShowColumns();
        }

        /// <summary>
        /// Handles the Click event of unhideRowsMenuItem object.
        /// </summary>
        private void unhideRowsMenuItem_Click(object sender, RoutedEventArgs e)
        {
            VisualEditor.ShowRows();
        }

        #endregion     


        #region Panel events

        #region SpreadsheetVisualEditorPanel

        /// <summary>
        /// Handles the VisualEditorError event of SpreadsheetVisualEditorPanel object.
        /// </summary>
        private void SpreadsheetVisualEditorPanel_VisualEditorError(object sender, ExceptionEventArgs e)
        {
            DemosTools.ShowErrorMessage(e.Exception);
        }

        #endregion


        #region FilePanel  

        /// <summary>
        /// Handles the FilenameChanged event of FilePanel object.
        /// </summary>
        private void FilePanel_FilenameChanged(object sender, EventArgs e)
        {
            UpdateUI();
        }

        /// <summary>
        /// Handles the OpenFileRequest event of FilePanel object.
        /// </summary>
        private void FilePanel_OpenFileRequest(object sender, FilenameEventArgs e)
        {
            e.Filename = null;
            e.AllowAction = false;

            // show dialog for opening the XLSX file
            if (_openWorksheetFileDialog.ShowDialog() == true)
            {
                // get file path from open dialog
                string filename = _openWorksheetFileDialog.FileName;
                // if file is XLS file
                if (XlsxDecoder.IsXlsDocument(filename))
                {
                    if (MessageBox.Show("The loaded file is XLS file. To open XLS file application needs to convert XLS file to the XLSX file. Do you want to create XLSX file from XLS file?", "Convert XLS to XLSX", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.No)
                        return;

                    // create path to an XLSX file
                    filename = Path.Combine(Path.GetDirectoryName(filename), Path.GetFileNameWithoutExtension(filename) + ".xlsx");
                    // set file to the save dialog
                    _saveWorksheetFileDialog.FileName = filename;
                    // show the save dialog
                    if (_saveWorksheetFileDialog.ShowDialog() != true)
                        return;
                    // get file path from save dialog
                    filename = _saveWorksheetFileDialog.FileName;
                    // convert XLS file to the XLSX file
                    OpenXmlDocumentConverter.ConvertXlsToXlsx(_openWorksheetFileDialog.FileName, filename);
                }
                // if file is CSV file
                else if (XlsxDecoder.IsCsvFile(filename))
                {
                    if (MessageBox.Show("The loaded file is CSV file. To open CSV file application needs to convert CSV file to the XLSX file. Do you want to create XLSX file from CSV file?", "Convert CSV to XLSX", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.No)
                        return;

                    // create path to an XLSX file
                    filename = Path.Combine(Path.GetDirectoryName(filename), Path.GetFileNameWithoutExtension(filename) + ".xlsx");
                    // set file to the save dialog
                    _saveWorksheetFileDialog.FileName = filename;
                    // show the save dialog
                    if (_saveWorksheetFileDialog.ShowDialog() != true)
                        return;
                    // get file path from save dialog
                    filename = _saveWorksheetFileDialog.FileName;
                    // convert XLS file to the XLSX file
                    OpenXmlDocumentConverter.ConvertCsvToXlsx(_openWorksheetFileDialog.FileName, filename);
                }
                // if file is TSV file
                else if (XlsxDecoder.IsTsvFile(filename))
                {
                    if (MessageBox.Show("The loaded file is TSV file. To open TSV file application needs to convert TSV file to the XLSX file. Do you want to create XLSX file from TSV file?", "Convert TSV to XLSX", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.No)
                        return;

                    // create path to an XLSX file
                    filename = Path.Combine(Path.GetDirectoryName(filename), Path.GetFileNameWithoutExtension(filename) + ".xlsx");
                    // set file to the save dialog
                    _saveWorksheetFileDialog.FileName = filename;
                    // show the save dialog
                    if (_saveWorksheetFileDialog.ShowDialog() != true)
                        return;
                    // get file path from save dialog
                    filename = _saveWorksheetFileDialog.FileName;
                    // convert XLS file to the XLSX file
                    OpenXmlDocumentConverter.ConvertTsvToXlsx(_openWorksheetFileDialog.FileName, filename);
                }
                // if file is ODS file
                else if (XlsxDecoder.IsOdsDocument(filename))
                {
                    if (MessageBox.Show("The loaded file is ODS file. To open ODS file application needs to convert ODS file to the XLSX file. Do you want to create XLSX file from ODS file?", "Convert ODS to XLSX", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.No)
                        return;

                    // create path to an XLSX file
                    filename = Path.Combine(Path.GetDirectoryName(filename), Path.GetFileNameWithoutExtension(filename) + ".xlsx");
                    // set file to the save dialog
                    _saveWorksheetFileDialog.FileName = filename;
                    // show the save dialog
                    if (_saveWorksheetFileDialog.ShowDialog() != true)
                        return;
                    // get file path from save dialog
                    filename = _saveWorksheetFileDialog.FileName;
                    // convert ODS file to the XLSX file
                    OpenXmlDocumentConverter.ConvertOdsToXlsx(_openWorksheetFileDialog.FileName, filename);
                }
                // if file is encrypted XLSX file
                else if (Path.GetExtension(filename).ToUpperInvariant() == ".XLSX" && OfficeDocumentCryptography.IsSecuredOfficeDocument(filename))
                {
                    if (MessageBox.Show("The loaded file is secured XLSX file. To open secured file application needs to decrypt XLSX file. Do you want to create decrypted XLSX file?", "Decrypt XLSX", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.No)
                        return;

                    // set file to the save dialog
                    _saveWorksheetFileDialog.FileName = filename;
                    // show the save dialog
                    if (_saveWorksheetFileDialog.ShowDialog() != true)
                        return;

                    while (true)
                    {
                        DocumentPasswordWindow passwordWindow = new DocumentPasswordWindow();
                        passwordWindow.Filename = filename;

                        // enter password
                        if (passwordWindow.ShowDialog() != true)
                            return;

                        // try decrypt XLSX document
                        if (OfficeDocumentCryptography.TryDecryptOfficeDocument(filename, passwordWindow.Password, _saveWorksheetFileDialog.FileName))
                        {
                            break;
                        }
                        else
                        {
                            passwordWindow.ShowIncorrectPasswordMessage();
                        }
                    }

                    // get file path from save dialog
                    filename = _saveWorksheetFileDialog.FileName;
                }

                e.Filename = filename;
                e.AllowAction = true;
            }
        }

        /// <summary>
        /// Handles the ExportFile event of FilePanel object.
        /// </summary>
        private void FilePanel_ExportFile(object sender, EventArgs e)
        {
            _exportFileDialog.FileName = Path.GetFileNameWithoutExtension(DocumentEditorControl.MainMenuPanel.Filename);
            if (_exportFileDialog.ShowDialog() == true)
            {
                try
                {
                    string extension = Path.GetExtension(_exportFileDialog.FileName);

                    if (string.Equals(extension, ".tsv", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(extension, ".csv", StringComparison.OrdinalIgnoreCase))
                    {
                        if (MessageBox.Show(
                            "The selected file type does not support workbooks that contain multiple sheets.\r\nTo save only the active sheet, click OK.\r\nTo save all sheets, save them individually using a different file name for each, or choose a file type that supports multiple sheets.",
                            "Export document", MessageBoxButton.OKCancel, MessageBoxImage.Warning) == MessageBoxResult.OK)
                        {
                            // create a temporary stream
                            using (MemoryStream tempStream = new MemoryStream())
                            {
                                // save XLSX file to a temporary stream
                                VisualEditor.SaveDocumentTo(tempStream);

                                tempStream.Position = 0;

                                using (Stream stream = File.Create(_exportFileDialog.FileName))
                                {
                                    DocumentEnvironmentProperties environmentProperties = DocumentEnvironmentProperties.Default;
                                    environmentProperties.Culture = CultureInfo.CurrentCulture;
                                    environmentProperties.UICulture = CultureInfo.CurrentUICulture;

                                    if (string.Equals(extension, ".tsv", StringComparison.OrdinalIgnoreCase))
                                        OpenXmlDocumentConverter.ConvertXlsxToTsv(environmentProperties, tempStream, VisualEditor.FocusedWorksheetIndex, stream);
                                    else
                                        OpenXmlDocumentConverter.ConvertXlsxToCsv(environmentProperties, tempStream, VisualEditor.FocusedWorksheetIndex, stream, System.Text.Encoding.UTF8);
                                }
                            }
                        }
                    }
                    else
                    {
                        if (!InitLayoutSettings())
                            return;

                        // create a temporary stream
                        using (MemoryStream tempStream = new MemoryStream())
                        {
                            // save XLSX file to a temporary stream
                            VisualEditor.SaveDocumentTo(tempStream);

                            // add XLSX file to the image collection of document converter
                            _documentConverter.Images.Add(tempStream);

                            // create dialog that displays progress for document conversion process
                            ActionProgressWindow dlg = new ActionProgressWindow(ExportDocument, 1, "Export document");
                            // specify that dialog should be closed when conversion is finished
                            dlg.CloseAfterComplete = true;
                            // show dialog and run conversion process
                            dlg.RunAndShowDialog(Application.Current.MainWindow);

                            // clear image collection of document converter
                            _documentConverter.Images.ClearAndDisposeItems();
                        }
                    }
                }
                catch (Exception ex)
                {
                    DemosTools.ShowErrorMessage(ex);
                }
            }
        }

        /// <summary>
        /// Handles the PrintDocument event of FilePanel object.
        /// </summary>
        private void FilePanel_PrintDocument(object sender, EventArgs e)
        {
            PrintDocument();
        }

        /// <summary>
        /// Handles the ShowPrintLayoutSettings event of FilePanel object.
        /// </summary>
        private void FilePanel_ShowPrintLayoutSettings(object sender, EventArgs e)
        {
            if (_layoutSettingsManager.EditLayoutSettingsUseDialog(Application.Current.MainWindow))
                _isLayoutSettingsInitialized = true;
        }

        /// <summary>
        /// Handles the ShowPrintPageSettings event of FilePanel object.
        /// </summary>
        private void FilePanel_ShowPrintPageSettings(object sender, EventArgs e)
        {
            // show dialog with page setup setings

            PageSettingsWindow pageSettingsWindow = new PageSettingsWindow(
                _imagePrintManager, _imagePrintManager.PagePadding, _imagePrintManager.ImagePadding);
            pageSettingsWindow.WindowStartupLocation = WindowStartupLocation.CenterOwner;
            pageSettingsWindow.Owner = Application.Current.MainWindow;
            if (pageSettingsWindow.ShowDialog() == true)
            {
                _imagePrintManager.PagePadding = pageSettingsWindow.PagePadding;
                _imagePrintManager.ImagePadding = pageSettingsWindow.ImagePadding;
            }
        }

        /// <summary>
        /// Handles the ShowPrintPreview event of FilePanel object.
        /// </summary>
        private void FilePanel_ShowPrintPreview(object sender, EventArgs e)
        {
            try
            {
                // if layout settings are not initialized
                if (!_isLayoutSettingsInitialized)
                {
                    // set layout settings
                    if (_layoutSettingsManager.EditLayoutSettingsUseDialog(Application.Current.MainWindow))
                        _isLayoutSettingsInitialized = true;
                    else
                        return;
                }

                // create a temporary stream
                using (MemoryStream tempStream = new MemoryStream())
                {
                    // save XLSX file to a temporary stream
                    VisualEditor.SaveDocumentTo(tempStream);

                    // create a dialog that allows to preview and print XLSX document
                    PrintPreviewWindow window = new PrintPreviewWindow(tempStream);
                    window.WindowStartupLocation = WindowStartupLocation.CenterOwner;
                    window.Owner = Application.Current.MainWindow;
                    // show the dialog
                    window.ShowDialog();
                }
            }
            catch (Exception ex)
            {
                // show error message
                DemosTools.ShowErrorMessage(ex);
            }
        }

        /// <summary>
        /// Handles the SaveChangesRequest event of FilePanel object.
        /// </summary>
        private void FilePanel_SaveChangesRequest(object sender, ActionRequestEventArgs e)
        {
            MessageBoxResult result = MessageBox.Show("Current workbook is changed. Do you want to save changes?", "New Workbook", MessageBoxButton.YesNoCancel, MessageBoxImage.Question, MessageBoxResult.Cancel);
            switch (result)
            {
                case MessageBoxResult.Yes:
                    e.AllowAction = true;
                    break;

                case MessageBoxResult.No:
                    break;

                case MessageBoxResult.Cancel:
                    e.Cancel = true;
                    break;
            }
        }

        /// <summary>
        /// Handles the SaveAsRequest event of FilePanel object.
        /// </summary>
        private void FilePanel_SaveAsRequest(object sender, FilenameEventArgs e)
        {
            _saveWorksheetFileDialog.FileName = DocumentEditorControl.MainMenuPanel.Filename;
            if (_saveWorksheetFileDialog.ShowDialog() == true)
            {
                e.Filename = _saveWorksheetFileDialog.FileName;
                e.AllowAction = true;
            }
            else
            {
                e.Filename = null;
                e.AllowAction = false;
            }
        }

        /// <summary>
        /// Handles the ShowDocumentInfo event of FilePanel object.
        /// </summary>
        private void FilePanel_ShowDocumentInfo(object sender, EventArgs e)
        {
            DocumentInfoWindow.ShowDialog(VisualEditor);
        }

        /// <summary>
        /// Handles the ShowVisualEditorOptions event of FilePanel object.
        /// </summary>
        private void FilePanel_ShowVisualEditorOptions(object sender, EventArgs e)
        {
            OptionsWindow.ShowDialog(VisualEditor);
        }

        /// <summary>
        /// Prints the XLSX document.
        /// </summary>
        private bool PrintDocument()
        {
            if (VisualEditor.Document == null)
                return false;
            try
            {
                // if layout settings are not initialized
                if (!_isLayoutSettingsInitialized)
                {
                    // set layout settings
                    if (_layoutSettingsManager.EditLayoutSettingsUseDialog(Application.Current.MainWindow))
                        _isLayoutSettingsInitialized = true;
                    else
                        return true;
                }

                // create a temporary stream
                using (MemoryStream tempStream = new MemoryStream())
                {
                    // save XLSX file to a temporary stream
                    VisualEditor.SaveDocumentTo(tempStream);

                    using (ImageCollection images = new ImageCollection())
                    {
                        images.Add(tempStream);
                        try
                        {
                            _imagePrintManager.Images = images;
                            _imagePrintManager.PrintDialog.MinPage = 1;
                            _imagePrintManager.PrintDialog.MaxPage = (uint)images.Count;

                            if (_imagePrintManager.PrintDialog.ShowDialog() == true)
                                // print XLSX document
                                _imagePrintManager.Print("XLSX document");
                        }
                        finally
                        {
                            _imagePrintManager.Images = null;
                            images.ClearAndDisposeItems();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // show error message
                DemosTools.ShowErrorMessage(ex);
            }
            return true;
        }

        /// <summary>
        /// Exports the XLSX document.
        /// </summary>
        /// <param name="progressController">Progress controller.</param>
        private void ExportDocument(Vintasoft.Imaging.Utils.IActionProgressController progressController)
        {
            // set progress controller for document converter
            _documentConverter.ProgressController = progressController;

            // convert XLSX to the selected format            
            _documentConverter.Convert(_exportFileDialog.FileName);
        }

        /// <summary>
        /// Initialize the layout settings of XLSX document.
        /// </summary>
        private bool InitLayoutSettings()
        {
            // if layout settings are not initialized
            if (!_isLayoutSettingsInitialized)
            {
                // set layout settings
                if (_layoutSettingsManager.EditLayoutSettingsUseDialog(Application.Current.MainWindow))
                    _isLayoutSettingsInitialized = true;
                else
                    return false;
            }

            return true;
        }

        #endregion


        #region CellsEditorPanel

        /// <summary>
        /// Handles the SetRowHeight event of CellsEditorPanel object.
        /// </summary>
        private void CellsEditorPanel_SetRowHeight(object sender, EventArgs e)
        {
            SetRowHeight();
        }

        private void SetRowHeight()
        {
            // show value editor dialog
            NumberValueEditorWindow dlg = new NumberValueEditorWindow(VisualEditor, VisualEditor.RowsHeight, 0, CellsEditorPanel.MAX_ROW_HEIGHT, "Row height");
            dlg.WindowStartupLocation = WindowStartupLocation.CenterOwner;
            dlg.Owner = Application.Current.MainWindow;
            if (dlg.ShowDialog() == true)
            {
                // set height of focused rows
                VisualEditor.RowsHeight = dlg.Value;
            }
        }

        /// <summary>
        /// Handles the SetDefaultRowHeight event of CellsEditorPanel object.
        /// </summary>
        private void CellsEditorPanel_SetDefaultRowHeight(object sender, EventArgs e)
        {
            WorksheetFormat sheetFormat = VisualEditor.FocusedWorksheet.Format;
            // show value editor dialog
            NumberValueEditorWindow dlg = new NumberValueEditorWindow(VisualEditor, sheetFormat.RowHeight, 0, CellsEditorPanel.MAX_ROW_HEIGHT, "Default row height");
            dlg.WindowStartupLocation = WindowStartupLocation.CenterOwner;
            dlg.Owner = Application.Current.MainWindow;
            if (dlg.ShowDialog() == true)
            {
                // set visual format with new default row height
                VisualEditor.SetWorksheetFormat(new WorksheetFormat(sheetFormat.ColumnWidth, dlg.Value, sheetFormat.AutoHeight, sheetFormat.RowsHiddenByDefault));
            }
        }

        /// <summary>
        /// Handles the SetColumnWidth event of CellsEditorPanel object.
        /// </summary>
        private void CellsEditorPanel_SetColumnWidth(object sender, EventArgs e)
        {
            SetColumnWidth();
        }

        private void SetColumnWidth()
        {
            // show value editor dialog
            NumberValueEditorWindow dlg = new NumberValueEditorWindow(VisualEditor, VisualEditor.ColumnsWidth, 0, CellsEditorPanel.MAX_COLUMN_WIDTH, "Column width");
            dlg.WindowStartupLocation = WindowStartupLocation.CenterOwner;
            dlg.Owner = Application.Current.MainWindow;
            if (dlg.ShowDialog() == true)
            {
                // set width of focused columns
                VisualEditor.ColumnsWidth = dlg.Value;
            }
        }

        /// <summary>
        /// Handles the SetDefaultColumnWidth event of CellsEditorPanel object.
        /// </summary>
        private void CellsEditorPanel_SetDefaultColumnWidth(object sender, EventArgs e)
        {
            WorksheetFormat sheetFormat = VisualEditor.FocusedWorksheet.Format;
            // show value editor dialog
            NumberValueEditorWindow dlg = new NumberValueEditorWindow(VisualEditor, sheetFormat.ColumnWidth, 0, CellsEditorPanel.MAX_COLUMN_WIDTH, "Default column width");
            dlg.WindowStartupLocation = WindowStartupLocation.CenterOwner;
            dlg.Owner = Application.Current.MainWindow;

            if (dlg.ShowDialog() == true)
            {
                // set visual format with new default column width
                VisualEditor.SetWorksheetFormat(new WorksheetFormat(dlg.Value, sheetFormat.RowHeight, sheetFormat.AutoHeight, sheetFormat.RowsHiddenByDefault));
            }
        }

        /// <summary>
        /// Handles the AddChart event of CellsEditorPanel object.
        /// </summary>
        private void CellsEditorPanel_AddChart(object sender, EventArgs e)
        {
            try
            {
                InsertChartWindow chartWindow = new InsertChartWindow();
                chartWindow.SetChartDataSource(VisualEditor, ChartTemplatesResourceName);
                chartWindow.WindowStartupLocation = WindowStartupLocation.CenterOwner;
                chartWindow.Owner = Application.Current.MainWindow;
                chartWindow.ShowDialog();
            }
            catch (Exception ex)
            {
                DemosTools.ShowWarningMessage("Insert Chart", ex.Message);
            }
        }

        /// <summary>
        /// Handles the EditDrawing event of CellsEditorPanel object.
        /// </summary>
        private void CellsEditorPanel_EditDrawing(object sender, EventArgs e)
        {
            try
            {
                DrawingPropertiesWindow window = new DrawingPropertiesWindow(VisualEditor, VisualEditor.FocusedDrawing);
                window.WindowStartupLocation = WindowStartupLocation.CenterOwner;
                window.Owner = Application.Current.MainWindow;
                window.ShowDialog();
            }
            catch (Exception ex)
            {
                DemosTools.ShowErrorMessage(ex);
            }
        }

        /// <summary>
        /// Handles the ImageStreamRequest event of CellsEditorPanel object.
        /// </summary>
        private void CellsEditorPanel_ImageStreamRequest(object sender, StreamRequestEventArgs e)
        {
            // create dialog
            OpenFileDialog dialog = new OpenFileDialog();
            // specify that dialog should open folder with demo images
            DemosTools.SetTestFilesFolder(dialog);
            // set image filters
            CodecsFileFilters.SetFilters(dialog);

            // if image must be loaded
            if (dialog.ShowDialog() == true)
            {
                e.Stream = dialog.OpenFile();
                // return image stream
                return;
            }

            e.Stream = null;
        }

        /// <summary>
        /// Handles the AddHyperlink event of CellsEditorPanel object.
        /// </summary>
        private void CellsEditorPanel_AddHyperlink(object sender, EventArgs e)
        {
            EditHyperlinkWindow.ShowDialog(VisualEditor, false);
        }

        /// <summary>
        /// Handles the EditHyperlink event of CellsEditorPanel object.
        /// </summary>
        private void CellsEditorPanel_EditHyperlink(object sender, EventArgs e)
        {
            EditHyperlinkWindow.ShowDialog(VisualEditor, true);
        }

        #endregion


        #region CommentsPanel

        /// <summary>
        /// Handles the NewComment event of CommentsPanel object.
        /// </summary>
        private void CommentsPanel_NewComment(object sender, EventArgs e)
        {
            NewComment();
        }

        private void NewComment()
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
        /// Handles the EditComment event of CommentsPanel object.
        /// </summary>
        private void CommentsPanel_EditComment(object sender, EventArgs e)
        {
            EditComment();
        }

        private void EditComment()
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


        #region CopyPastePanel

        /// <summary>
        /// Handles the ShowCellPasteSpecial event of CopyPastePanel object.
        /// </summary>
        private void CopyPastePanel_ShowCellPasteSpecial(object sender, ActionRequestEventArgs e)
        {
            CellPasteSpecialWindow dlg = new CellPasteSpecialWindow(VisualEditor);
            dlg.WindowStartupLocation = WindowStartupLocation.CenterOwner;
            dlg.Owner = Application.Current.MainWindow;
            if (dlg.ShowDialog() == true)
            {
                e.AllowAction = true;
            }
        }

        #endregion 


        #region DefinedNamesPanel

        /// <summary>
        /// Handles the InsertDefinedNameInFormula event of DefinedNamesPanel object.
        /// </summary>
        private void DefinedNamesPanel_InsertDefinedNameInFormula(object sender, EventArgs e)
        {
            // get a list of defined names, which are defined on focused worksheet
            DefinedName[] definedNames = VisualEditor.GetFocusedWorksheetDefinedNames();

            // create dialog that allows to select defined name
            SelectDefinedNameWindow dlg = new SelectDefinedNameWindow(definedNames);
            dlg.WindowStartupLocation = WindowStartupLocation.CenterOwner;
            dlg.Owner = Application.Current.MainWindow;
            // show the dialog
            if (dlg.ShowDialog() == true)
            {
                // get selected defined name
                string selectedDefinedName = dlg.SelectedDefinedName;

                // insert formula in focused cell
                VisualEditor.InsertFormulaInFocusedCell(selectedDefinedName);
            }
        }

        /// <summary>
        /// Handles the AddDefinedName event of DefinedNamesPanel object.
        /// </summary>
        private void DefinedNamesPanel_AddDefinedName(object sender, EventArgs e)
        {
            // get value for defined name
            string value = VisualEditor.GetFixedSelectedCells().ToString(VisualEditor.Document.Defaults.FormattingProperties);

            // create dialog that allows to add new defined name
            EditDefinedNameWindow dlg = new EditDefinedNameWindow(VisualEditor, value);
            dlg.WindowStartupLocation = WindowStartupLocation.CenterOwner;
            dlg.Owner = Application.Current.MainWindow;
            // show the dialog
            dlg.ShowDialog();
        }

        /// <summary>
        /// Handles the ShowDefinedNamesManager event of DefinedNamesPanel object.
        /// </summary>
        private void DefinedNamesPanel_ShowDefinedNamesManager(object sender, EventArgs e)
        {
            DefinedNameManagerWindow dlg = new DefinedNameManagerWindow(VisualEditor);
            dlg.WindowStartupLocation = WindowStartupLocation.CenterOwner;
            dlg.Owner = Application.Current.MainWindow;
            dlg.ShowDialog();
        }

        #endregion


        #region FontPropertiesPanel

        /// <summary>
        /// Handles the ColorRequest event of FontPropertiesPanel object.
        /// </summary>
        private void FontPropertiesPanel_ColorRequest(object sender, ColorEventArgs e)
        {
            ColorPickerDialog colorDialog1 = new ColorPickerDialog();
            colorDialog1.WindowStartupLocation = WindowStartupLocation.CenterOwner;
            colorDialog1.Owner = Application.Current.MainWindow;
            colorDialog1.StartingColor = e.Color;
            colorDialog1.CanEditAlphaChannel = false;

            if (colorDialog1.ShowDialog() == true)
            {
                e.Color = colorDialog1.SelectedColor;
                e.AllowAction = true;
            }
        }

        /// <summary>
        /// Handles the FontProperties event of FontPropertiesPanel object.
        /// </summary>
        private void FontPropertiesPanel_FontProperties(object sender, EventArgs e)
        {
            CellsStyleWindow.ShowFontDialog(VisualEditor);
        }

        /// <summary>
        /// Handles the Borders event of FontPropertiesPanel object.
        /// </summary>
        private void FontPropertiesPanel_Borders(object sender, EventArgs e)
        {
            CellsStyleWindow.ShowBordersDialog(VisualEditor);
        }

        #endregion   


        #region FormulaPanel     

        /// <summary>
        /// Handles the EditCellFormulaError event of FormulaPanel object.
        /// </summary>
        private void FormulaPanel_EditCellFormulaError(object sender, ExceptionActionEventArgs e)
        {
            Exception exception = e.Exception;

            if (MessageBox.Show(exception.Message, "Formula Syntax Error", MessageBoxButton.OKCancel, MessageBoxImage.Warning) == MessageBoxResult.OK)
                e.AllowAction = true;
        }

        /// <summary>
        /// Handles the InsertFunction event of FormulaPanel object.
        /// </summary>
        private void FormulaPanel_InsertFunction(object sender, XlsxFunctionNameEventArgs e)
        {
            e.FunctionName = SelectFunctionWindow.SelectFunction(VisualEditor.Document);
            e.AllowAction = true;
        }

        #endregion    


        #region FunctionPanel    

        /// <summary>
        /// Handles the InsertFunction event of FunctionsPanel object.
        /// </summary>
        private void FunctionsPanel_InsertFunction(object sender, XlsxFunctionNameEventArgs e)
        {
            e.FunctionName = SelectFunctionWindow.SelectFunction(VisualEditor.Document);
            e.AllowAction = true;
        }

        #endregion    


        #region NavigationPanel    

        /// <summary>
        /// Handles the RemoveWorksheet event of NavigationPanel object.
        /// </summary>
        private void NavigationPanel_RemoveWorksheet(object sender, ActionRequestEventArgs e)
        {
            if (MessageBox.Show(string.Format("Do you want to remove worksheet '{0}'?", VisualEditor.FocusedWorksheet.Name),
                "Remove Worksheet", MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.Yes)
            {
                e.AllowAction = true;
            }
        }

        /// <summary>
        /// Handles the MoveWorksheet event of NavigationPanel object.
        /// </summary>
        private void NavigationPanel_MoveWorksheet(object sender, ActionRequestEventArgs e)
        {
            // show "Move worksheet" dialog
            MoveWorksheetWindow dlg = new MoveWorksheetWindow(VisualEditor);
            dlg.WindowStartupLocation = WindowStartupLocation.CenterOwner;
            dlg.Owner = Application.Current.MainWindow;
            if (dlg.ShowDialog() == true)
            {
                if (dlg.IsWorksheetOrderChanged)
                {
                    e.AllowAction = true;
                }
            }
        }

        /// <summary>
        /// Handles the RenameWorksheet event of NavigationPanel object.
        /// </summary>
        private void NavigationPanel_RenameWorksheet(object sender, ActionRequestEventArgs e)
        {
            // show "Rename worksheet" dialog
            RenameWorksheetWindow dlg = new RenameWorksheetWindow(VisualEditor);
            dlg.WindowStartupLocation = WindowStartupLocation.CenterOwner;
            dlg.Owner = Application.Current.MainWindow;
            if (dlg.ShowDialog() == true)
            {
                if (dlg.IsWorksheetNameChanged)
                {
                    e.AllowAction = true;
                }
            }
        }

        /// <summary>
        /// Handles the ColorRequest event of NavigationPanel object.
        /// </summary>
        private void NavigationPanel_ColorRequest(object sender, ColorEventArgs e)
        {
            // create "Color" dialog
            ColorPickerDialog colorDialog = new ColorPickerDialog();
            colorDialog.WindowStartupLocation = WindowStartupLocation.CenterOwner;
            colorDialog.Owner = Application.Current.MainWindow;
            colorDialog.CanEditAlphaChannel = false;
            colorDialog.StartingColor = WpfObjectConverter.Convert(VisualEditor.GridColor);

            if (colorDialog.ShowDialog() == true)
            {
                e.Color = colorDialog.SelectedColor;
                e.AllowAction = true;
            }
        }

        /// <summary>
        /// Handles the WorksheetFormat event of NavigationPanel object.
        /// </summary>
        private void NavigationPanel_WorksheetFormat(object sender, EventArgs e)
        {
            // show "Worksheet Format" dialog
            WorksheetFormatWindow.ShowDialog(VisualEditor);
        }

        #endregion   


        #region NumberFormatPanel    

        /// <summary>
        /// Handles the NumberFormatProperties event of NumberFormatPanel object.
        /// </summary>
        private void NumberFormatPanel_NumberFormatProperties(object sender, EventArgs e)
        {
            CellsStyleWindow.ShowNumberFormatDialog(VisualEditor);
        }

        #endregion  


        #region TextAlignmentPanel    

        /// <summary>
        /// Handles the AlignmentProperties event of TextAlignmentPanel object.
        /// </summary>
        private void TextAlignmentPanel_AlignmentProperties(object sender, EventArgs e)
        {
            CellsStyleWindow.ShowAlignmentDialog(VisualEditor);
        }

        #endregion

        #endregion


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

        #endregion



        #region Delegates

        /// <summary>
        /// SetStatus delegate.
        /// </summary>
        /// <param name="status">The status.</param>
        delegate void SetStatusDelegate(string status);

        #endregion

    }
}

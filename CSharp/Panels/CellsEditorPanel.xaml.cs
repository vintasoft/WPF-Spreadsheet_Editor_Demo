using System;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

using Microsoft.Win32;

using Vintasoft.Imaging;
using Vintasoft.Imaging.Office.Spreadsheet;
using Vintasoft.Imaging.Office.Spreadsheet.Document;
using Vintasoft.Imaging.Office.Spreadsheet.Wpf.UI;
using WpfDemosCommonCode;
using WpfDemosCommonCode.Imaging.Codecs;

namespace WpfSpreadsheetEditorDemo
{
    /// <summary>
    /// Provides a "Cells Editor" panel.
    /// </summary>
    public partial class CellsEditorPanel : SpreadsheetVisualEditorPanel
    {

        #region Constants

        /// <summary>
        /// Maximum column width in DIP.
        /// </summary>
        const double MAX_COLUMN_WIDTH = 1908;

        /// <summary>
        /// Maximum row height in DIP.
        /// </summary>
        const double MAX_ROW_HEIGHT = 545;

        /// <summary>
        /// The chart templates resource name.
        /// </summary>
        const string ChartTemplatesResourceName = "ChartSource.xlsx";

        /// <summary>
        /// The "Insert Rows" menu item.
        /// </summary>
        readonly MenuItem insertRowsMenuItem;

        /// <summary>
        /// The "Insert Columns" menu item.
        /// </summary>
        readonly MenuItem insertColumnsMenuItem;

        /// <summary>
        /// The "Set Picture..." menu item.
        /// </summary>
        readonly MenuItem setPictureMenuItem;

        /// <summary>
        /// The "Properties..." menu item.
        /// </summary>
        readonly MenuItem picturePropertiesMenuItem;

        /// <summary>
        /// The "Remove Picture" menu item.
        /// </summary>
        readonly MenuItem removePictureMenuItem;

        /// <summary>
        /// The "Edit Hyperlink..." menu item.
        /// </summary>
        readonly MenuItem editHyperlinkMenuItem;

        /// <summary>
        /// The "Remove Hyperlink" menu item.
        /// </summary>
        readonly MenuItem removeHyperlinkMenuItem;

        /// <summary>
        /// The "Add Chart..." menu item.
        /// </summary>
        readonly MenuItem addChartMenuItem;

        /// <summary>
        /// The "Properties..." menu item.
        /// </summary>
        readonly MenuItem chartPropertiesMenuItem;

        /// <summary>
        /// The "Remove Chart" menu item.
        /// </summary>
        readonly MenuItem removeChartMenuItem;

        /// <summary>
        /// The "Switch Rows/Columns" menu item.
        /// </summary>
        readonly MenuItem switchRowColumnMenuItem;

        /// <summary>
        /// The "Select Chart Values" menu item.
        /// </summary>
        readonly MenuItem selectChartValuesMenuItem;

        #endregion



        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="CellsEditorPanel"/> class.
        /// </summary>
        public CellsEditorPanel()
        {
            InitializeComponent();

            insertRowsMenuItem = (MenuItem)insertButton.Items[0];
            insertColumnsMenuItem = (MenuItem)insertButton.Items[1];

            addChartMenuItem = (MenuItem)chartButton.Items[0];
            removeChartMenuItem = (MenuItem)chartButton.Items[1];
            switchRowColumnMenuItem = (MenuItem)chartButton.Items[2];
            selectChartValuesMenuItem = (MenuItem)chartButton.Items[4];
            chartPropertiesMenuItem = (MenuItem)chartButton.Items[5];

            setPictureMenuItem = (MenuItem)pictureButton.Items[1];
            picturePropertiesMenuItem = (MenuItem)pictureButton.Items[2];
            removePictureMenuItem = (MenuItem)pictureButton.Items[3];

            editHyperlinkMenuItem = (MenuItem)hypelinkSplitButton.Items[1];
            removeHyperlinkMenuItem = (MenuItem)hypelinkSplitButton.Items[2];
        }

        #endregion



        #region Methods

        #region Common

        /// <summary>
        /// Raises the <see cref="E:SpreadsheetEditorChanged" /> event.
        /// </summary>
        /// <param name="args">The <see cref="PropertyChangedEventArgs{SpreadsheetEditorControl}"/> instance containing the event data.</param>
        protected override void OnSpreadsheetEditorChanged(PropertyChangedEventArgs<WpfSpreadsheetEditorControl> args)
        {
            base.OnSpreadsheetEditorChanged(args);

            if (args.OldValue != null)
            {
                args.OldValue.VisualEditor.FocusedCellsChanged -= VisualEditor_FocusedCellsChanged;
                args.OldValue.VisualEditor.FocusedDrawingChanged -= VisualEditor_FocusedDrawingChanged;
                args.OldValue.VisualEditor.ChartTemplatesRequest -= VisualEditor_ChartTemplatesRequest;
                args.OldValue.MouseDoubleClick -= SpreadsheetEditorControl_MouseDoubleClick;
            }

            if (args.NewValue != null)
            {
                args.NewValue.VisualEditor.FocusedCellsChanged += VisualEditor_FocusedCellsChanged;
                args.NewValue.VisualEditor.FocusedDrawingChanged += VisualEditor_FocusedDrawingChanged;
                args.NewValue.VisualEditor.ChartTemplatesRequest += VisualEditor_ChartTemplatesRequest;
                args.NewValue.MouseDoubleClick += SpreadsheetEditorControl_MouseDoubleClick;
            }

            UpdateUI();
        }

        private void VisualEditor_FocusedDrawingChanged(object sender, PropertyChangedEventArgs<SheetDrawing> e)
        {
            UpdateUI();
        }

        private void VisualEditor_FocusedCellsChanged(object sender, PropertyChangedEventArgs<CellReferences> e)
        {
            UpdateUI();
        }

        /// <summary>
        /// Updates the user interface.
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
                bool hasFocusedCells = VisualEditor.FocusedCells != null;
                insertButton.IsEnabled = hasFocusedCells;
                deleteButton.IsEnabled = hasFocusedCells;
                formatButton.IsEnabled = hasFocusedCells;
                mergeMenuButton.IsEnabled = hasFocusedCells;
                clearButton.IsEnabled = hasFocusedCells;
                fillButton.IsEnabled = hasFocusedCells;
                insertColumnsMenuItem.IsEnabled = VisualEditor.CanInsertEmptyColumns;
                insertRowsMenuItem.IsEnabled = VisualEditor.CanInsertEmptyRows;
                addChartMenuItem.IsEnabled = VisualEditor.FocusedCells != null;

                bool chartIsSelected = VisualEditor.FocusedDrawing != null && VisualEditor.FocusedDrawing.ChartProperties != null;
                pictureButton.IsEnabled = !chartIsSelected || VisualEditor.FocusedDrawing == null;
                chartButton.IsEnabled = chartIsSelected || VisualEditor.FocusedDrawing == null;
                removeChartMenuItem.IsEnabled = chartIsSelected;
                chartPropertiesMenuItem.IsEnabled = chartIsSelected;
                selectChartValuesMenuItem.IsEnabled = chartIsSelected;
                switchRowColumnMenuItem.IsEnabled = chartIsSelected;
            }
        }

        #endregion


        #region Insert

        private void insertButton_ButtonClick(object sender, RoutedEventArgs e)
        {
            try
            {
                // if focused cells contain more rows than columns
                if (VisualEditor.FocusedCells.RowCount > VisualEditor.FocusedCells.ColumnCount)
                {
                    // if full columns are focused
                    if (VisualEditor.FocusedCells.IsCoverColumns)
                        VisualEditor.InsertEmptyColumns();
                    else
                        VisualEditor.InsertCellsAndShiftRight();
                }
                else
                {
                    // if full rows are focused
                    if (VisualEditor.FocusedCells.IsCoverRows)
                        VisualEditor.InsertEmptyRows();
                    else
                        VisualEditor.InsertCellsAndShiftDown();
                }
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
        /// Handles the Click event of InsertCellsShiftRightMenuItem object.
        /// </summary>
        private void insertCellsShiftRightMenuItem_Click(object sender, RoutedEventArgs e)
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
        /// Handles the Click event of InserCellsShiftDownMenuItem object.
        /// </summary>
        private void inserCellsShiftDownMenuItem_Click(object sender, RoutedEventArgs e)
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

        #endregion


        #region Delete

        private void deleteButton_ButtonClick(object sender, RoutedEventArgs e)
        {
            try
            {
                // if focused cells contain more rows than columns
                if (VisualEditor.FocusedCells.RowCount > VisualEditor.FocusedCells.ColumnCount)
                {
                    // if full columns are focused
                    if (VisualEditor.FocusedCells.IsCoverColumns)
                        VisualEditor.RemoveColumns();
                    else
                        VisualEditor.RemoveCellsAndShiftLeft();
                }
                else
                {
                    // if full rows are focused
                    if (VisualEditor.FocusedCells.IsCoverRows)
                        VisualEditor.RemoveRows();
                    else
                        VisualEditor.RemoveCellsAndShiftUp();
                }
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
        /// Handles the Click event of DeleteCellsShiftLeftMenuItem object.
        /// </summary>
        private void deleteCellsShiftLeftMenuItem_Click(object sender, RoutedEventArgs e)
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
        /// Handles the Click event of DeleteCellsShiftUpMenuItem object.
        /// </summary>
        private void deleteCellsShiftUpMenuItem_Click(object sender, RoutedEventArgs e)
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

        #endregion


        #region Format

        private void formatButton_ButtonClick(object sender, RoutedEventArgs e)
        {
            ((MenuItem)sender).IsSubmenuOpen = true;
        }

        /// <summary>
        /// Handles the Click event of AutoFitRowHeightMenuItem object.
        /// </summary>
        private void autoFitRowHeightMenuItem_Click(object sender, RoutedEventArgs e)
        {
            VisualEditor.AutoFitRowHeight();
        }

        /// <summary>
        /// Handles the Click event of SetRowAutoheightMenuItem object.
        /// </summary>
        private void setRowAutoheightMenuItem_Click(object sender, RoutedEventArgs e)
        {
            VisualEditor.SetRowAutoHeight();
        }

        /// <summary>
        /// Handles the Click event of CalculateRowAutoHeightMenuItem object.
        /// </summary>
        private void calculateRowAutoHeightMenuItem_Click(object sender, RoutedEventArgs e)
        {
            VisualEditor.CalculateRowAutoHeight();
        }

        /// <summary>
        /// Handles the Click event of RowHeightMenuItem object.
        /// </summary>
        private void rowHeightMenuItem_Click(object sender, RoutedEventArgs e)
        {
            SetRowHeight();
        }

        /// <summary>
        /// Shows a dialog, which allows to enter a new row height value, and applies height to the selected rows.
        /// </summary>
        public void SetRowHeight()
        {
            // show value editor dialog
            NumberValueEditorWindow dlg = new NumberValueEditorWindow(VisualEditor, VisualEditor.RowsHeight, 0, MAX_ROW_HEIGHT, "Row height");
            dlg.WindowStartupLocation = WindowStartupLocation.CenterOwner;
            dlg.Owner = Application.Current.MainWindow;
            if (dlg.ShowDialog() == true)
            {
                // set height of focused rows
                VisualEditor.RowsHeight = dlg.Value;
            }
        }

        /// <summary>
        /// Handles the Click event of DefaultRowHeightMenuItem object.
        /// </summary>
        private void defaultRowHeightMenuItem_Click(object sender, RoutedEventArgs e)
        {
            WorksheetFormat sheetFormat = VisualEditor.FocusedWorksheet.Format;
            // show value editor dialog
            NumberValueEditorWindow dlg = new NumberValueEditorWindow(VisualEditor, sheetFormat.RowHeight, 0, MAX_ROW_HEIGHT, "Default row height");
            dlg.WindowStartupLocation = WindowStartupLocation.CenterOwner;
            dlg.Owner = Application.Current.MainWindow;
            if (dlg.ShowDialog() == true)
            {
                // set visual format with new default row height
                VisualEditor.SetWorksheetFormat(new WorksheetFormat(sheetFormat.ColumnWidth, dlg.Value, sheetFormat.AutoHeight, sheetFormat.RowsHiddenByDefault));
            }
        }

        /// <summary>
        /// Handles the Click event of ColumnWidthMenuItem object.
        /// </summary>
        private void columnWidthMenuItem_Click(object sender, RoutedEventArgs e)
        {
            SetColumnWidth();
        }

        /// <summary>
        /// Shows a dialog, which allows to enter a new column width value, and applies width to the selected columns.
        /// </summary>
        public void SetColumnWidth()
        {
            // show value editor dialog
            NumberValueEditorWindow dlg = new NumberValueEditorWindow(VisualEditor, VisualEditor.ColumnsWidth, 0, MAX_COLUMN_WIDTH, "Column width");
            dlg.WindowStartupLocation = WindowStartupLocation.CenterOwner;
            dlg.Owner = Application.Current.MainWindow;
            if (dlg.ShowDialog() == true)
            {
                // set width of focused columns
                VisualEditor.ColumnsWidth = dlg.Value;
            }
        }

        /// <summary>
        /// Handles the Click event of DefaultColumnWidthMenuItem object.
        /// </summary>
        private void defaultColumnWidthMenuItem_Click(object sender, RoutedEventArgs e)
        {
            WorksheetFormat sheetFormat = VisualEditor.FocusedWorksheet.Format;
            // show value editor dialog
            NumberValueEditorWindow dlg = new NumberValueEditorWindow(VisualEditor, sheetFormat.ColumnWidth, 0, MAX_COLUMN_WIDTH, "Default column width");
            dlg.WindowStartupLocation = WindowStartupLocation.CenterOwner;
            dlg.Owner = Application.Current.MainWindow;

            if (dlg.ShowDialog() == true)
            {
                // set visual format with new default column width
                VisualEditor.SetWorksheetFormat(new WorksheetFormat(dlg.Value, sheetFormat.RowHeight, sheetFormat.AutoHeight, sheetFormat.RowsHiddenByDefault));
            }
        }

        /// <summary>
        /// Handles the Click event of AutoFitColumnWidthMenuItem object.
        /// </summary>
        private void autoFitColumnWidthMenuItem_Click(object sender, RoutedEventArgs e)
        {
            VisualEditor.AutoFitColumnWidth();
        }

        /// <summary>
        /// Handles the Click event of HideRowsMenuItem object.
        /// </summary>
        private void hideRowsMenuItem_Click(object sender, RoutedEventArgs e)
        {
            VisualEditor.HideRows();
        }

        /// <summary>
        /// Handles the Click event of HideColumnsMenuItem object.
        /// </summary>
        private void hideColumnsMenuItem_Click(object sender, RoutedEventArgs e)
        {
            VisualEditor.HideColumns();
        }

        /// <summary>
        /// Handles the Click event of ShowRowsMenuItem object.
        /// </summary>
        private void showRowsMenuItem_Click(object sender, RoutedEventArgs e)
        {
            VisualEditor.ShowRows();
        }

        /// <summary>
        /// Handles the Click event of ShowColumnsMenuItem object.
        /// </summary>
        private void showColumnsMenuItem_Click(object sender, RoutedEventArgs e)
        {
            VisualEditor.ShowColumns();
        }

        #endregion


        #region Merge

        /// <summary>
        /// Handles the Click event of MergeCenterMenuItem object.
        /// </summary>
        private void mergeCenterMenuItem_Click(object sender, RoutedEventArgs e)
        {
            try
            { 
                VisualEditor.MergeCellsAndCenterText();
            }
            catch (Exception ex)
            {
                DemosTools.ShowErrorMessage(ex);
            }
        }

        /// <summary>
        /// Handles the Click event of MergeMenuItem object.
        /// </summary>
        private void mergeMenuItem_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                VisualEditor.MergeCells();
            }
            catch (Exception ex)
            {
                DemosTools.ShowErrorMessage(ex);
            }
        }

        /// <summary>
        /// Handles the Click event of MergeAcrossMenuItem object.
        /// </summary>
        private void mergeAcrossMenuItem_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                VisualEditor.MergeCellsAcross();
            }
            catch (Exception ex)
            {
                DemosTools.ShowErrorMessage(ex);
            }
        }

        /// <summary>
        /// Handles the Click event of UnmergeMenuItem object.
        /// </summary>
        private void unmergeMenuItem_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                VisualEditor.UnmergeCells();
            }
            catch (Exception ex)
            {
                DemosTools.ShowErrorMessage(ex);
            }
        }

        #endregion


        #region Clear

        private void clearButton_ButtonClick(object sender, RoutedEventArgs e)
        {
            ((MenuItem)sender).IsSubmenuOpen = true;
        }

        /// <summary>
        /// Handles the Click event of ClearAllMenuItem object.
        /// </summary>
        private void clearAllMenuItem_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                VisualEditor.ClearCells();
            }
            catch (Exception ex)
            {
                DemosTools.ShowErrorMessage(ex);
            }
        }

        /// <summary>
        /// Handles the Click event of ClearStylesMenuItem object.
        /// </summary>
        private void clearStylesMenuItem_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                VisualEditor.ClearCellsStyle();
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
        /// Handles the Click event of ClearHyperlinksMenuItem object.
        /// </summary>
        private void clearHyperlinksMenuItem_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                VisualEditor.RemoveHyperlinks();
            }
            catch (Exception ex)
            {
                DemosTools.ShowErrorMessage(ex);
            }
        }

        /// <summary>
        /// Handles the Click event of ClearCommentsMenuItem object.
        /// </summary>
        private void clearCommentsMenuItem_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                VisualEditor.RemoveComments();
            }
            catch (Exception ex)
            {
                DemosTools.ShowErrorMessage(ex);
            }
        }

        #endregion


        #region Fill

        private void fillButton_ButtonClick(object sender, RoutedEventArgs e)
        {
            ((MenuItem)sender).IsSubmenuOpen = true;
        }

        /// <summary>
        /// Handles the Click event of FillDownMenuItem object.
        /// </summary>
        private void fillDownMenuItem_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                VisualEditor.FillCells(CellsFillDirection.Down);
            }
            catch (Exception ex)
            {
                DemosTools.ShowErrorMessage(ex);
            }
        }

        /// <summary>
        /// Handles the Click event of FillRightMenuItem object.
        /// </summary>
        private void fillRightMenuItem_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                VisualEditor.FillCells(CellsFillDirection.Right);
            }
            catch (Exception ex)
            {
                DemosTools.ShowErrorMessage(ex);
            }
        }

        /// <summary>
        /// Handles the Click event of FillUpMenuItem object.
        /// </summary>
        private void fillUpMenuItem_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                VisualEditor.FillCells(CellsFillDirection.Up);
            }
            catch (Exception ex)
            {
                DemosTools.ShowErrorMessage(ex);
            }
        }

        /// <summary>
        /// Handles the Click event of FillLeftMenuItem object.
        /// </summary>
        private void fillLeftMenuItem_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                VisualEditor.FillCells(CellsFillDirection.Left);
            }
            catch (Exception ex)
            {
                DemosTools.ShowErrorMessage(ex);
            }
        }

        #endregion


        #region Chart

        private void VisualEditor_ChartTemplatesRequest(object sender, StreamRequestEventArgs e)
        {
            e.Stream = DemosResourcesManager.GetResourceAsStream(ChartTemplatesResourceName);
        }

        private void chartButton_ButtonClick(object sender, RoutedEventArgs e)
        {
            ((MenuItem)sender).IsSubmenuOpen = true;
        }

        /// <summary>
        /// Handles the Click event of AddChartMenuItem object.
        /// </summary>
        private void addChartMenuItem_Click(object sender, RoutedEventArgs e)
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
        /// Handles the Click event of RemoveChartMenuItem object.
        /// </summary>
        private void removeChartMenuItem_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                VisualEditor.RemoveFocusedDrawing();
            }
            catch (Exception ex)
            {
                DemosTools.ShowErrorMessage(ex);
            }
        }

        /// <summary>
        /// Handles the Click event of SwitchRowColumnMenuItem object.
        /// </summary>
        private void switchRowColumnMenuItem_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                VisualEditor.ChartSwitchRowColumn();
            }
            catch (Exception ex)
            {
                DemosTools.ShowErrorMessage(ex);
            }
        }

        /// <summary>
        /// Handles the Click event of SelectChartValuesMenuItem object.
        /// </summary>
        private void selectChartValuesMenuItem_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                CellReferencesSet chartData = VisualEditor.FocusedDrawing.ChartProperties.GetSeriesValuesReferencesSet();
                VisualEditor.SetFocusedAndSelectedCells(chartData);
            }
            catch (Exception ex)
            {
                DemosTools.ShowWarningMessage(ex.Message);
            }
            SpreadsheetEditor.Focus();
        }

        /// <summary>
        /// Handles the Click event of ChartPropertiesMenuItem object.
        /// </summary>
        private void chartPropertiesMenuItem_Click(object sender, RoutedEventArgs e)
        {
            EditDrawing();
        }

        #endregion


        #region Image

        /// <summary>
        /// Handles the SubmenuOpened event of PictureButton object.
        /// </summary>
        private void pictureButton_SubmenuOpened(object sender, RoutedEventArgs e)
        {
            setPictureMenuItem.IsEnabled = VisualEditor.FocusedDrawing != null && VisualEditor.FocusedDrawing.Type == DrawingType.Picture;
            picturePropertiesMenuItem.IsEnabled = VisualEditor.FocusedDrawing != null;
            removePictureMenuItem.IsEnabled = VisualEditor.FocusedDrawing != null;
        }

        private void pictureButton_ButtonClick(object sender, RoutedEventArgs e)
        {
            ((MenuItem)sender).IsSubmenuOpen = true;
        }

        /// <summary>
        /// Handles the Click event of AddPictureMenuItem object.
        /// </summary>
        private void addPictureMenuItem_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Stream imageStream = GetImageStream();
                if (imageStream != null)
                    using (imageStream)
                        VisualEditor.AddPicture(new ImageData(imageStream));
            }
            catch (Exception ex)
            {
                DemosTools.ShowErrorMessage(ex);
            }
        }

        /// <summary>
        /// Handles the Click event of SetPictureMenuItem object.
        /// </summary>
        private void setPictureMenuItem_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Stream imageStream = GetImageStream();

                if (imageStream != null)
                    using (imageStream)
                        VisualEditor.SetDrawingPicture(new ImageData(imageStream));
            }
            catch (Exception ex)
            {
                DemosTools.ShowErrorMessage(ex);
            }
        }


        /// <summary>
        /// Handles the MouseDoubleClick event of SpreadsheetEditorControl object.
        /// </summary>
        private void SpreadsheetEditorControl_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left && VisualEditor.FocusedDrawing != null)
            {
                e.Handled = true;
                EditDrawing();
            }
        }

        /// <summary>
        /// Handles the Click event of PicturePropertiesMenuItem object.
        /// </summary>
        private void picturePropertiesMenuItem_Click(object sender, RoutedEventArgs e)
        {
            EditDrawing();
        }


        /// <summary>
        /// Handles the Click event of RemovePictureMenuItem object.
        /// </summary>
        private void removePictureMenuItem_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                VisualEditor.RemoveFocusedDrawing();
            }
            catch (Exception ex)
            {
                DemosTools.ShowErrorMessage(ex);
            }
        }

        /// <summary>
        /// Edits the focused drawing.
        /// </summary>
        private void EditDrawing()
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
        /// Returns the image stream.
        /// </summary>
        private Stream GetImageStream()
        {
            // create dialog
            OpenFileDialog dialog = new OpenFileDialog();
            // specify that dialog should open folder with demo images
            DemosTools.SetTestFilesFolder(dialog);
            // set image filters
            CodecsFileFilters.SetFilters(dialog);

            // if image must be loaded
            if (dialog.ShowDialog() == true)
                // return image stream
                return dialog.OpenFile();

            return null;
        }

        #endregion


        #region Link

        /// <summary>
        /// Handles the SubmenuOpened event of HypelinkSplitButton object.
        /// </summary>
        private void hypelinkSplitButton_SubmenuOpened(object sender, RoutedEventArgs e)
        {
            editHyperlinkMenuItem.IsEnabled = VisualEditor.FocusedHyperlink != null;
            removeHyperlinkMenuItem.IsEnabled = !VisualEditor.SelectionContainsSingleCell || VisualEditor.FocusedHyperlink != null;
        }

        /// <summary>
        /// Handles the Click event of AddHyperlinkMenuItem object.
        /// </summary>
        private void addHyperlinkMenuItem_Click(object sender, RoutedEventArgs e)
        {
            EditHyperlinkWindow.ShowDialog(VisualEditor, false);
        }

        /// <summary>
        /// Handles the Click event of EditHyperlinkMenuItem object.
        /// </summary>
        private void editHyperlinkMenuItem_Click(object sender, RoutedEventArgs e)
        {
            EditHyperlinkWindow.ShowDialog(VisualEditor, true);
        }

        /// <summary>
        /// Handles the Click event of RemoveHyperlinkMenuItem object.
        /// </summary>
        private void removeHyperlinkMenuItem_Click(object sender, RoutedEventArgs e)
        {
            VisualEditor.RemoveHyperlinks();
        }

        #endregion

        #endregion

    }
}

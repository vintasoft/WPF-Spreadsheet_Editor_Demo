using System;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

using Vintasoft.Imaging;
using Vintasoft.Imaging.Office.Spreadsheet.Document;
using Vintasoft.Imaging.Office.Spreadsheet.UI;
using Vintasoft.Imaging.Office.Spreadsheet.Wpf.UI;
using Vintasoft.Imaging.Wpf;
using WpfDemosCommonCode;
using WpfDemosCommonCode.CustomControls;

namespace WpfSpreadsheetEditorDemo
{
    /// <summary>
    /// Provides the "Naviagation" panel.
    /// </summary>
    public partial class NavigationPanel : SpreadsheetVisualEditorPanel
    {

        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="NavigationPanel"/> class.
        /// </summary>
        public NavigationPanel()
        {
            InitializeComponent();
        }

        #endregion



        #region Methods

        #region Common        

        /// <summary>
        /// Raises the <see cref="E:SpreadsheetEditorChanged" /> event.
        /// </summary>
        /// <param name="args">The <see cref="PropertyChangedEventArgs{WpfSpreadsheetEditorControl}"/> instance containing the event data.</param>
        protected override void OnSpreadsheetEditorChanged(PropertyChangedEventArgs<WpfSpreadsheetEditorControl> args)
        {
            base.OnSpreadsheetEditorChanged(args);
            if (args.OldValue != null)
            {
                SpreadsheetVisualEditor visualEditor = args.OldValue.VisualEditor;
                visualEditor.FocusedWorksheetChanged -= VisualEditor_FocusedWorksheetChanged;
                visualEditor.EditorChanged -= VisualEditor_EditorChanged;
                visualEditor.InitializationFinished -= VisualEditor_InitializationFinished;
                visualEditor.ZoomChanged -= VisualEditor_ZoomChanged;
                visualEditor.SynchronizationStarted -= VisualEditor_SynchronizationStarted;
                visualEditor.SynchronizationFinished -= VisualEditor_SynchronizationFinished;
                visualEditor.EditCellValueStarted -= VisualEditor_EditingCellValueStateChanged;
                visualEditor.EditCellValueFinished -= VisualEditor_EditingCellValueStateChanged;
            }
            if (args.NewValue != null)
            {
                SpreadsheetVisualEditor visualEditor = args.NewValue.VisualEditor;
                visualEditor.FocusedWorksheetChanged += VisualEditor_FocusedWorksheetChanged;
                visualEditor.EditorChanged += VisualEditor_EditorChanged;
                visualEditor.InitializationFinished += VisualEditor_InitializationFinished;
                visualEditor.ZoomChanged += VisualEditor_ZoomChanged;
                visualEditor.SynchronizationStarted += VisualEditor_SynchronizationStarted;
                visualEditor.SynchronizationFinished += VisualEditor_SynchronizationFinished;
                visualEditor.EditCellValueStarted += VisualEditor_EditingCellValueStateChanged;
                visualEditor.EditCellValueFinished += VisualEditor_EditingCellValueStateChanged;
            }
            UpdateUI();
        }

        /// <summary>
        /// Updates the user interface.
        /// </summary>
        private void UpdateUI()
        {
            worksheetComboBox.BeginInit();
            worksheetComboBox.Items.Clear();
            if (VisualEditor != null && VisualEditor.Document != null)
            {
                foreach (Worksheet worksheet in VisualEditor.Document.Worksheets)
                {
                    worksheetComboBox.Items.Add(worksheet.Name);
                }
            }
            UpdateFocusedWorksheetUI();
            worksheetComboBox.EndInit();
        }

        /// <summary>
        /// Updates the user interface of focused worksheet.
        /// </summary>
        private void UpdateFocusedWorksheetUI()
        {
            // if there is no focused worksheet, disable panel
            if (VisualEditor == null || VisualEditor.Document == null || VisualEditor.Document.FocusedWorksheet == null)
            {
                firstWorksheetButton.IsEnabled = false;
                prevWorksheetButton.IsEnabled = false;
                worksheetComboBox.IsEnabled = false;
                nextWorksheetButton.IsEnabled = false;
                lastWorksheetButton.IsEnabled = false;
                addWorksheetButton.IsEnabled = false;
                worksheetsActionsButton.IsEnabled = false;
                zoomInButton.IsEnabled = false;
                zoomComboBox.IsEnabled = false;
                zoomOutButton.IsEnabled = false;
                copyMenuItem.IsEnabled = false;
            }
            else
            {
                bool isChangingFocusedCellValue = VisualEditor.IsChangingFocusedCellValue;

                int worksheetCount = VisualEditor.Document.Worksheets.Count;
                int worksheetIndex = VisualEditor.FocusedWorksheetIndex;

                // navigation
                firstWorksheetButton.IsEnabled = worksheetIndex > 0 && !isChangingFocusedCellValue;
                prevWorksheetButton.IsEnabled = worksheetIndex > 0 && !isChangingFocusedCellValue;
                worksheetComboBox.IsEnabled = !isChangingFocusedCellValue;
                if (VisualEditor.FocusedWorksheet != null)
                    worksheetComboBox.SelectedItem = VisualEditor.FocusedWorksheet.Name;
                else
                    worksheetComboBox.Text = "";
                nextWorksheetButton.IsEnabled = worksheetIndex < worksheetCount - 1 && !isChangingFocusedCellValue;
                lastWorksheetButton.IsEnabled = worksheetIndex < worksheetCount - 1 && !isChangingFocusedCellValue;

                // worksheet manipulation
                addWorksheetButton.IsEnabled = !VisualEditor.IsSynchronizing && !isChangingFocusedCellValue;
                worksheetsActionsButton.IsEnabled = !isChangingFocusedCellValue;
                removeMenuItem.IsEnabled = !VisualEditor.IsSynchronizing && worksheetCount > 1 && !isChangingFocusedCellValue;
                moveMenuItem.IsEnabled = !VisualEditor.IsSynchronizing && worksheetCount > 1 && !isChangingFocusedCellValue;
                copyMenuItem.IsEnabled= !VisualEditor.IsSynchronizing && !isChangingFocusedCellValue;

                // worksheet properties
                showHeadingsMenuItem.IsChecked = VisualEditor.ShowHeadings;
                showFormulasMenuItem.IsChecked = VisualEditor.ShowFormulas;
                showGridMenuItem.IsChecked = VisualEditor.ShowGrid;

                // zoom
                int zoom = (int)Math.Round(VisualEditor.Zoom);
                zoomOutButton.IsEnabled = zoom > 10;
                zoomComboBox.IsEnabled = true;
                zoomComboBox.Text = zoom.ToString() + "%";
                zoomInButton.IsEnabled = zoom < 400;
            }
        }

        /// <summary>
        /// Handles the Click event of PanModeButton object.
        /// </summary>
        private void panModeButton_Click(object sender, RoutedEventArgs e)
        {
            panModeButton.IsChecked = !panModeButton.IsChecked;
            if (panModeButton.IsChecked)
                VisualEditor.InteractionMode = SpreadsheetVisualEditorInteractionMode.PanAndSelection;
            else
                VisualEditor.InteractionMode = SpreadsheetVisualEditorInteractionMode.Selection;
        }

        private void VisualEditor_InitializationFinished(object sender, EventArgs e)
        {
            UpdateUI();
        }

        private void VisualEditor_EditorChanged(object sender, PropertyChangedEventArgs<Vintasoft.Imaging.Office.Spreadsheet.SpreadsheetEditor> e)
        {
            UpdateUI();
        }

        private void VisualEditor_SynchronizationFinished(object sender, EventArgs e)
        {
            UpdateFocusedWorksheetUI();
        }

        private void VisualEditor_SynchronizationStarted(object sender, EventArgs e)
        {
            UpdateFocusedWorksheetUI();
        }


        private void VisualEditor_FocusedWorksheetChanged(object sender, PropertyChangedEventArgs<Worksheet> e)
        {
            UpdateFocusedWorksheetUI();
        }

        /// <summary>
        /// Handles the EditCellValueStarted/EditCellValueFinished event of the VisualEditor.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
        private void VisualEditor_EditingCellValueStateChanged(object sender, EventArgs e)
        {
            UpdateFocusedWorksheetUI();
        }


        #endregion


        #region Navigation

        /// <summary>
        /// Handles the Click event of FirstWorksheetButton object.
        /// </summary>
        private void firstWorksheetButton_Click(object sender, RoutedEventArgs e)
        {
            // move focused worksheet to the first position
            VisualEditor.FocusedWorksheetIndex = 0;
        }

        /// <summary>
        /// Handles the Click event of PrevWorksheetButton object.
        /// </summary>
        private void prevWorksheetButton_Click(object sender, RoutedEventArgs e)
        {
            // move focused worksheet to the previous position
            VisualEditor.FocusedWorksheetIndex--;
        }

        /// <summary>
        /// Handles the SelectionChanged event of WorksheetComboBox object.
        /// </summary>
        private void worksheetComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // set focused worksheet
            if (worksheetComboBox.SelectedIndex != -1)
                VisualEditor.FocusedWorksheetIndex = worksheetComboBox.SelectedIndex;
        }

        /// <summary>
        /// Handles the Click event of NextWorksheetButton object.
        /// </summary>
        private void nextWorksheetButton_Click(object sender, RoutedEventArgs e)
        {
            // move focused worksheet to the next position
            VisualEditor.FocusedWorksheetIndex++;
        }

        /// <summary>
        /// Handles the Click event of LastWorksheetButton object.
        /// </summary>
        private void lastWorksheetButton_Click(object sender, RoutedEventArgs e)
        {
            // move focused worksheet to the last position
            VisualEditor.FocusedWorksheetIndex = VisualEditor.Document.Worksheets.Count - 1;
        }


        #endregion


        #region Worksheet manipulations and properties

        /// <summary>
        /// Handles the Click event of AddWorksheetButton object.
        /// </summary>
        private void addWorksheetButton_Click(object sender, RoutedEventArgs e)
        {
            // add new worksheet
            try
            {
                VisualEditor.AddWorksheet();
            }
            catch (Exception ex)
            {
                DemosTools.ShowErrorMessage(ex);
            }
            UpdateUI();
        }

        /// <summary>
        /// Handles the Click event of CopyMenuItem object.
        /// </summary>
        private void copyMenuItem_Click(object sender, RoutedEventArgs e)
        {
            // copy focused worksheet
            try
            {
                VisualEditor.CopyWorksheet();
            }
            catch (Exception ex)
            {
                DemosTools.ShowErrorMessage(ex);
            }
            UpdateUI();
        }

        /// <summary>
        /// Handles the Click event of RemoveMenuItem object.
        /// </summary>
        private void removeMenuItem_Click(object sender, RoutedEventArgs e)
        {
            if (MessageBox.Show(string.Format("Remove worksheet '{0}'?", VisualEditor.FocusedWorksheet.Name),
                "Remove Worksheet", MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.Yes)
            {
                // remove focused worksheet
                try
                {
                    VisualEditor.RemoveWorksheet();
                }
                catch (Exception ex)
                {
                    DemosTools.ShowErrorMessage(ex);
                }
                UpdateUI();
            }
        }

        /// <summary>
        /// Handles the Click event of MoveMenuItem object.
        /// </summary>
        private void moveMenuItem_Click(object sender, RoutedEventArgs e)
        {
            // show "Move worksheet" dialog
            MoveWorksheetWindow dlg = new MoveWorksheetWindow(VisualEditor);
            dlg.WindowStartupLocation = WindowStartupLocation.CenterOwner;
            dlg.Owner = Application.Current.MainWindow;
            if (dlg.ShowDialog() == true)
            {
                if (dlg.IsWorksheetOrderChanged)
                {
                    UpdateUI();
                }
            }
        }

        /// <summary>
        /// Handles the Click event of RenameMenuItem object.
        /// </summary>
        private void renameMenuItem_Click(object sender, RoutedEventArgs e)
        {
            // show "Rename worksheet" dialog
            RenameWorksheetWindow dlg = new RenameWorksheetWindow(VisualEditor);
            dlg.WindowStartupLocation = WindowStartupLocation.CenterOwner;
            dlg.Owner = Application.Current.MainWindow;
            if (dlg.ShowDialog() == true)
            {
                if (dlg.IsWorksheetNameChanged)
                {
                    UpdateUI();
                }
            }
        }

        /// <summary>
        /// Handles the Click event of ShowHeadingsMenuItem object.
        /// </summary>
        private void showHeadingsMenuItem_Click(object sender, RoutedEventArgs e)
        {
            VisualEditor.ShowHeadings = !VisualEditor.ShowHeadings;
            showHeadingsMenuItem.IsChecked = VisualEditor.ShowHeadings;
        }

        /// <summary>
        /// Handles the Click event of ShowFormulasMenuItem object.
        /// </summary>
        private void showFormulasMenuItem_Click(object sender, RoutedEventArgs e)
        {
            VisualEditor.ShowFormulas = !VisualEditor.ShowFormulas;
            showFormulasMenuItem.IsChecked = VisualEditor.ShowFormulas;
        }

        /// <summary>
        /// Handles the Click event of ShowGridMenuItem object.
        /// </summary>
        private void showGridMenuItem_Click(object sender, RoutedEventArgs e)
        {
            VisualEditor.ShowGrid = !VisualEditor.ShowGrid;
            showGridMenuItem.IsChecked = VisualEditor.ShowGrid;
        }

        /// <summary>
        /// Handles the Click event of GridColorMenuItem object.
        /// </summary>
        private void gridColorMenuItem_Click(object sender, RoutedEventArgs e)
        {
            // create "Color" dialog
            ColorPickerDialog colorDialog = new ColorPickerDialog();
            colorDialog.WindowStartupLocation = WindowStartupLocation.CenterOwner;
            colorDialog.Owner = Application.Current.MainWindow;
            colorDialog.CanEditAlphaChannel = false;
            colorDialog.StartingColor = WpfObjectConverter.Convert(VisualEditor.GridColor);

            // show "Color" dialog
            if (colorDialog.ShowDialog() == true)
                VisualEditor.GridColor = WpfObjectConverter.Convert(colorDialog.SelectedColor);
        }

        /// <summary>
        /// Handles the Click event of FormatMenuItem object.
        /// </summary>
        private void formatMenuItem_Click(object sender, RoutedEventArgs e)
        {
            // show "Worksheet Format" dialog
            WorksheetFormatWindow.ShowDialog(VisualEditor);
        }

        #endregion


        #region Zoom

        /// <summary>
        /// Handles the Click event of ZoomOutButton object.
        /// </summary>
        private void zoomOutButton_Click(object sender, RoutedEventArgs e)
        {
            // get current zoom comboBox value
            int index = GetPredefinedZoomIndex();
            int predefinedZoom = int.Parse(GetZoomText((ComboBoxItem)zoomComboBox.Items[index]));

            // set lower zoom value
            if (predefinedZoom < VisualEditor.Zoom)
                zoomComboBox.SelectedIndex = index;
            else if (index < zoomComboBox.Items.Count - 1)
                zoomComboBox.SelectedIndex = index + 1;
        }

        /// <summary>
        /// Handles the Click event of ZoomInButton object.
        /// </summary>
        private void zoomInButton_Click(object sender, RoutedEventArgs e)
        {
            // get current zoom comboBox value
            int index = GetPredefinedZoomIndex();
            int predefinedZoom = int.Parse(GetZoomText((ComboBoxItem)zoomComboBox.Items[index]));

            // set higher zoom value
            if (predefinedZoom > VisualEditor.Zoom)
                zoomComboBox.SelectedIndex = index;
            else if (index > 0)
                zoomComboBox.SelectedIndex = index - 1;
        }

        /// <summary>
        /// Handles the SelectionChanged event of ZoomComboBox object.
        /// </summary>
        private void zoomComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (e.AddedItems.Count == 1)
                SetZoom(GetZoomText((ComboBoxItem)e.AddedItems[0]));
        }

        /// <summary>
        /// Handles the PreviewKeyDown event of ZoomComboBox object.
        /// </summary>
        private void zoomComboBox_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
                SetZoom();
        }

        /// <summary>
        /// Handles the LostFocus event of ZoomComboBox object.
        /// </summary>
        private void zoomComboBox_LostFocus(object sender, RoutedEventArgs e)
        {
            SetZoom();
        }

        /// <summary>
        /// Sets the zoom in visual editor.
        /// </summary>
        private void SetZoom()
        {
            SetZoom(GetZoomText(zoomComboBox.Text));
        }

        /// <summary>
        /// Sets the zoom in visual editor.
        /// </summary>
        /// <param name="zoomText">The zoom text.</param>
        private void SetZoom(string zoomText)
        {
            int zoom;
            if (int.TryParse(zoomText, out zoom))
            {
                if (zoom < 10)
                    zoom = 10;
                else if (zoom > 400)
                    zoom = 400;
                VisualEditor.Zoom = zoom;
            }
        }

        /// <summary>
        /// Returns the zoom value without percent sign.
        /// </summary>
        /// <param name="item">The <see cref="ComboBoxItem"/>.</param>
        private string GetZoomText(ComboBoxItem item)
        {
            return GetZoomText(item.Content.ToString());
        }

        /// <summary>
        /// Returns the zoom value without percent sign.
        /// </summary>
        /// <param name="text">Zoom value with percent sign.</param>
        private string GetZoomText(string text)
        {
            return text.Replace("%", "").Trim();
        }

        /// <summary>
        /// Returns the index of predefined zoom.
        /// </summary>
        private int GetPredefinedZoomIndex()
        {
            double zoom;
            if (!double.TryParse(GetZoomText(zoomComboBox.Text), NumberStyles.Number, UICulture, out zoom))
                zoom = 100;
            for (int i = zoomComboBox.Items.Count - 1; i >= 0; i--)
            {
                string zoomText = GetZoomText((ComboBoxItem)zoomComboBox.Items[i]);
                if (int.Parse(zoomText) >= zoom)
                    return i;
            }
            return zoomComboBox.Items.Count - 1;
        }

        /// <summary>
        /// Handles the ZoomChanged event of VisualEditor object.
        /// </summary>
        private void VisualEditor_ZoomChanged(object sender, PropertyChangedEventArgs<double> e)
        {
            UpdateFocusedWorksheetUI();
        }

        #endregion

        #endregion

    }
}

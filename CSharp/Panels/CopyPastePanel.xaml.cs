using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Vintasoft.Imaging;
using Vintasoft.Imaging.Office.Spreadsheet;
using Vintasoft.Imaging.Office.Spreadsheet.Document;
using Vintasoft.Imaging.Office.Spreadsheet.Wpf.UI;

using WpfDemosCommonCode;

namespace WpfSpreadsheetEditorDemo
{
    /// <summary>
    /// Provides a "Copy/Paste" panel.
    /// </summary>
    public partial class CopyPastePanel : SpreadsheetVisualEditorPanel
    {

        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="CopyPastePanel"/> class.
        /// </summary>
        public CopyPastePanel()
        {
            InitializeComponent();
        }

        #endregion



        #region Methods

        /// <summary>
        /// Raises the <see cref="E:SpreadsheetEditorChanged" /> event.
        /// </summary>
        /// <param name="args">The <see cref="PropertyChangedEventArgs{SpreadsheetEditorControl}"/> instance containing the event data.</param>
        protected override void OnSpreadsheetEditorChanged(PropertyChangedEventArgs<WpfSpreadsheetEditorControl> args)
        {
            base.OnSpreadsheetEditorChanged(args);

            if (args.OldValue != null)
            {
                args.OldValue.VisualEditor.CellsClipboardChanged -= VisualEditor_CellsBufferChanged;
                args.OldValue.VisualEditor.FocusedCellsChanged -= VisualEditor_FocusedCellsChanged;
                args.OldValue.VisualEditor.FocusedWorksheetChanged -= VisualEditor_FocusedWorksheetChanged;
            }

            if (args.NewValue != null)
            {
                args.NewValue.VisualEditor.CellsClipboardChanged += VisualEditor_CellsBufferChanged;
                args.NewValue.VisualEditor.FocusedCellsChanged += VisualEditor_FocusedCellsChanged;                
                args.NewValue.VisualEditor.FocusedWorksheetChanged += VisualEditor_FocusedWorksheetChanged;
            }

            UpdateUI();
        }

        private void VisualEditor_FocusedCellsChanged(object sender, PropertyChangedEventArgs<CellReferences> e)
        {
            UpdateUI();
        }

        private void VisualEditor_CellsBufferChanged(object sender, PropertyChangedEventArgs<SheetCellsClipboard> e)
        {
            UpdateUI();
        }

        private void VisualEditor_FocusedWorksheetChanged(object sender, PropertyChangedEventArgs<Worksheet> e)
        {
            UpdateUI();
        }

        /// <summary>
        /// Updates the user interface.
        /// </summary>
        private void UpdateUI()
        {
            if (VisualEditor.IsFocusedWorksheetChanging)
                return;

            if (VisualEditor.FocusedWorksheet == null || VisualEditor.FocusedCells == null)
            {
                IsEnabled = false;
            }
            else
            {
                IsEnabled = true;
                if (VisualEditor.CellsClipboard != null)
                {
                    pasteButton.IsEnabled = true;
                    ((MenuItem)pasteButton.Items[0]).IsEnabled = !VisualEditor.IsChangingFocusedCellValue;
                    ((MenuItem)pasteButton.Items[3]).IsEnabled = !VisualEditor.IsChangingFocusedCellValue;
                    ((MenuItem)pasteButton.Items[5]).IsEnabled = !VisualEditor.IsChangingFocusedCellValue;
                    ((MenuItem)pasteButton.Items[1]).IsEnabled = !VisualEditor.IsChangingFocusedCellValue;
                    ((MenuItem)pasteButton.Items[2]).IsEnabled = !VisualEditor.IsChangingFocusedCellValue;
                }
                else
                {
                    ((MenuItem)pasteButton.Items[0]).IsEnabled = false;
                    ((MenuItem)pasteButton.Items[3]).IsEnabled = false;
                    ((MenuItem)pasteButton.Items[5]).IsEnabled = false;
                    ((MenuItem)pasteButton.Items[1]).IsEnabled = false;
                    ((MenuItem)pasteButton.Items[2]).IsEnabled = false;
                }
                pasteButton.IsEnabled = true;
            }
        }

        private void copyButton_ButtonClick(object sender, RoutedEventArgs e)
        {
            VisualEditor.Copy();
        }

        /// <summary>
        /// Handles the Click event of CutMenuItem object.
        /// </summary>
        private void cutMenuItem_Click(object sender, RoutedEventArgs e)
        {
            VisualEditor.Cut();
        }

        private void pasteButton_ButtonClick(object sender, RoutedEventArgs e)
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
        /// Handles the Click event of PasteContentsMenuItem object.
        /// </summary>
        private void pasteContentsMenuItem_Click(object sender, RoutedEventArgs e)
        {
            VisualEditor.PasteCellsContent();
        }

        /// <summary>
        /// Handles the Click event of PasteValuesAndStyleMenuItem object.
        /// </summary>
        private void pasteValuesAndStyleMenuItem_Click(object sender, RoutedEventArgs e)
        {
            VisualEditor.PasteCellsValueAndStyle();
        }

        /// <summary>
        /// Handles the Click event of PasteValuesMenuItem object.
        /// </summary>
        private void pasteValuesMenuItem_Click(object sender, RoutedEventArgs e)
        {
            VisualEditor.PasteCellsValue();
        }

        /// <summary>
        /// Handles the Click event of PasteFormulasMenuItem object.
        /// </summary>
        private void pasteFormulasMenuItem_Click(object sender, RoutedEventArgs e)
        {
            VisualEditor.PasteCellsFormula();
        }

        /// <summary>
        /// Handles the Click event of PasteSpecialMenuItem object.
        /// </summary>
        private void pasteSpecialMenuItem_Click(object sender, RoutedEventArgs e)
        {
            CellPasteSpecialWindow dlg = new CellPasteSpecialWindow(VisualEditor);
            dlg.WindowStartupLocation = WindowStartupLocation.CenterOwner;
            dlg.Owner = Application.Current.MainWindow;
            if (dlg.ShowDialog() == true)
            {
                UpdateUI();
            }
        }

        /// <summary>
        /// Handles the SubmenuOpened event of PasteButton object.
        /// </summary>
        private void pasteButton_SubmenuOpened(object sender, RoutedEventArgs e)
        {
            UpdateUI();
        }

        #endregion

    }
}

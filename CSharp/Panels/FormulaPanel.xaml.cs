using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

using Vintasoft.Imaging;
using Vintasoft.Imaging.Office.OpenXml.Wpf.UI;
using Vintasoft.Imaging.Office.Spreadsheet.Document;
using Vintasoft.Imaging.Office.Spreadsheet.UI;
using Vintasoft.Imaging.Office.Spreadsheet.Wpf.UI;
using Vintasoft.Imaging.Wpf.UI;

namespace WpfSpreadsheetEditorDemo
{
    /// <summary>
    /// Provides a "Formula" panel.
    /// </summary>
    public partial class FormulaPanel : SpreadsheetVisualEditorPanel
    {

        #region Fields

        /// <summary>
        /// Indicates when changing <see cref="cellsReferenceComboBox"/>.Text.
        /// </summary>
        bool _changingCellsReferenceComboBoxText = false;

        #endregion



        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="FormulaPanel"/> class.
        /// </summary>
        public FormulaPanel()
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
                SpreadsheetVisualEditor visualEditor = args.OldValue.VisualEditor;
                visualEditor.EditCellValueStarted -= VisualEditor_EditCellValueStarted;
                visualEditor.EditCellValueFinished -= VisualEditor_EditCellValueFinished;
                visualEditor.FocusedCellChanged -= VisualEditor_FocusedCellChanged;
                visualEditor.FocusedCellsChanged -= VisualEditor_FocusedCellsChanged;
                visualEditor.FormulaSyntaxError -= VisualEditor_FormulaSyntaxError;
            }
            if (args.NewValue != null)
            {
                SpreadsheetVisualEditor visualEditor = args.NewValue.VisualEditor;
                visualEditor.CellValueExternalEditor = new WpfTextBoxProvider(cellValueTextBox);
                visualEditor.EditCellValueStarted += VisualEditor_EditCellValueStarted;
                visualEditor.EditCellValueFinished += VisualEditor_EditCellValueFinished;
                visualEditor.FocusedCellChanged += VisualEditor_FocusedCellChanged;
                visualEditor.FocusedCellsChanged += VisualEditor_FocusedCellsChanged;
                visualEditor.FormulaSyntaxError += VisualEditor_FormulaSyntaxError;
            }
            else
            {
                cellsReferenceComboBox.Text = "";
            }
            buttonCancel.IsEnabled = VisualEditor.IsChangingFocusedCellValue;
            buttonOk.IsEnabled = VisualEditor.IsChangingFocusedCellValue;
        }

        /// <summary>
        /// Handles the FocusedCellsChanged event of the VisualEditor control.
        /// </summary>
        private void VisualEditor_FocusedCellsChanged(object sender, PropertyChangedEventArgs<CellReferences> e)
        {
            if (e.NewValue != null)
            {
                if (VisualEditor.IsFocusedCellsChanging)
                {
                    if (e.NewValue.ColumnCount == 1 && e.NewValue.RowCount == 1)
                    {
                        cellsReferenceComboBox.Text = VisualEditor.FocusedCell.GetA1Name();
                    }
                    else
                    {
                        cellsReferenceComboBox.Text = string.Format("{0}R x {1}C", e.NewValue.RowCount, e.NewValue.ColumnCount);
                        InvalidateVisual();
                    }
                }
            }
            else
            {
                cellsReferenceComboBox.Text = "";
            }
        }

        /// <summary>
        /// Handles the FocusedCellChanged event of the VisualEditor control.
        /// </summary>
        private void VisualEditor_FocusedCellChanged(object sender, PropertyChangedEventArgs<CellReference> e)
        {
            if (!VisualEditor.IsFocusedCellsChanging)
            {
                _changingCellsReferenceComboBoxText = true;
                if (e.NewValue != null)
                {
                    cellsReferenceComboBox.Text = e.NewValue.GetA1Name();
                }
                else
                {
                    cellsReferenceComboBox.Text = "";
                }
                _changingCellsReferenceComboBoxText = false;
            }
        }

        /// <summary>
        /// Handles the EditCellValueStarted event of the VisualEditor control.
        /// </summary>
        private void VisualEditor_EditCellValueStarted(object sender, EventArgs e)
        {
            buttonCancel.IsEnabled = true;
            buttonOk.IsEnabled = true;
        }

        /// <summary>
        /// Handles the EditCellValueFinished event of the VisualEditor control.
        /// </summary>
        private void VisualEditor_EditCellValueFinished(object sender, EventArgs e)
        {
            buttonCancel.IsEnabled = false;
            buttonOk.IsEnabled = false;
        }
        
        /// <summary>
        /// Handles the PreviewKeyDown event of CellsReferenceComboBox object.
        /// </summary>
        private void cellsReferenceComboBox_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                CellReferences cellReferences = SetCellReference();
                if (cellReferences != null)
                    if (!cellsReferenceComboBox.Items.Contains(cellReferences))
                        cellsReferenceComboBox.Items.Insert(0, cellReferences);
            }
        }

        /// <summary>
        /// Sets the cell reference.
        /// </summary>
        /// <returns></returns>
        private CellReferences SetCellReference()
        {
            try
            {
                if (!string.IsNullOrEmpty(cellsReferenceComboBox.Text))
                {
                    VisualEditor.SetFocusedAndSelectedCells(cellsReferenceComboBox.Text);
                    return VisualEditor.FocusedCells;
                }
            }
            catch (Exception ex)
            {
                WpfDemosCommonCode.DemosTools.ShowErrorMessage(ex);
            }
            return null;
        }


        /// <summary>
        /// Handles the Click event of the buttonCancel control.
        /// </summary>
        private void buttonCancel_Click(object sender, RoutedEventArgs e)
        {
            VisualEditor.CancelEditCellValue();
        }
        
        /// <summary>
        /// Handles the Click event of the buttonOk control.
        /// </summary>
        private void buttonOk_Click(object sender, RoutedEventArgs e)
        {
            Exception parsingException = null;
            if (!VisualEditor.TryFinishEditCellFormula(out parsingException))
            {
                if (MessageBox.Show(parsingException.Message, "Formula Syntax Error", MessageBoxButton.OKCancel, MessageBoxImage.Warning) == MessageBoxResult.Cancel)
                    VisualEditor.CancelEditCellValue();
            }
        }

        /// <summary>
        /// Handles the SelectedIndexChanged event of the cellsReferenceComboBox control.
        /// </summary>
        private void cellsReferenceComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!_changingCellsReferenceComboBoxText && !VisualEditor.IsFocusedCellsChanging)
                SetCellReference();
        }

        /// <summary>
        /// Handles the Click event of InsertFunctionToolStripButton object.
        /// </summary>
        private void insertFunctionToolStripButton_Click(object sender, RoutedEventArgs e)
        {
            string function = SelectFunctionWindow.SelectFunction(VisualEditor.Document);
            if (function != null)
            {
                VisualEditor.InsertFormulaInFocusedCell(function + "()");
            }
            else
            {
                VisualEditor.StartEditFocusedCellFormula(true);
            }
        }

        /// <summary>
        /// Handles the FormulaSyntaxError event of the VisualEditor.
        /// </summary>
        private void VisualEditor_FormulaSyntaxError(object sender, Vintasoft.Imaging.ExceptionEventArgs e)
        {
            MessageBox.Show(e.Exception.Message, "Formula Syntax Error", MessageBoxButton.OK, MessageBoxImage.Warning);
        }

        #endregion

    }
}

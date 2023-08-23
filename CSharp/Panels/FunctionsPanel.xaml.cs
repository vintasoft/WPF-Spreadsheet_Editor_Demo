using System;
using System.Windows;

using Vintasoft.Imaging;
using Vintasoft.Imaging.Office.Spreadsheet.Document;
using Vintasoft.Imaging.Office.Spreadsheet.Wpf.UI;

namespace WpfSpreadsheetEditorDemo
{
    /// <summary>
    /// Provides a "Functions" panel.
    /// </summary>
    public partial class FunctionsPanel : SpreadsheetVisualEditorPanel
    {

        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="FunctionsPanel"/> class.
        /// </summary>
        public FunctionsPanel()
        {
            InitializeComponent();
        } 

        #endregion



        #region Methods

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
                args.OldValue.VisualEditor.FocusedWorksheetChanged -= VisualEditor_FocusedWorksheetChanged;
            }

            if (args.NewValue != null)
            {
                args.NewValue.VisualEditor.FocusedWorksheetChanged += VisualEditor_FocusedWorksheetChanged;
            }

            UpdateUI();
        }

        #endregion


        #region PRIVATE

        private void VisualEditor_FocusedWorksheetChanged(object sender, PropertyChangedEventArgs<Worksheet> e)
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
                showFomulasButton.IsChecked = false;
            }
            else
            {
                IsEnabled = true;
                showFomulasButton.IsChecked = VisualEditor.ShowFormulas;
            }
        }

        /// <summary>
        /// Handles the Click event of InsertFunctionButton object.
        /// </summary>
        private void insertFunctionButton_Click(object sender, RoutedEventArgs e)
        {
            // get the function selected in dialog
            string function = SelectFunctionWindow.SelectFunction(VisualEditor.Document);
            // if function is found
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
        /// Handles the Click event of ShowFomulasButton object.
        /// </summary>
        private void showFomulasButton_Click(object sender, RoutedEventArgs e)
        {
            VisualEditor.ShowFormulas = !VisualEditor.ShowFormulas;
            UpdateUI();
        }

        #endregion

        #endregion

    }
}

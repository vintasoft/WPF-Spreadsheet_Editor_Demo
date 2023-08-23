using System;
using System.Windows;

using Vintasoft.Imaging.Office.Spreadsheet.UI;
using Vintasoft.Imaging.Wpf;

namespace WpfSpreadsheetEditorDemo
{
    /// <summary>
    /// A window that allows to view and change settings of the spreadsheet visual editor.
    /// </summary>
    public partial class OptionsWindow : Window
    {

        #region Fields

        /// <summary>
        /// Spreadsheet visual editor.
        /// </summary>
        SpreadsheetVisualEditor _visualEditor;

        #endregion



        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="OptionsWindow"/> class.
        /// </summary>
        public OptionsWindow()
        {
            InitializeComponent();
        }

        #endregion



        #region Methods

        /// <summary>
        /// Shows this form with current visual editor settings.
        /// </summary>
        /// <param name="visualEditor">Spreadsheet visual editor.</param>
        public static bool? ShowDialog(SpreadsheetVisualEditor visualEditor)
        {
            OptionsWindow window = new OptionsWindow();
            window.WindowStartupLocation = WindowStartupLocation.CenterOwner;
            window.Owner = Application.Current.MainWindow;
            window.SetVisualEditor(visualEditor);
            return window.ShowDialog();
        }

        /// <summary>
        /// Sets the visual editor settings to this form UI.
        /// </summary>
        /// <param name="visualEditor">Spreadsheet visual editor.</param>
        private void SetVisualEditor(SpreadsheetVisualEditor visualEditor)
        {
            _visualEditor = visualEditor;
            UpdateUI();
        }

        /// <summary>
        /// Updates the user interface of this form.
        /// </summary>
        private void UpdateUI()
        {
            // culture
            cultureComboBox.Text = _visualEditor.DocumentCulture;
            uiCultureComboBox.Text = _visualEditor.DocumentUICulture;

            calculationMinIntervalNumericUpDown.Value = _visualEditor.EditorSettings.CalculationMinInterval;

            // appearances
            focusedCellsAppearanceEditor.CellsAppearance = _visualEditor.FocusedCellsAppearance;
            bufferCellsAppearanceEditor.CellsAppearance = _visualEditor.CellsClipboardAppearance;
            formulaCellsAppearanceEditor.CellsAppearance = _visualEditor.FormulaAppearance;
            focusedFormulaCellsAppearanceEditor.CellsAppearance = _visualEditor.FormulaFocusedAppearance;

            // headings
            headingsColorPanelControl.Color = WpfObjectConverter.Convert(_visualEditor.HeadingsColor);
            headingsTextColorPanelControl.Color = WpfObjectConverter.Convert(_visualEditor.HeadingsTextColor);
            headingsBorderColorPanelControl.Color = WpfObjectConverter.Convert(_visualEditor.HeadingsBorderColor);
            selectedCellColorPanelControl.Color = WpfObjectConverter.Convert(_visualEditor.SelectedCellColor);
            selectedHeaderColorPanelControl.Color = WpfObjectConverter.Convert(_visualEditor.SelectedHeaderColor);
            coveredHeaderColorPanelControl.Color = WpfObjectConverter.Convert(_visualEditor.CoveredHeaderColor);

            // errors
            errorIndicatorColorPanelControl.Color = WpfObjectConverter.Convert(_visualEditor.ErrorIndicatorColor);
            errorIndicatorSizeNumericUpDown.Value = (int)Math.Round(_visualEditor.ErrorIndicatorSize, 0);

            // comments
            commentIndicatorColorPanelControl.Color = WpfObjectConverter.Convert(_visualEditor.CommentIndicatorColor);
            commentIndicatorSizeNumericUpDown.Value = (int)Math.Round(_visualEditor.CommentIndicatorSize, 0);
            commentAppearanceEditor.CellsAppearance = _visualEditor.CommentAppearance;
            focusedCommentAppearanceEditor.CellsAppearance = _visualEditor.CommentFocusedAppearance;

            // miscellaneous
            hyperlinkColorPanelControl.Color = WpfObjectConverter.Convert(_visualEditor.HyperlinkColor);
            gridColorAlphaNumericUpDown.Value = (int)Math.Round(255 - _visualEditor.GridColorAlpha * 255, 0);
        }

        /// <summary>
        /// Handles the Click event of ButtonOk object.
        /// </summary>
        private void buttonOk_Click(object sender, RoutedEventArgs e)
        {
            // culture
            if (_visualEditor.DocumentCulture != cultureComboBox.Text ||
                _visualEditor.DocumentUICulture != uiCultureComboBox.Text)
            {
                _visualEditor.SetDocumentCulture(cultureComboBox.Text, uiCultureComboBox.Text);
            }

            _visualEditor.EditorSettings.CalculationMinInterval = (int)calculationMinIntervalNumericUpDown.Value;

            // appearances
            _visualEditor.FocusedCellsAppearance = focusedCellsAppearanceEditor.CellsAppearance;
            _visualEditor.CellsClipboardAppearance = bufferCellsAppearanceEditor.CellsAppearance;
            _visualEditor.FormulaAppearance = formulaCellsAppearanceEditor.CellsAppearance;
            _visualEditor.FormulaFocusedAppearance = focusedFormulaCellsAppearanceEditor.CellsAppearance;

            // headings
            _visualEditor.HeadingsColor = WpfObjectConverter.Convert(headingsColorPanelControl.Color);
            _visualEditor.HeadingsTextColor = WpfObjectConverter.Convert(headingsTextColorPanelControl.Color);
            _visualEditor.HeadingsBorderColor = WpfObjectConverter.Convert(headingsBorderColorPanelControl.Color);
            _visualEditor.SelectedCellColor = WpfObjectConverter.Convert(selectedCellColorPanelControl.Color);
            _visualEditor.SelectedHeaderColor = WpfObjectConverter.Convert(selectedHeaderColorPanelControl.Color);
            _visualEditor.CoveredHeaderColor = WpfObjectConverter.Convert(coveredHeaderColorPanelControl.Color);

            // errors
            _visualEditor.ErrorIndicatorColor = WpfObjectConverter.Convert(errorIndicatorColorPanelControl.Color);
            _visualEditor.ErrorIndicatorSize = (double)errorIndicatorSizeNumericUpDown.Value;

            // comments
            _visualEditor.CommentIndicatorColor = WpfObjectConverter.Convert(commentIndicatorColorPanelControl.Color);
            _visualEditor.CommentIndicatorSize = (double)commentIndicatorSizeNumericUpDown.Value;
            _visualEditor.CommentAppearance = commentAppearanceEditor.CellsAppearance;
            _visualEditor.CommentFocusedAppearance = focusedCommentAppearanceEditor.CellsAppearance;

            // miscellaneous
            _visualEditor.HyperlinkColor = WpfObjectConverter.Convert(hyperlinkColorPanelControl.Color);
            _visualEditor.GridColorAlpha = 1 - (double)gridColorAlphaNumericUpDown.Value / 255;

            DialogResult = true;
        }

        #endregion

    }
}

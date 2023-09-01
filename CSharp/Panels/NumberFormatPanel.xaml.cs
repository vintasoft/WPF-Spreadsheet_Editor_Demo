using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;

using Vintasoft.Imaging;
using Vintasoft.Imaging.Office.Spreadsheet.Document;
using Vintasoft.Imaging.Office.Spreadsheet.NumberFormats;
using Vintasoft.Imaging.Office.Spreadsheet.Wpf.UI;

namespace WpfSpreadsheetEditorDemo
{
    /// <summary>
    /// Provides a "Number Format" panel.
    /// </summary>
    public partial class NumberFormatPanel : SpreadsheetVisualEditorPanel
    {

        #region Fields

        /// <summary>
        /// The formats, which can be set using this panel.
        /// </summary>
        Dictionary<string, NumberFormat> _panelFormats = new Dictionary<string, NumberFormat>();

        /// <summary>
        /// Indicates whether the UI is currently updating.
        /// </summary>
        bool _isUpdateUI = false;

        #endregion



        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="NumberFormatPanel"/> class.
        /// </summary>
        public NumberFormatPanel()
        {
            InitializeComponent();

            // create formats for panel comboBox
            _panelFormats.Add("General", new GeneralFormat());
            _panelFormats.Add("Number", new NumberingFormat(2, false));
            _panelFormats.Add("Date", DateFormat.Create("mm-dd-yy"));
            _panelFormats.Add("Time", TimeFormat.Create("h:mm"));
            _panelFormats.Add("Currency", new CurrencyFormat(2, "[$$-en-US]", true));
            _panelFormats.Add("Percentage", new PercentageFormat(2));
            _panelFormats.Add("Scientific", new ScientificFormat(2));
            _panelFormats.Add("Text", new TextFormat());

            foreach (string formatName in _panelFormats.Keys)
                numberFormatComboBox.Items.Add(formatName);

            numberFormatComboBox.Items.Add("Custom");
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
                args.OldValue.VisualEditor.FocusedCellChanged -= VisualEditor_FocusedCellChanged;
                args.OldValue.VisualEditor.CellsStylePropertiesChanged -= VisualEditor_CellsStylePropertiesChanged;
                args.NewValue.VisualEditor.FocusedWorksheetChanged -= VisualEditor_FocusedWorksheetChanged;
            }
            if (args.NewValue != null)
            {
                args.NewValue.VisualEditor.FocusedCellChanged += VisualEditor_FocusedCellChanged;
                args.NewValue.VisualEditor.CellsStylePropertiesChanged += VisualEditor_CellsStylePropertiesChanged;
                args.NewValue.VisualEditor.FocusedWorksheetChanged += VisualEditor_FocusedWorksheetChanged;
            }
            UpdateUI();
        }

        private void VisualEditor_FocusedWorksheetChanged(object sender, PropertyChangedEventArgs<Worksheet> e)
        {
            UpdateUI();
        }

        private void VisualEditor_FocusedCellChanged(object sender, PropertyChangedEventArgs<CellReference> e)
        {
            UpdateUI();
        }

        private void VisualEditor_CellsStylePropertiesChanged(object sender, EventArgs e)
        {
            UpdateUI();
        }

        /// <summary>
        /// Updates the user interface of this panel.
        /// </summary>
        private void UpdateUI()
        {
            if (VisualEditor.IsFocusedWorksheetChanging)
                return;

            _isUpdateUI = true;

            try
            {

                if (VisualEditor.FocusedCell == null)
                {
                    IsEnabled = false;
                }
                else
                {
                    IsEnabled = true;

                    // get focused cell format
                    string formatString = VisualEditor.NumberFormat;
                    NumberFormat format = VisualEditor.Document.ParseNumberFormat(formatString);

                    // update comboBox value
                    if (format is GeneralFormat)
                        numberFormatComboBox.SelectedValue = "General";
                    else if (format is NumberingFormat)
                        numberFormatComboBox.SelectedValue = "Number";
                    else if (format is DateFormat)
                        numberFormatComboBox.SelectedValue = "Date";
                    else if (format is TimeFormat)
                        numberFormatComboBox.SelectedValue = "Time";
                    else if (format is CurrencyFormat)
                        numberFormatComboBox.SelectedValue = "Currency";
                    else if (format is PercentageFormat)
                        numberFormatComboBox.SelectedValue = "Percentage";
                    else if (format is ScientificFormat)
                        numberFormatComboBox.SelectedValue = "Scientific";
                    else if (format is TextFormat)
                        numberFormatComboBox.SelectedValue = "Text";
                    else
                        numberFormatComboBox.SelectedValue = "Custom";
                }
            }
            finally
            {
                _isUpdateUI = false;
            }
        }

        /// <summary>
        /// Handles the SelectionChanged event of NumberFormatComboBox object.
        /// </summary>
        private void numberFormatComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!_isUpdateUI)
            {
                string selectedFormat = numberFormatComboBox.SelectedValue.ToString();
                if (_panelFormats.ContainsKey(selectedFormat))
                {
                    // update cell number format
                    VisualEditor.NumberFormat = _panelFormats[selectedFormat].ToString(VisualEditor.Document.Defaults.FormattingProperties);
                }

                SpreadsheetEditor.Focus();
            }
        }

        /// <summary>
        /// Handles the Click event of EnglishUnitedStatesMenuItem object.
        /// </summary>
        private void englishUnitedStatesMenuItem_Click(object sender, RoutedEventArgs e)
        {
            CurrencyFormat format = new CurrencyFormat(2, "[$$-en-US]", true);
            VisualEditor.NumberFormat = format.ToString(VisualEditor.Document.Defaults.FormattingProperties);
            UpdateUI();
        }

        /// <summary>
        /// Handles the Click event of EnglishUnitedKingdomMenuItem object.
        /// </summary>
        private void englishUnitedKingdomMenuItem_Click(object sender, RoutedEventArgs e)
        {
            CurrencyFormat format = new CurrencyFormat(2, "[$£-en-GB]", true);
            VisualEditor.NumberFormat = format.ToString(VisualEditor.Document.Defaults.FormattingProperties);
            UpdateUI();
        }

        /// <summary>
        /// Handles the Click event of EuroMenuItem object.
        /// </summary>
        private void euroMenuItem_Click(object sender, RoutedEventArgs e)
        {
            CurrencyFormat format = new CurrencyFormat(2, "[$€-x-euro2]\\ ", true);
            VisualEditor.NumberFormat = format.ToString(VisualEditor.Document.Defaults.FormattingProperties);
            UpdateUI();
        }

        /// <summary>
        /// Handles the Click event of ChineseSimplifiedMenuItem object.
        /// </summary>
        private void chineseSimplifiedMenuItem_Click(object sender, RoutedEventArgs e)
        {
            CurrencyFormat format = new CurrencyFormat(2, "[$¥-zh-CN]", true);
            VisualEditor.NumberFormat = format.ToString(VisualEditor.Document.Defaults.FormattingProperties);
            UpdateUI();
        }

        /// <summary>
        /// Handles the Click event of RussianMenuItem object.
        /// </summary>
        private void russianMenuItem_Click(object sender, RoutedEventArgs e)
        {
            CurrencyFormat format = new CurrencyFormat(2, "\\ [$₽-ru-RU]", false);
            VisualEditor.NumberFormat = format.ToString(VisualEditor.Document.Defaults.FormattingProperties);
            UpdateUI();
        }

        /// <summary>
        /// Handles the Click event of EnglishIndiaMenuItem object.
        /// </summary>
        private void englishIndiaMenuItem_Click(object sender, RoutedEventArgs e)
        {
            CurrencyFormat format = new CurrencyFormat(2, "[$₹-en-IN]\\ ", true);
            VisualEditor.NumberFormat = format.ToString(VisualEditor.Document.Defaults.FormattingProperties);
            UpdateUI();
        }

        /// <summary>
        /// Handles the Click event of PercentStyleButton object.
        /// </summary>
        private void percentStyleButton_Click(object sender, RoutedEventArgs e)
        {
            VisualEditor.NumberFormat = "0%";
            UpdateUI();
        }

        /// <summary>
        /// Handles the Click event of NumberFormatPropertiesButton object.
        /// </summary>
        private void numberFormatPropertiesButton_Click(object sender, RoutedEventArgs e)
        {
            CellsStyleWindow.ShowNumberFormatDialog(VisualEditor);
            UpdateUI();
        }

        /// <summary>
        /// Handles the Click event of DecreaseDecimalButton object.
        /// </summary>
        private void decreaseDecimalButton_Click(object sender, RoutedEventArgs e)
        {
            VisualEditor.NumberFormatDecimalPlaces--;
            UpdateUI();
        }

        /// <summary>
        /// Handles the Click event of IncreaseDecimalButton object.
        /// </summary>
        private void increaseDecimalButton_Click(object sender, RoutedEventArgs e)
        {
            VisualEditor.NumberFormatDecimalPlaces++;
            UpdateUI();
        }

        #endregion

    }
}

﻿using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

using Vintasoft.Imaging.Fonts;
using Vintasoft.Imaging.Office.Spreadsheet;
using Vintasoft.Imaging.Office.Spreadsheet.Document;
using Vintasoft.Imaging.Office.Spreadsheet.NumberFormats;
using Vintasoft.Imaging.Office.Spreadsheet.UI;
using Vintasoft.Imaging.Wpf;
using Vintasoft.Primitives;

using WpfDemosCommonCode;

namespace WpfSpreadsheetEditorDemo
{
    /// <summary>
    /// A window that allows to view and change the style of cells.
    /// </summary>
    public partial class CellsStyleWindow : Window
    {

        #region Fields

        /// <summary>
        /// Current spreadsheet visual editor.
        /// </summary>
        SpreadsheetVisualEditor _visualEditor;

        /// <summary>
        /// A value indicating whether form initialization is finished.
        /// </summary>
        bool _initializationFinished = false;

        /// <summary>
        /// Dictionary that contains changed properties: property name => new property value.
        /// </summary>
        Dictionary<CellStyleProperty, object> _changedProperties = new Dictionary<CellStyleProperty, object>();

        /// <summary>
        /// The preview manager for cell borders.
        /// </summary>
        CellBordersPreviewManager _bordersPreview;

        /// <summary>
        /// Current number format.
        /// </summary>
        NumberFormat _currentFormat;

        /// <summary>
        /// Standard number formats.
        /// </summary>
        List<NumberFormat> _standartFormats = new List<NumberFormat>();

        /// <summary>
        /// Dictionary: format string => number format.
        /// </summary>
        Dictionary<string, NumberFormat> _formatStringToStandartFormat = new Dictionary<string, NumberFormat>();

        /// <summary>
        /// Dictionary: currency format name => currency format.
        /// </summary>
        Dictionary<string, CurrencyFormat> _currencyFormatNameToFormat = new Dictionary<string, CurrencyFormat>();

        #endregion



        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="CellsStyleWindow"/> class.
        /// </summary>
        public CellsStyleWindow()
        {
            InitializeComponent();

            // Number Format

            // add currency formats
            foreach (CurrencyFormat currencyFormat in SupportedNumberFormats.GetNumberFormats(NumberFormatCategory.Currency))
                _currencyFormatNameToFormat.Add(currencyFormat.Name, currencyFormat);

            // add date formats
            foreach (DateFormat dateFormat in SupportedNumberFormats.GetNumberFormats(NumberFormatCategory.Date))
                dateFormatsListBox.Items.Add(dateFormat.ToString(FormattingProperties));

            // add time formats
            foreach (TimeFormat timeFormat in SupportedNumberFormats.GetNumberFormats(NumberFormatCategory.Time))
                timeFormatsListBox.Items.Add(timeFormat.ToString(FormattingProperties));

            // add all supported formats
            _standartFormats = SupportedNumberFormats.GetNumberFormats(NumberFormatCategory.All);


            // Alignment
            foreach (TextHorizontalAlign horizontalAlignment in Enum.GetValues(typeof(TextHorizontalAlign)))
                textHorizontalAlignmentComboBox.Items.Add(horizontalAlignment);

            foreach (TextVerticalAlign verticalAlignment in Enum.GetValues(typeof(TextVerticalAlign)))
                textVerticalAlignmentComboBox.Items.Add(verticalAlignment);


            // Font
            foreach (string fontName in GetAvailableFontNames())
            {
                ListBoxItem item = new ListBoxItem();
                item.Content = fontName;
                fontNamesListBox.Items.Add(item);
            }
        }

        #endregion



        #region Properties

        /// <summary>
        /// Gets the current culture.
        /// </summary>
        public CultureInfo Culture
        {
            get
            {
                if (_visualEditor != null)
                {
                    try
                    {
                        return CultureInfo.GetCultureInfo(_visualEditor.DocumentCulture);
                    }
                    catch
                    {
                    }
                }
                return CultureInfo.CurrentCulture;
            }
        }

        /// <summary>
        /// Gets the formatting properties.
        /// </summary>
        public FormattingProperties FormattingProperties
        {
            get
            {
                if (_visualEditor != null)
                    return _visualEditor.Document.Defaults.FormattingProperties;

                return null;
            }
        }

        /// <summary>
        /// Gets or sets current cells borders.
        /// </summary>
        public CellsBorders CurrentBorders
        {
            get
            {
                return _bordersPreview.Borders;
            }
            set
            {
                _bordersPreview.Borders = value;
            }
        }

        #endregion



        #region Methods

        #region Number Format

        /// <summary>
        /// Shows this form with opened 'Number' tab.
        /// </summary>
        /// <param name="visualEditor">Spreadsheet visual editor.</param>
        public static bool? ShowNumberFormatDialog(SpreadsheetVisualEditor visualEditor)
        {
            CellsStyleWindow window = new CellsStyleWindow();
            window.WindowStartupLocation = WindowStartupLocation.CenterOwner;
            window.Owner = Application.Current.MainWindow;
            window.SetVisualEditor(visualEditor);
            window.cellStyleTabControl.SelectedItem = window.numberFormatTabPage;
            return window.ShowDialog();
        }

        /// <summary>
        /// Initializes the UI with number format properties.
        /// </summary>
        /// <param name="properties">Initial properties.</param>
        private void InitNumberFormatPropertiesUI(Dictionary<CellStyleProperty, object> commonProperties)
        {
            // fill the formats listBox
            foreach (NumberFormat numberFormat in _standartFormats)
            {
                string standartFormatString = numberFormat.ToString(FormattingProperties);
                customFormatsListBox.Items.Add(standartFormatString);
                _formatStringToStandartFormat.Add(standartFormatString, numberFormat);
            }

            // fill the currency symbol formats comboBox
            foreach (string currencyFormatName in _currencyFormatNameToFormat.Keys)
                currencySymbolComboBox.Items.Add(currencyFormatName);

            currencySymbolComboBox.SelectedItem = currencySymbolComboBox.Items[0];

            // if selection contains different formats
            if (!commonProperties.ContainsKey(CellStyleProperty.NumberFormat))
            {
                _currentFormat = null;
                return;
            }

            // get format string
            string formatString = (string)commonProperties[CellStyleProperty.NumberFormat];
            customFormatTextBox.Text = formatString;

            // get number format
            NumberFormat format = GetNumberingFormat(formatString);

            // select format tab page

            if (format is GeneralFormat)
            {
                formatCategoriesTabControl.SelectedItem = generalTabPage;
            }
            else if (format is NumberingFormat)
            {
                NumberingFormat numberingFormat = (NumberingFormat)format;
                formatCategoriesTabControl.SelectedItem = numberTabPage;
                numberDecimalPlacesNumericUpDown.Value = numberingFormat.DecimalPlaces;
                useThousandsSeparatorCheckBox.IsChecked = numberingFormat.UseThousandsSeparator;
                useRedColorForNegativeCheckBox.IsChecked = numberingFormat.NegativeValueColor == "Red";
                hideMinusSignCheckBox.IsChecked = numberingFormat.HideMinusSign;
            }
            else if (format is DateFormat)
            {
                DateFormat dateFormat = (DateFormat)format;
                formatCategoriesTabControl.SelectedItem = dateTabPage;
                dateFormatsListBox.SelectedItem = dateFormat.ToString(FormattingProperties);
            }
            else if (format is TimeFormat)
            {
                TimeFormat timeFormat = (TimeFormat)format;
                formatCategoriesTabControl.SelectedItem = timeTabPage;
                timeFormatsListBox.SelectedItem = timeFormat.ToString(FormattingProperties);
            }
            else if (format is CurrencyFormat)
            {
                CurrencyFormat currencyFormat = (CurrencyFormat)format;
                formatCategoriesTabControl.SelectedItem = currencyTabPage;
                currencyDecimalPlacesNumericUpDown.Value = currencyFormat.DecimalPlaces;
                if (currencySymbolComboBox.Items.Contains(currencyFormat.Name))
                    currencySymbolComboBox.SelectedItem = currencyFormat.Name;
            }
            else if (format is PercentageFormat)
            {
                PercentageFormat percentageFormat = (PercentageFormat)format;
                formatCategoriesTabControl.SelectedItem = percentageTabPage;
                percentageDecimalPlacesNumericUpDown.Value = percentageFormat.DecimalPlaces;
            }
            else if (format is ScientificFormat)
            {
                ScientificFormat scientificFormat = (ScientificFormat)format;
                formatCategoriesTabControl.SelectedItem = scientificTabPage;
                scientificDecimalPlacesNumericUpDown.Value = scientificFormat.DecimalPlaces;
            }
            else if (format is TextFormat)
            {
                formatCategoriesTabControl.SelectedItem = textTabPage;
            }
            else
            {
                formatCategoriesTabControl.SelectedItem = customTabPage;
                // add custom format to listBox
                if (!customFormatsListBox.Items.Contains(customFormatTextBox.Text))
                    customFormatsListBox.Items.Add(customFormatTextBox.Text);
                // set selected format in listBox
                customFormatsListBox.SelectedItem = customFormatTextBox.Text;
                customFormatsListBox.ScrollIntoView(customFormatsListBox.SelectedItem);
            }

            _currentFormat = format;
        }


        #region Format tab control

        /// <summary>
        /// Handles the SelectionChanged event of formatCategoriesTabControl object.
        /// </summary>
        private void formatCategoriesTabControl_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (e.OriginalSource != formatCategoriesTabControl)
                return;

            if (!_initializationFinished)
                return;

            NumberFormat format = _currentFormat;

            // set a new format from selected tab page

            if (formatCategoriesTabControl.SelectedItem == generalTabPage)
            {
                // create general format
                format = new GeneralFormat();
            }
            else if (formatCategoriesTabControl.SelectedItem == numberTabPage)
            {
                string negativeColor = null;

                if (useRedColorForNegativeCheckBox.IsChecked == true)
                    negativeColor = "Red";

                // create numbering format
                bool thousandsSeparator = useThousandsSeparatorCheckBox.IsChecked != null && useThousandsSeparatorCheckBox.IsChecked.Value == true;
                bool hideMinusSign = hideMinusSignCheckBox.IsChecked != null && hideMinusSignCheckBox.IsChecked.Value == true;
                format = new NumberingFormat(
                    (int)numberDecimalPlacesNumericUpDown.Value,
                    thousandsSeparator,
                    negativeColor,
                    hideMinusSign);
            }
            else if (formatCategoriesTabControl.SelectedItem == currencyTabPage)
            {
                CurrencyFormat currencyFormat = _currencyFormatNameToFormat[currencySymbolComboBox.SelectedItem.ToString()];
                // create currency format
                format = new CurrencyFormat(
                    (int)currencyDecimalPlacesNumericUpDown.Value,
                    currencyFormat.IsCurrencySymbolBeforeValue,
                    currencyFormat.CurrencySymbolFormat);
            }
            else if (formatCategoriesTabControl.SelectedItem == dateTabPage)
            {
                // create date format
                format = DateFormat.Create(dateFormatsListBox.Items[0].ToString());
                dateFormatsListBox.SelectedIndex = 0;
            }
            else if (formatCategoriesTabControl.SelectedItem == timeTabPage)
            {
                // create time format
                format = TimeFormat.Create(timeFormatsListBox.Items[0].ToString());
                timeFormatsListBox.SelectedIndex = 0;
            }
            else if (formatCategoriesTabControl.SelectedItem == percentageTabPage)
            {
                // create percentage format
                format = new PercentageFormat((int)percentageDecimalPlacesNumericUpDown.Value);
            }
            else if (formatCategoriesTabControl.SelectedItem == scientificTabPage)
            {
                // create scientific format
                format = new ScientificFormat((int)scientificDecimalPlacesNumericUpDown.Value);
            }
            else if (formatCategoriesTabControl.SelectedItem == textTabPage)
            {
                // create text format
                format = new TextFormat();
            }
            else if (formatCategoriesTabControl.SelectedItem == customTabPage)
            {
                if (_currentFormat != null)
                {
                    customFormatTextBox.Text = _currentFormat.ToString(FormattingProperties);
                    // add custom format to the listBox
                    if (!customFormatsListBox.Items.Contains(customFormatTextBox.Text))
                        customFormatsListBox.Items.Add(customFormatTextBox.Text);

                    // set selected format in listBox
                    customFormatsListBox.SelectedItem = customFormatTextBox.Text;
                    customFormatsListBox.ScrollIntoView(customFormatsListBox.SelectedItem);
                }
            }

            if (_currentFormat != null)
            {
                // update current format
                _currentFormat = format;
                // save changes about number format
                _changedProperties[CellStyleProperty.NumberFormat] = format.ToString(FormattingProperties);
            }
        }

        #endregion


        #region Number tab page

        /// <summary>
        /// Handles the ValueChanged event of numberDecimalPlacesNumericUpDown object.
        /// </summary>
        private void numberDecimalPlacesNumericUpDown_ValueChanged(object sender, EventArgs e)
        {
            if (!_initializationFinished)
                return;

            ((NumberingFormat)_currentFormat).DecimalPlaces = (int)numberDecimalPlacesNumericUpDown.Value;
            // save changes about number format
            _changedProperties[CellStyleProperty.NumberFormat] = _currentFormat.ToString(FormattingProperties);
        }

        /// <summary>
        /// Handles the CheckedChanged event of useThousandsSeparatorCheckBox object.
        /// </summary>
        private void useThousandsSeparatorCheckBox_CheckedChanged(object sender, RoutedEventArgs e)
        {
            if (!_initializationFinished)
                return;

            ((NumberingFormat)_currentFormat).UseThousandsSeparator = useThousandsSeparatorCheckBox.IsChecked != null && useThousandsSeparatorCheckBox.IsChecked.Value == true;
            // save changes about number format
            _changedProperties[CellStyleProperty.NumberFormat] = _currentFormat.ToString(FormattingProperties);
        }

        /// <summary>
        /// Handles the CheckedChanged event of useRedColorForNegativeCheckBox object.
        /// </summary>
        private void useRedColorForNegativeCheckBox_CheckedChanged(object sender, RoutedEventArgs e)
        {
            if (!_initializationFinished)
                return;

            if (useRedColorForNegativeCheckBox.IsChecked == true)
                ((NumberingFormat)_currentFormat).NegativeValueColor = "Red";
            else
                ((NumberingFormat)_currentFormat).NegativeValueColor = null;

            // save changes about number format
            _changedProperties[CellStyleProperty.NumberFormat] = _currentFormat.ToString(FormattingProperties);
        }

        /// <summary>
        /// Handles the CheckedChanged event of hideMinusSignCheckBox object.
        /// </summary>
        private void hideMinusSignCheckBox_CheckedChanged(object sender, RoutedEventArgs e)
        {
            if (!_initializationFinished)
                return;

            ((NumberingFormat)_currentFormat).HideMinusSign = hideMinusSignCheckBox.IsChecked != null && hideMinusSignCheckBox.IsChecked.Value == true;

            // save changes about number format
            _changedProperties[CellStyleProperty.NumberFormat] = _currentFormat.ToString(FormattingProperties);
        }

        #endregion


        #region Currency tab page

        /// <summary>
        /// Handles the ValueChanged event of currencyDecimalPlacesNumericUpDown object.
        /// </summary>
        private void currencyDecimalPlacesNumericUpDown_ValueChanged(object sender, EventArgs e)
        {
            if (!_initializationFinished)
                return;

            ((CurrencyFormat)_currentFormat).DecimalPlaces = (int)currencyDecimalPlacesNumericUpDown.Value;
            // save changes about number format
            _changedProperties[CellStyleProperty.NumberFormat] = _currentFormat.ToString(FormattingProperties);
        }

        /// <summary>
        /// Handles the SelectionChanged event of currencySymbolComboBox object.
        /// </summary>
        private void currencySymbolComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!_initializationFinished)
                return;

            // get currency symbol format
            CurrencyFormat currencyFormat = _currencyFormatNameToFormat[currencySymbolComboBox.SelectedItem.ToString()];

            CurrencyFormat currentCurrencyFormat = _currentFormat as CurrencyFormat;
            // create format with new currency symbol
            _currentFormat = new CurrencyFormat(currentCurrencyFormat.DecimalPlaces, currencyFormat.IsCurrencySymbolBeforeValue, currencyFormat.CurrencySymbolFormat);
            // save changes about number format
            _changedProperties[CellStyleProperty.NumberFormat] = _currentFormat.ToString(FormattingProperties);
        }

        #endregion


        #region Date tab page

        /// <summary>
        /// Handles the SelectionChanged event of dateFormatsListBox object.
        /// </summary>
        private void dateFormatsListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!_initializationFinished)
                return;

            // create date format
            _currentFormat = DateFormat.Create(dateFormatsListBox.SelectedItem.ToString());

            if (_currentFormat != null)
            {
                // save changes about number format
                _changedProperties[CellStyleProperty.NumberFormat] = _currentFormat.ToString(FormattingProperties);
            }
        }

        #endregion


        #region Time tab page

        /// <summary>
        /// Handles the SelectionChanged event of timeFormatsListBox object.
        /// </summary>
        private void timeFormatsListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!_initializationFinished)
                return;

            // create time format
            _currentFormat = TimeFormat.Create(timeFormatsListBox.SelectedItem.ToString());

            if (_currentFormat != null)
            {
                // save changes about number format
                _changedProperties[CellStyleProperty.NumberFormat] = _currentFormat.ToString(FormattingProperties);
            }
        }

        #endregion


        #region Percentage tab page

        /// <summary>
        /// Handles the ValueChanged event of percentageDecimalPlacesNumericUpDown object.
        /// </summary>
        private void percentageDecimalPlacesNumericUpDown_ValueChanged(object sender, EventArgs e)
        {
            if (!_initializationFinished)
                return;

            ((PercentageFormat)_currentFormat).DecimalPlaces = (int)percentageDecimalPlacesNumericUpDown.Value;
            // save changes about number format
            _changedProperties[CellStyleProperty.NumberFormat] = _currentFormat.ToString(FormattingProperties);
        }

        #endregion


        #region Scientific tab page

        /// <summary>
        /// Handles the ValueChanged event of scientificDecimalPlacesNumericUpDown object.
        /// </summary>
        private void scientificDecimalPlacesNumericUpDown_ValueChanged(object sender, EventArgs e)
        {
            if (!_initializationFinished)
                return;

            ((ScientificFormat)_currentFormat).DecimalPlaces = (int)scientificDecimalPlacesNumericUpDown.Value;
            // save changes about number format
            _changedProperties[CellStyleProperty.NumberFormat] = _currentFormat.ToString(FormattingProperties);
        }

        #endregion


        #region Custom tab page

        /// <summary>
        /// Handles the SelectionChanged event of customFormatsListBox object.
        /// </summary>
        private void customFormatsListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!_initializationFinished)
                return;

            if (customFormatsListBox.SelectedItem == null)
                return;

            customFormatTextBox.Text = customFormatsListBox.SelectedItem.ToString();
            // get selected format
            _currentFormat = _visualEditor.Document.ParseNumberFormat(customFormatTextBox.Text);
            // save changes about number format
            _changedProperties[CellStyleProperty.NumberFormat] = _currentFormat.ToString(FormattingProperties);

            customFormatsListBox.ScrollIntoView(customFormatsListBox.SelectedItem);
        }

        /// <summary>
        /// Handles the LostFocus event of customFormatTextBox object.
        /// </summary>
        private void customFormatTextBox_LostFocus(object sender, RoutedEventArgs e)
        {
            // get format
            _currentFormat = GetNumberingFormat(customFormatTextBox.Text);
            // save changes about number format
            _changedProperties[CellStyleProperty.NumberFormat] = _currentFormat.ToString(FormattingProperties);
        }

        /// <summary>
        /// Handles the PreviewKeyDown event of customFormatTextBox object.
        /// </summary>
        private void customFormatTextBox_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                // get format
                _currentFormat = GetNumberingFormat(customFormatTextBox.Text);
                // save changes about number format
                _changedProperties[CellStyleProperty.NumberFormat] = _currentFormat.ToString(FormattingProperties);
            }
        }

        #endregion


        /// <summary>
        /// Returns number format, which is created from specified format string.
        /// </summary>
        /// <param name="formatString">The format string.</param>
        /// <returns>The number format.</returns>
        private NumberFormat GetNumberingFormat(string formatString)
        {
            // if format string corresponds to one of the standart formats
            if (_formatStringToStandartFormat.ContainsKey(formatString))
                return _formatStringToStandartFormat[formatString];

            // parse number format
            return _visualEditor.Document.ParseNumberFormat(formatString);
        }

        #endregion


        #region Alignment

        /// <summary>
        /// Shows this form with opened 'Alignment' tab.
        /// </summary>
        /// <param name="visualEditor">Spreadsheet visual editor.</param>
        public static bool? ShowAlignmentDialog(SpreadsheetVisualEditor visualEditor)
        {
            CellsStyleWindow window = new CellsStyleWindow();
            window.WindowStartupLocation = WindowStartupLocation.CenterOwner;
            window.Owner = Application.Current.MainWindow;
            window.SetVisualEditor(visualEditor);
            window.cellStyleTabControl.SelectedItem = window.alignmentTabPage;
            return window.ShowDialog();
        }

        /// <summary>
        /// Initializes the UI with alignment properties.
        /// </summary>
        /// <param name="properties">Initial properties.</param>
        private void InitAlignmentPropertiesUI(Dictionary<CellStyleProperty, object> properties)
        {
            // init text horizontal alignment
            if (properties.ContainsKey(CellStyleProperty.TextHorizontalAlign))
            {
                textHorizontalAlignmentComboBox.SelectedItem =
                    properties[CellStyleProperty.TextHorizontalAlign];
            }

            // init text vertical alignment
            if (properties.ContainsKey(CellStyleProperty.TextVerticalAlign))
            {
                textVerticalAlignmentComboBox.SelectedItem =
                    properties[CellStyleProperty.TextVerticalAlign];
            }

            // init text indent
            if (properties.ContainsKey(CellStyleProperty.TextIndent))
                textIndentNumericUpDown.Value = (int)properties[CellStyleProperty.TextIndent];
            else
                textIndentNumericUpDown.Value = 0;

            // init text wrap
            if (properties.ContainsKey(CellStyleProperty.TextWordWrap))
            {
                wrapTextCheckBox.IsChecked = (bool)properties[CellStyleProperty.TextWordWrap];
            }
            else
            {
                wrapTextCheckBox.IsThreeState = true;
                wrapTextCheckBox.IsChecked = null;
            }
        }

        /// <summary>
        /// Handles the SelectionChanged event of textHorizontalAlignmentComboBox object.
        /// </summary>
        private void textHorizontalAlignmentComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!_initializationFinished)
                return;

            if (textHorizontalAlignmentComboBox.SelectedItem != null)
            {
                // save changes about horizontal alignment
                _changedProperties[CellStyleProperty.TextHorizontalAlign] =
                    (TextHorizontalAlign)textHorizontalAlignmentComboBox.SelectedIndex;

                if ((TextHorizontalAlign)textHorizontalAlignmentComboBox.SelectedItem == TextHorizontalAlign.Default ||
                    (TextHorizontalAlign)textHorizontalAlignmentComboBox.SelectedItem == TextHorizontalAlign.Center)
                {
                    // reset the text indent
                    _changedProperties[CellStyleProperty.TextIndent] = 0;
                    textIndentNumericUpDown.Value = 0;
                }
            }
        }

        /// <summary>
        /// Handles the SelectionChanged event of textVerticalAlignmentComboBox object.
        /// </summary>
        private void textVerticalAlignmentComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!_initializationFinished)
                return;

            if (textVerticalAlignmentComboBox.SelectedItem != null)
            {
                // save changes about vertical alignment
                _changedProperties[CellStyleProperty.TextVerticalAlign] =
                    (TextVerticalAlign)textVerticalAlignmentComboBox.SelectedIndex;
            }
        }

        /// <summary>
        /// Handles the ValueChanged event of textIndentNumericUpDown object.
        /// </summary>
        private void textIndentNumericUpDown_ValueChanged(object sender, EventArgs e)
        {
            if (!_initializationFinished)
                return;

            // save changes about text indent
            SetIndentValue();
        }

        /// <summary>
        /// Handles the CheckStateChanged event of wrapTextCheckBox object.
        /// </summary>
        private void wrapTextCheckBox_CheckStateChanged(object sender, RoutedEventArgs e)
        {
            if (!_initializationFinished)
                return;

            bool isWordWrapIndeterminated = wrapTextCheckBox.IsChecked == null;

            if (isWordWrapIndeterminated)
            {
                // clear changes about word wrap
                _changedProperties.Remove(CellStyleProperty.TextWordWrap);
            }
            else
            {
                // save changes about word wrap
                _changedProperties[CellStyleProperty.TextWordWrap] = wrapTextCheckBox.IsChecked == true;
            }
        }

        /// <summary>
        /// Sets the text indent from the corresponding control.
        /// </summary>
        private void SetIndentValue()
        {
            // save changes about text indent
            _changedProperties[CellStyleProperty.TextIndent] = (int)textIndentNumericUpDown.Value;

            if (textIndentNumericUpDown.Value > 0)
            {
                // if text alignment does not support text indent
                if (textHorizontalAlignmentComboBox.SelectedItem == null ||
                    (TextHorizontalAlign)textHorizontalAlignmentComboBox.SelectedItem == TextHorizontalAlign.Default ||
                    (TextHorizontalAlign)textHorizontalAlignmentComboBox.SelectedItem == TextHorizontalAlign.Center)
                {
                    textHorizontalAlignmentComboBox.SelectedItem = TextHorizontalAlign.Left;
                }
            }
        }

        #endregion


        #region Font

        /// <summary>
        /// Shows this form with opened 'Font' tab.
        /// </summary>
        /// <param name="visualEditor">Spreadsheet visual editor.</param>
        public static bool? ShowFontDialog(SpreadsheetVisualEditor visualEditor)
        {
            CellsStyleWindow window = new CellsStyleWindow();
            window.WindowStartupLocation = WindowStartupLocation.CenterOwner;
            window.Owner = Application.Current.MainWindow;
            window.SetVisualEditor(visualEditor);
            window.cellStyleTabControl.SelectedItem = window.fontTabPage;
            return window.ShowDialog();
        }

        /// <summary>
        /// Initializes the UI with font properties.
        /// </summary>
        /// <param name="properties">Initial properties.</param>
        private void InitFontPropertiesUI(Dictionary<CellStyleProperty, object> properties)
        {
            // init font name
            if (properties.ContainsKey(CellStyleProperty.FontName))
            {
                fontNameTextBox.Text = (string)properties[CellStyleProperty.FontName];
                DemosTools.SetSelectedItem(fontNamesListBox, fontNameTextBox.Text);
            }

            // init font style
            bool? isBold = null;
            bool? isItalic = null;
            if (properties.ContainsKey(CellStyleProperty.FontIsBold))
                isBold = (bool)properties[CellStyleProperty.FontIsBold];
            if (properties.ContainsKey(CellStyleProperty.FontIsItalic))
                isItalic = (bool)properties[CellStyleProperty.FontIsItalic];

            if (isBold.HasValue && isItalic.HasValue)
            {
                if (isBold.Value && isItalic.Value)
                {
                    fontStyleTextBox.Text = "Bold Italic";
                    DemosTools.SetSelectedItem(fontStylesListBox, "Bold Italic");
                }
                else if (isBold.Value)
                {
                    fontStyleTextBox.Text = "Bold";
                    DemosTools.SetSelectedItem(fontStylesListBox, "Bold");
                }
                else if (isItalic.Value)
                {
                    fontStyleTextBox.Text = "Italic";
                    DemosTools.SetSelectedItem(fontStylesListBox, "Italic");
                }
                else
                {
                    fontStyleTextBox.Text = "Regular";
                    DemosTools.SetSelectedItem(fontStylesListBox, "Regular");
                }
            }

            // init font size
            if (properties.ContainsKey(CellStyleProperty.FontSize))
            {
                double fontSize = (double)properties[CellStyleProperty.FontSize];
                fontSizeTextBox.Text = fontSize.ToString(Culture);
                DemosTools.SetSelectedItem(fontSizesListBox, fontSizeTextBox.Text);
            }

            // init text underline
            if (properties.ContainsKey(CellStyleProperty.FontIsUnderline))
            {
                underlineCheckBox.IsChecked = (bool)properties[CellStyleProperty.FontIsUnderline];
            }
            else
            {
                underlineCheckBox.IsThreeState = true;
                underlineCheckBox.IsChecked = null;
            }

            // init text strikeout
            if (properties.ContainsKey(CellStyleProperty.FontIsStrikeout))
            {
                strikethroughCheckBox.IsChecked = (bool)properties[CellStyleProperty.FontIsStrikeout];
            }
            else
            {
                strikethroughCheckBox.IsThreeState = true;
                strikethroughCheckBox.IsChecked = null;
            }

            // init font color
            if (properties.ContainsKey(CellStyleProperty.FontColor))
            {
                VintasoftColor fontColor = (VintasoftColor)properties[CellStyleProperty.FontColor];
                fontColorPanelControl.Color = WpfObjectConverter.Convert(fontColor);
            }
            else
            {
                fontColorPanelControl.Color = new Color();
            }
        }

        /// <summary>
        /// Handles the SelectionChanged event of fontNamesListBox object.
        /// </summary>
        private void fontNamesListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!_initializationFinished)
                return;

            if (fontNamesListBox.SelectedItem != null)
            {
                fontNameTextBox.Text = ((ListBoxItem)fontNamesListBox.SelectedItem).Content.ToString();
                SetFontNameValue();
            }
        }

        /// <summary>
        /// Handles the LostFocus event of fontNameTextBox object.
        /// </summary>
        private void fontNameTextBox_LostFocus(object sender, RoutedEventArgs e)
        {
            SetFontNameValue();

            DemosTools.SetSelectedItem(fontNamesListBox, fontNameTextBox.Text);
        }

        /// <summary>
        /// Handles the PreviewKeyDown event of fontNameTextBox object.
        /// </summary>
        private void fontNameTextBox_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
                SetFontNameValue();
        }

        /// <summary>
        /// Handles the SelectionChanged event of fontStylesListBox object.
        /// </summary>
        private void fontStylesListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!_initializationFinished)
                return;

            if (fontStylesListBox.SelectedItem != null)
            {
                fontStyleTextBox.Text = ((ListBoxItem)fontStylesListBox.SelectedItem).Content.ToString();
                SetFontStyleValue();
            }
        }

        /// <summary>
        /// Handles the LostFocus event of fontStyleTextBox object.
        /// </summary>
        private void fontStyleTextBox_LostFocus(object sender, RoutedEventArgs e)
        {
            SetFontStyleValue();

            DemosTools.SetSelectedItem(fontStylesListBox, fontStyleTextBox.Text);
        }

        /// <summary>
        /// Handles the PreviewKeyDown event of fontStyleTextBox object.
        /// </summary>
        private void fontStyleTextBox_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
                SetFontStyleValue();
        }

        /// <summary>
        /// Handles the SelectionChanged event of fontSizesListBox object.
        /// </summary>
        private void fontSizesListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!_initializationFinished)
                return;

            if (fontSizesListBox.SelectedItem != null)
            {
                fontSizeTextBox.Text = ((ListBoxItem)fontSizesListBox.SelectedItem).Content.ToString();
                SetFontSizeValue();
            }
        }

        /// <summary>
        /// Handles the LostFocus event of fontSizeTextBox object.
        /// </summary>
        private void fontSizeTextBox_LostFocus(object sender, RoutedEventArgs e)
        {
            SetFontSizeValue();

            DemosTools.SetSelectedItem(fontSizesListBox, fontSizeTextBox.Text);
        }

        /// <summary>
        /// Handles the PreviewKeyDown event of fontSizeTextBox object.
        /// </summary>
        private void fontSizeTextBox_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
                SetFontSizeValue();
        }

        /// <summary>
        /// Handles the CheckStateChanged event of underlineCheckBox object.
        /// </summary>
        private void underlineCheckBox_CheckStateChanged(object sender, RoutedEventArgs e)
        {
            if (!_initializationFinished)
                return;

            bool isUnderlineIndeterminated = underlineCheckBox.IsChecked == null;

            if (isUnderlineIndeterminated)
            {
                // clear changes about underline state
                _changedProperties.Remove(CellStyleProperty.FontIsUnderline);
            }
            else
            {
                // save changes about underline state
                _changedProperties[CellStyleProperty.FontIsUnderline] = underlineCheckBox.IsChecked != null && underlineCheckBox.IsChecked.Value == true;
            }
        }

        /// <summary>
        /// Handles the CheckStateChanged event of strikethroughCheckBox object.
        /// </summary>
        private void strikethroughCheckBox_CheckStateChanged(object sender, RoutedEventArgs e)
        {
            if (!_initializationFinished)
                return;

            bool isStrikethroughIndeterminated = strikethroughCheckBox.IsChecked == null;

            if (isStrikethroughIndeterminated)
            {
                // clear changes about strikethrough state
                _changedProperties.Remove(CellStyleProperty.FontIsStrikeout);
            }
            else
            {
                // save changes about strikethrough state
                _changedProperties[CellStyleProperty.FontIsStrikeout] = strikethroughCheckBox.IsChecked == true;
            }
        }

        /// <summary>
        /// Handles the ColorChanged event of fontColorPanelControl object.
        /// </summary>
        private void fontColorPanelControl_ColorChanged(object sender, EventArgs e)
        {
            if (!_initializationFinished)
                return;

            _changedProperties[CellStyleProperty.FontColor] = WpfObjectConverter.Convert(fontColorPanelControl.Color);
        }

        /// <summary>
        /// Handles the Click event of normalFontButton object.
        /// </summary>
        private void normalFontButton_Click(object sender, RoutedEventArgs e)
        {
            if (!_initializationFinished)
                return;

            // get the "Normal" style from spreadsheet document
            CellStyle normalStyle = _visualEditor.Document.Styles[0];

            // set the UI to the "Normal" style
            // style properties will change in event handlers

            DemosTools.SetSelectedItem(fontNamesListBox, normalStyle.FontProperties.Name);

            if (normalStyle.FontProperties.IsBold && normalStyle.FontProperties.IsItalic)
                DemosTools.SetSelectedItem(fontStylesListBox, "Bold Italic");
            else if (normalStyle.FontProperties.IsBold)
                DemosTools.SetSelectedItem(fontStylesListBox, "Bold");
            else if (normalStyle.FontProperties.IsItalic)
                DemosTools.SetSelectedItem(fontStylesListBox, "Italic");
            else
                DemosTools.SetSelectedItem(fontStylesListBox, "Regular");

            fontSizeTextBox.Text = normalStyle.FontProperties.Size.ToString(Culture);
            SetFontSizeValue();
            DemosTools.SetSelectedItem(fontSizesListBox, fontSizeTextBox.Text);

            if (normalStyle.FontProperties.IsUnderline)
            {
                underlineCheckBox.IsChecked = true;
            }
            else
            {
                underlineCheckBox.IsChecked = false;
            }

            if (normalStyle.FontProperties.IsStrikeout)
            {
                strikethroughCheckBox.IsChecked = true;
            }
            else
            {
                strikethroughCheckBox.IsChecked = false;
            }

            _changedProperties[CellStyleProperty.FontColor] = normalStyle.FontProperties.Color;
            fontColorPanelControl.Color = WpfObjectConverter.Convert(normalStyle.FontProperties.Color);
        }

        /// <summary>
        /// Sets the font name from the corresponding text box.
        /// </summary>
        private void SetFontNameValue()
        {
            if (!_initializationFinished)
                return;

            _changedProperties[CellStyleProperty.FontName] = fontNameTextBox.Text;
        }

        /// <summary>
        /// Sets the font style from the corresponding text box.
        /// </summary>
        private void SetFontStyleValue()
        {
            if (!_initializationFinished)
                return;

            if (string.IsNullOrEmpty(fontStyleTextBox.Text))
                return;

            if (DemosTools.Contains(fontStylesListBox, fontStyleTextBox.Text))
            {
                bool isBold = fontStyleTextBox.Text == "Bold" || fontStyleTextBox.Text == "Bold Italic";
                bool isItalic = fontStyleTextBox.Text == "Italic" || fontStyleTextBox.Text == "Bold Italic";

                _changedProperties[CellStyleProperty.FontIsBold] = isBold;
                _changedProperties[CellStyleProperty.FontIsItalic] = isItalic;
            }
            else
            {
                MessageBox.Show(string.Format("Font style with name '{0}' does not exist.", fontStyleTextBox.Text));
            }
        }

        /// <summary>
        /// Sets the font size from the corresponding text box.
        /// </summary>
        private void SetFontSizeValue()
        {
            if (!_initializationFinished)
                return;

            if (!string.IsNullOrEmpty(fontSizeTextBox.Text))
            {
                double fontSize;
                if (double.TryParse(fontSizeTextBox.Text, NumberStyles.Any, Culture, out fontSize))
                    _changedProperties[CellStyleProperty.FontSize] = Math.Round(fontSize, 2);
                else
                    MessageBox.Show("Font size must be an integer or decimal number.");
            }
        }

        /// <summary>
        /// Returns an array of avaliable font names.
        /// </summary>
        public static string[] GetAvailableFontNames()
        {
            FileFontProgramsController fontProgramsController = 
                (FileFontProgramsController)Vintasoft.Imaging.Drawing.DrawingFactory.Default.FontProgramsController;
            Dictionary<string, string> systemFonts = fontProgramsController.GetSystemInstalledFonts();
            List<string> result = new List<string>();
            foreach (string fontName in systemFonts.Keys)
            {
                if (fontName.ToUpperInvariant().Contains(" BOLD") || fontName.ToUpperInvariant().Contains(" ITALIC"))
                    continue;
                result.Add(fontName);
            }
            return result.ToArray();
        }

        #endregion


        #region Border

        /// <summary>
        /// Shows this form with opened 'Border' tab.
        /// </summary>
        /// <param name="visualEditor">Spreadsheet visual editor.</param>
        public static bool? ShowBordersDialog(SpreadsheetVisualEditor visualEditor)
        {
            CellsStyleWindow window = new CellsStyleWindow();
            window.WindowStartupLocation = WindowStartupLocation.CenterOwner;
            window.Owner = Application.Current.MainWindow;
            window.SetVisualEditor(visualEditor);
            window.cellStyleTabControl.SelectedItem = window.bordersTabPage;
            return window.ShowDialog();
        }

        /// <summary>
        /// Initializes the UI with border properties.
        /// </summary>
        /// <param name="properties">Initial properties.</param>
        /// <param name="styleIndexes">Cells style indexes.</param>
        private void InitBorderPropertiesUI(CellReferences[] selectedCells)
        {
            lineStylesListBox.SelectedIndex = 6;
            lineColorPanelControl.Color = Colors.Black;

            // init borders preview control
            bordersPreviewControl.AutoScroll = false;
            bordersPreviewControl.IsEnabled = false;

            // create the preview manager for cell borders
            _bordersPreview = new CellBordersPreviewManager(bordersPreviewControl.VisualEditor, selectedCells);
            // set borders to the border preview
            CurrentBorders = _visualEditor.CellsBorders;

            // update border buttons
            verticalBorderButton.IsEnabled = _bordersPreview.CanEditVerticalBorder;
            horizontalBorderButton.IsEnabled = _bordersPreview.CanEditHorizontalBorder;
            insideBorderPresetButton.IsEnabled = _bordersPreview.CanEditVerticalBorder || _bordersPreview.CanEditHorizontalBorder;
        }

        /// <summary>
        /// Handles the Click event of noneBorderPresetButton object.
        /// </summary>
        private void noneBorderPresetButton_Click(object sender, RoutedEventArgs e)
        {
            CurrentBorders = new CellsBorders(new CellBorders(CellBorder.Invisible), CellBorder.Invisible, CellBorder.Invisible);
            _changedProperties[CellStyleProperty.Borders] = CurrentBorders;
        }

        /// <summary>
        /// Handles the Click event of outlineBorderPresetButton object.
        /// </summary>
        private void outlineBorderPresetButton_Click(object sender, RoutedEventArgs e)
        {
            CellBorder border = GetSelectedBorder();
            CurrentBorders = new CellsBorders(new CellBorders(border), CurrentBorders.HorizontalBorder, CurrentBorders.VerticalBorder);
            _changedProperties[CellStyleProperty.Borders] = CurrentBorders;
        }

        /// <summary>
        /// Handles the Click event of insideBorderPresetButton object.
        /// </summary>
        private void insideBorderPresetButton_Click(object sender, RoutedEventArgs e)
        {
            CellBorder border = GetSelectedBorder();
            CurrentBorders = new CellsBorders(CurrentBorders.OutsideBorders, border, border);
            _changedProperties[CellStyleProperty.Borders] = CurrentBorders;
        }

        /// <summary>
        /// Handles the Click event of topBorderButton object.
        /// </summary>
        private void topBorderButton_Click(object sender, RoutedEventArgs e)
        {
            CellBorder border = GetSelectedBorder();

            if (Equals(CurrentBorders.OutsideBorders.Top, border))
                border = CellBorder.Invisible;

            CellBorders borders = new CellBorders(
                CurrentBorders.OutsideBorders.Left,
                CurrentBorders.OutsideBorders.Right,
                border,
                CurrentBorders.OutsideBorders.Bottom);

            CurrentBorders = new CellsBorders(borders, CurrentBorders.HorizontalBorder, CurrentBorders.VerticalBorder);
            _changedProperties[CellStyleProperty.Borders] = CurrentBorders;
        }

        /// <summary>
        /// Handles the Click event of horizontalBorderButton object.
        /// </summary>
        private void horizontalBorderButton_Click(object sender, RoutedEventArgs e)
        {
            CellBorder border = GetSelectedBorder();

            if (Equals(CurrentBorders.HorizontalBorder, border))
                border = CellBorder.Invisible;

            CurrentBorders = new CellsBorders(CurrentBorders.OutsideBorders, border, CurrentBorders.VerticalBorder);
            _changedProperties[CellStyleProperty.Borders] = CurrentBorders;
        }

        /// <summary>
        /// Handles the Click event of bottomBorderButton object.
        /// </summary>
        private void bottomBorderButton_Click(object sender, RoutedEventArgs e)
        {
            CellBorder border = GetSelectedBorder();

            if (Equals(CurrentBorders.OutsideBorders.Bottom, border))
                border = CellBorder.Invisible;

            CellBorders borders = new CellBorders(
                CurrentBorders.OutsideBorders.Left,
                CurrentBorders.OutsideBorders.Right,
                CurrentBorders.OutsideBorders.Top,
                border);

            CurrentBorders = new CellsBorders(borders, CurrentBorders.HorizontalBorder, CurrentBorders.VerticalBorder);
            _changedProperties[CellStyleProperty.Borders] = CurrentBorders;
        }

        /// <summary>
        /// Handles the Click event of leftBorderButton object.
        /// </summary>
        private void leftBorderButton_Click(object sender, RoutedEventArgs e)
        {
            CellBorder border = GetSelectedBorder();

            if (Equals(CurrentBorders.OutsideBorders.Left, border))
                border = CellBorder.Invisible;

            CellBorders borders = new CellBorders(
                border,
                CurrentBorders.OutsideBorders.Right,
                CurrentBorders.OutsideBorders.Top,
                CurrentBorders.OutsideBorders.Bottom);

            CurrentBorders = new CellsBorders(borders, CurrentBorders.HorizontalBorder, CurrentBorders.VerticalBorder);
            _changedProperties[CellStyleProperty.Borders] = CurrentBorders;
        }

        /// <summary>
        /// Handles the Click event of verticalBorderButton object.
        /// </summary>
        private void verticalBorderButton_Click(object sender, RoutedEventArgs e)
        {
            CellBorder border = GetSelectedBorder();

            if (Equals(CurrentBorders.VerticalBorder, border))
                border = CellBorder.Invisible;

            CurrentBorders = new CellsBorders(CurrentBorders.OutsideBorders, CurrentBorders.HorizontalBorder, border);
            _changedProperties[CellStyleProperty.Borders] = CurrentBorders;
        }

        /// <summary>
        /// Handles the Click event of rightBorderButton object.
        /// </summary>
        private void rightBorderButton_Click(object sender, RoutedEventArgs e)
        {
            CellBorder border = GetSelectedBorder();

            if (Equals(CurrentBorders.OutsideBorders.Right, border))
                border = CellBorder.Invisible;

            CellBorders borders = new CellBorders(
                CurrentBorders.OutsideBorders.Left,
                border,
                CurrentBorders.OutsideBorders.Top,
                CurrentBorders.OutsideBorders.Bottom);

            CurrentBorders = new CellsBorders(borders, CurrentBorders.HorizontalBorder, CurrentBorders.VerticalBorder);
            _changedProperties[CellStyleProperty.Borders] = CurrentBorders;
        }

        /// <summary>
        /// Returns the border with selected border style and color.
        /// </summary>
        private CellBorder GetSelectedBorder()
        {
            CellBorderStyle style = CellBorderStyle.None;

            string selectedItemText = null;
            if (lineStylesListBox.SelectedItem != null)
                selectedItemText = ((ListBoxItem)lineStylesListBox.SelectedItem).Content.ToString();

            switch (selectedItemText)
            {
                case "Hair":
                    style = CellBorderStyle.Hair;
                    break;
                case "Dotted":
                    style = CellBorderStyle.Dotted;
                    break;
                case "Dash Dot Dot":
                    style = CellBorderStyle.DashDotDot;
                    break;
                case "Dash Dot":
                    style = CellBorderStyle.DashDot;
                    break;
                case "Dashed":
                    style = CellBorderStyle.Dashed;
                    break;
                case "Thin":
                    style = CellBorderStyle.Thin;
                    break;
                case "Medium Dash Dot Dot":
                    style = CellBorderStyle.MediumDashDotDot;
                    break;
                case "Medium Dash Dot":
                    style = CellBorderStyle.MediumDashDot;
                    break;
                case "Medium Dashed":
                    style = CellBorderStyle.MediumDashed;
                    break;
                case "Medium":
                    style = CellBorderStyle.Medium;
                    break;
                case "Thick":
                    style = CellBorderStyle.Thick;
                    break;
                case "Double":
                    style = CellBorderStyle.Double;
                    break;

                case "None":
                    style = CellBorderStyle.None;
                    break;
            }

            if (style == CellBorderStyle.None)
                return CellBorder.Invisible;

            return new CellBorder(style, WpfObjectConverter.Convert(lineColorPanelControl.Color));
        }

        #endregion


        #region Fill

        /// <summary>
        /// Initializes the UI with fill properties.
        /// </summary>
        /// <param name="properties">Initial properties.</param>
        /// <param name="styleIndexes">Cells style indexes.</param>
        private void InitFillPropertiesUI(Dictionary<CellStyleProperty, object> properties)
        {
            object fillColor = null;
            if (properties.TryGetValue(CellStyleProperty.FillColor, out fillColor))
                backgroundColorPanelControl.Color = WpfObjectConverter.Convert((VintasoftColor)fillColor);
        }

        /// <summary>
        /// Handles the ColorChanged event of backgroundColorPanelControl object.
        /// </summary>
        private void backgroundColorPanelControl_ColorChanged(object sender, EventArgs e)
        {
            _changedProperties[CellStyleProperty.FillColor] = WpfObjectConverter.Convert(backgroundColorPanelControl.Color);
        }

        /// <summary>
        /// Handles the Click event of noColorButton object.
        /// </summary>
        private void noColorButton_Click(object sender, RoutedEventArgs e)
        {
            backgroundColorPanelControl.Color = Colors.Transparent;
            _changedProperties[CellStyleProperty.FillColor] = VintasoftColor.Empty;
        }

        #endregion


        #region Common

        /// <summary>
        /// Sets the current style parameters.
        /// </summary>
        /// <param name="visualEditor">Spreadsheet visual editor.</param>
        private void SetVisualEditor(SpreadsheetVisualEditor visualEditor)
        {
            _visualEditor = visualEditor;
            CellStyle cellStyle = visualEditor.FocusedCellStyle;

            // get selected cells style indexes
            List<int> cellsStyleIndexes = new List<int>();
            CellReferences[] selectedCells = visualEditor.SelectedCells.ToArray();

            foreach (CellReferences cellReferences in selectedCells)
                cellsStyleIndexes.AddRange(visualEditor.FocusedWorksheet.GetCellsStyleIndexes(cellReferences));

            // get common properties
            Dictionary<CellStyleProperty, object> commonProperties = visualEditor.GetCommonCellsStyleProperties();

            // Number Format
            InitNumberFormatPropertiesUI(commonProperties);

            // Alignment
            InitAlignmentPropertiesUI(commonProperties);

            // Font
            InitFontPropertiesUI(commonProperties);

            // Borders
            InitBorderPropertiesUI(selectedCells);

            // Fill
            InitFillPropertiesUI(commonProperties);

            _initializationFinished = true;
        }

        /// <summary>
        /// Handles the Click event of okButton object.
        /// </summary>
        private void okButton_Click(object sender, RoutedEventArgs e)
        {
            if (_changedProperties.Count > 0)
                _visualEditor.ChangeStyleProperties(_changedProperties);
            DialogResult = true;
        }


        #endregion

        #endregion
    }
}

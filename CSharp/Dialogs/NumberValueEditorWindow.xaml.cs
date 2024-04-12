using System;
using System.Globalization;
using System.Windows;

using Vintasoft.Imaging.Office.Spreadsheet.UI;

using WpfDemosCommonCode;

namespace WpfSpreadsheetEditorDemo
{
    /// <summary>
    /// A dialog that allows to edit property value of <see cref="double"/> type.
    /// </summary>
    public partial class NumberValueEditorWindow : Window
    {

        #region Fields

        /// <summary>
        /// Spreadsheet visual editor.
        /// </summary>
        SpreadsheetVisualEditor _visualEditor;

        /// <summary>
        /// The minimum value.
        /// </summary>
        double _minValue;

        /// <summary>
        /// The maximum value.
        /// </summary>
        double _maxValue;

        /// <summary>
        /// The property name.
        /// </summary>
        string _propertyName;

        #endregion



        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="NumberValueEditorWindow"/> class.
        /// </summary>
        /// <param name="spreadsheetVisualEditor">The spreadsheet visual editor.</param>
        /// <param name="value">Current value.</param>
        /// <param name="minValue">Minimun value.</param>
        /// <param name="maxValue">Maximum value.</param>
        /// <param name="propertyName">Property name.</param>
        public NumberValueEditorWindow(SpreadsheetVisualEditor spreadsheetVisualEditor, double value, double minValue, double maxValue, string propertyName)
        {
            InitializeComponent();

            _visualEditor = spreadsheetVisualEditor;

            _value = value;
            _minValue = minValue;
            _maxValue = maxValue;
            _propertyName = propertyName;

            valueTextBox.Text = Math.Round(value, 2).ToString(Culture);
            valueTextBox.Focus();
            valueTextBox.CaretIndex = valueTextBox.Text.Length;
            propertyNameLabel.Text = propertyName + ":";
            Title = propertyName;
        }

        #endregion



        #region Properties

        double _value;
        /// <summary>
        /// Gets the property value.
        /// </summary>
        public double Value
        {
            get
            {
                return _value;
            }
        }

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

        #endregion



        #region Methods

        /// <summary>
        /// Handles the Click event of okButton object.
        /// </summary>
        private void okButton_Click(object sender, RoutedEventArgs e)
        {
            // parse the value
            double value;
            if (double.TryParse(valueTextBox.Text, NumberStyles.Float, Culture, out value))
            {
                // if value is in specified range
                if (value >= _minValue && value <= _maxValue)
                {
                    // update value
                    _value = value;

                    DialogResult = true;
                }
                else
                {
                    // create error message
                    string message = string.Format("{0} must be between {1} and {2}.", _propertyName, _minValue, _maxValue);
                    DemosTools.ShowWarningMessage("Spreadsheet Editor Demo", message);
                    return;
                }
            }
            else
            {
                // create error message
                string message = string.Format("{0} requires an integer or decimal number.", _propertyName);
                DemosTools.ShowWarningMessage("Spreadsheet Editor Demo", message);
                return;
            }
        }

        #endregion

    }
}

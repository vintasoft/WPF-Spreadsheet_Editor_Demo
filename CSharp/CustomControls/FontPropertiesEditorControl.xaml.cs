using System;
using System.ComponentModel;
using System.Globalization;
using System.Windows.Controls;

using Vintasoft.Imaging.Office.Spreadsheet.Document;
using Vintasoft.Imaging.Wpf;
using Vintasoft.Primitives;

namespace WpfSpreadsheetEditorDemo.CustomControls
{
    /// <summary>
    /// A control that allows to show and change the font properties.
    /// </summary>
    public partial class FontPropertiesEditorControl : UserControl
    {

        #region Fields

        /// <summary>
        /// A value indicating whether control initialization is finished.
        /// </summary>
        bool _initializationFinished = false;

        /// <summary>
        /// The initial font properties.
        /// </summary>
        FontProperties _initialFontProperties;

        #endregion



        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="FontPropertiesEditorControl"/> class.
        /// </summary>
        public FontPropertiesEditorControl()
        {
            InitializeComponent();

            fontStylesListBox.Items.Add("Regular");
            fontStylesListBox.Items.Add("Italic");
            fontStylesListBox.Items.Add("Bold");
            fontStylesListBox.Items.Add("Bold Italic");

            fontSizesListBox.Items.Add("8");
            fontSizesListBox.Items.Add("9");
            fontSizesListBox.Items.Add("10");
            fontSizesListBox.Items.Add("11");
            fontSizesListBox.Items.Add("12");
            fontSizesListBox.Items.Add("14");
            fontSizesListBox.Items.Add("16");
            fontSizesListBox.Items.Add("18");
            fontSizesListBox.Items.Add("20");
            fontSizesListBox.Items.Add("22");
            fontSizesListBox.Items.Add("24");
            fontSizesListBox.Items.Add("26");
            fontSizesListBox.Items.Add("28");
            fontSizesListBox.Items.Add("36");
            fontSizesListBox.Items.Add("48");
            fontSizesListBox.Items.Add("72");
        }

        #endregion



        #region Properties

        /// <summary>
        /// Gets or sets the font properties.
        /// </summary>
        [Browsable(false)]
        public FontProperties FontProperties
        {
            get
            {
                if (DesignerProperties.GetIsInDesignMode(this))
                    return null;
                return GetFontProperties();
            }
            set
            {
                _initialFontProperties = value;

                _initializationFinished = false;
                SetFontProperties(value);
                _initializationFinished = true;
            }
        }

        FontProperties _normalFontProperties;
        /// <summary>
        /// Gets or sets the "Normal" font properties.
        /// </summary>
        [Browsable(false)]
        public FontProperties NormalFontProperties
        {
            get
            {
                if (DesignerProperties.GetIsInDesignMode(this))
                    return null;
                return _normalFontProperties;
            }
            set
            {
                _normalFontProperties = value;
            }
        }

        CultureInfo _culture;
        /// <summary>
        /// The current culture.
        /// </summary>
        [Browsable(false)]
        public CultureInfo Culture
        {
            get
            {
                return _culture;
            }
            set
            {
                _culture = value;
            }
        }

        #endregion



        #region Methods

        /// <summary>
        /// Returns the font properties.
        /// </summary>
        /// <returns>The font properties.</returns>
        private FontProperties GetFontProperties()
        {
            // get font name
            string fontName;
            if (!string.IsNullOrEmpty(fontNameTextBox.Text))
                fontName = fontNameTextBox.Text;
            else
                fontName = _initialFontProperties.Name;

            // get font style
            bool isBold = fontStyleTextBox.Text == "Bold" || fontStyleTextBox.Text == "Bold Italic";
            bool isItalic = fontStyleTextBox.Text == "Italic" || fontStyleTextBox.Text == "Bold Italic";

            if (fontStyleTextBox.Text == "Regular")
            {
                isBold = false;
                isItalic = false;
            }
            else if (fontStyleTextBox.Text == "Bold")
            {
                isBold = true;
            }
            else if (fontStyleTextBox.Text == "Italic")
            {
                isItalic = true;
            }
            else if (fontStyleTextBox.Text == "Bold Italic")
            {
                isBold = true;
                isItalic = true;
            }
            else if (string.IsNullOrEmpty(fontStyleTextBox.Text))
            {
                isBold = _initialFontProperties.IsBold;
                isItalic = _initialFontProperties.IsItalic;
            }
            else
            {
                throw new Exception(string.Format("Font style with name '{0}' does not exist.", fontStyleTextBox.Text));
            }

            // get font size
            double fontSize;
            if (!string.IsNullOrEmpty(fontSizeTextBox.Text))
            {
                double parsedFontSize;
                if (double.TryParse(fontSizeTextBox.Text, NumberStyles.Float, _culture, out parsedFontSize))
                    fontSize = Math.Round(parsedFontSize, 2);
                else
                    throw new Exception("Font size must be an integer or decimal number.");
            }
            else
            {
                fontSize = _initialFontProperties.Size;
            }

            // get font effects
            bool isUnderline = underlineCheckBox.IsChecked != null && underlineCheckBox.IsChecked.Value == true;
            bool isStrikethrough = strikethroughCheckBox.IsChecked!=null && strikethroughCheckBox.IsChecked.Value == true;

            // get font color
            VintasoftColor fontColor = WpfObjectConverter.Convert(fontColorPanelControl.Color);

            return new FontProperties(fontName, fontSize, fontColor, isBold, isItalic, isUnderline, isStrikethrough);
        }

        /// <summary>
        /// Initializes the UI with font properties.
        /// </summary>
        /// <param name="fontProperties">The font properties.</param>
        private void SetFontProperties(FontProperties fontProperties)
        {
            if (fontProperties == null)
            {
                this.IsEnabled = false;
                return;
            }
            else
            {
                this.IsEnabled = true;
            }

            if (!_initializationFinished)
            {
                // initialize names of available fonts
                foreach (string fontName in CellsStyleWindow.GetAvailableFontNames())
                    fontNamesListBox.Items.Add(fontName);
            }

            // init font name
            fontNamesListBox.SelectedItem = fontProperties.Name;

            // init font style
            if (fontProperties.IsBold && fontProperties.IsItalic)
                fontStylesListBox.SelectedItem = "Bold Italic";
            else if (fontProperties.IsBold)
                fontStylesListBox.SelectedItem = "Bold";
            else if (fontProperties.IsItalic)
                fontStylesListBox.SelectedItem = "Italic";
            else
                fontStylesListBox.SelectedItem = "Regular";

            // init font size
            double fontSize = fontProperties.Size;
            fontSizeTextBox.Text = fontSize.ToString(_culture);
            if (fontSizesListBox.Items.Contains(fontSizeTextBox.Text))
                fontSizesListBox.SelectedItem = fontSizeTextBox.Text;
            else
                fontSizesListBox.SelectedItem = null;

            // init text underline
            underlineCheckBox.IsChecked = fontProperties.IsUnderline;

            // init text strikeout
            strikethroughCheckBox.IsChecked = fontProperties.IsStrikeout;

            // init font color
            VintasoftColor fontColor = fontProperties.Color;
            fontColorPanelControl.Color = WpfObjectConverter.Convert(fontColor);
        }

        /// <summary>
        /// Handles the SelectionChanged event of FontNamesListBox object.
        /// </summary>
        private void fontNamesListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (fontNamesListBox.SelectedItem != null)
                fontNameTextBox.Text = (string)fontNamesListBox.SelectedItem;
        }

        /// <summary>
        /// Handles the SelectionChanged event of FontStylesListBox object.
        /// </summary>
        private void fontStylesListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (fontStylesListBox.SelectedItem != null)
                fontStyleTextBox.Text = (string)fontStylesListBox.SelectedItem;
        }

        /// <summary>
        /// Handles the SelectionChanged event of FontSizesListBox object.
        /// </summary>
        private void fontSizesListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (fontSizesListBox.SelectedItem != null)
                fontSizeTextBox.Text = (string)fontSizesListBox.SelectedItem;
        }

        /// <summary>
        /// Handles the Click event of NormalFontButton object.
        /// </summary>
        private void normalFontButton_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            // set the UI to the "Normal" style
            SetFontProperties(NormalFontProperties);
        }

        #endregion
    }
}

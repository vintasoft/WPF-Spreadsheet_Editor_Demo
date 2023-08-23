using System;
using System.Collections.Generic;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

using Vintasoft.Imaging;
using Vintasoft.Imaging.Office.OpenXml.Wpf.UI;
using Vintasoft.Imaging.Office.Spreadsheet.Document;
using Vintasoft.Imaging.Office.Spreadsheet.Wpf.UI;
using Vintasoft.Imaging.Wpf;
using Vintasoft.Primitives;

using WpfDemosCommonCode.CustomControls;

namespace WpfSpreadsheetEditorDemo
{
    /// <summary>
    /// A panel control that allows to edit font properties.
    /// </summary>
    public partial class FontPropertiesPanel : SpreadsheetVisualEditorPanel
    {

        #region Fields

        /// <summary>
        /// Indicates whether the UI is updating now.
        /// </summary>
        bool _updateUI = false;

        /// <summary>
        /// Indicates whether style painting was executed.
        /// </summary>
        bool _setStyleCellIsExecuted = false;

        /// <summary>
        /// Current border color, that is applied when setting borders.
        /// </summary>
        Color _borderColor = Colors.Black;

        /// <summary>
        /// Current font color, that is applied by font color button.
        /// </summary>
        Color _fontColor = Colors.Red;

        /// <summary>
        /// Current fill color, that is applied by fill color button.
        /// </summary>
        Color _fillColor = Colors.Yellow;

        /// <summary>
        /// Dictionary: border button => cells borders.
        /// </summary>
        Dictionary<MenuItem, CellsBorders> _buttonsToBorders = new Dictionary<MenuItem, CellsBorders>();

        #endregion



        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="FontPropertiesPanel"/> class.
        /// </summary>
        public FontPropertiesPanel()
        {
            InitializeComponent();

            foreach (string fontName in CellsStyleWindow.GetAvailableFontNames())
                fontNameComboBox.Items.Add(fontName);

            fontColorButtonRectangle.Fill = new SolidColorBrush(_fontColor);
            fillColorButtonRectangle.Fill = new SolidColorBrush(_fillColor);

            InitBordersDropDownMenu();
        }

        #endregion



        #region Methods

        #region PROTECTED

        /// <summary>
        /// Raises the OnSpreadsheetEditorChanged event.
        /// </summary>
        /// <param name="args">The event data.</param>
        protected override void OnSpreadsheetEditorChanged(PropertyChangedEventArgs<WpfSpreadsheetEditorControl> args)
        {
            base.OnSpreadsheetEditorChanged(args);

            if (args.OldValue != null)
            {
                args.OldValue.VisualEditor.FocusedCellChanged -= VisualEditor_FocusedCellChanged;
                args.OldValue.VisualEditor.FocusedCommentChanged -= VisualEditor_FocusedCommentChanged;
                args.OldValue.VisualEditor.FocusedCellsChanged -= VisualEditor_FocusedCellsChanged;
                args.OldValue.VisualEditor.CellsStylePropertiesChanged -= VisualEditor_CellsStylePropertiesChanged;
                args.OldValue.VisualEditor.FocusedWorksheetChanged -= VisualEditor_FocusedWorksheetChanged;
            }
            if (args.NewValue != null)
            {
                args.NewValue.VisualEditor.FocusedCellChanged += VisualEditor_FocusedCellChanged;
                args.NewValue.VisualEditor.FocusedCommentChanged += VisualEditor_FocusedCommentChanged;
                args.NewValue.VisualEditor.FocusedCellsChanged += VisualEditor_FocusedCellsChanged;
                args.NewValue.VisualEditor.CellsStylePropertiesChanged += VisualEditor_CellsStylePropertiesChanged;
                args.NewValue.VisualEditor.FocusedWorksheetChanged += VisualEditor_FocusedWorksheetChanged;
            }
            UpdateUI();
        }

        #endregion


        #region PRIVATE

        #region UI

        private void FontPropertiesPanel_VisibleChanged(object sender, RoutedEventArgs e)
        {
            if (VisualEditor != null)
                UpdateUI();
        }

        #region Panel buttons

        #region Font

        private void fontNameComboBox_Leave(object sender, RoutedEventArgs e)
        {
            SetFontName(fontNameComboBox.Text);
        }

        /// <summary>
        /// Sets the font name.
        /// </summary>
        /// <param name="fontName">The font name.</param>
        private void SetFontName(string fontName)
        {
            if (!_updateUI)
            {
                if (VisualEditor.FontName != fontName)
                {
                    VisualEditor.FontName = fontName;
                    UpdateUI();
                }
            }
        }

        /// <summary>
        /// Handles the KeyDown event of FontNameComboBox object.
        /// </summary>
        private void fontNameComboBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                SetFontName(fontNameComboBox.Text);
            }
        }

        /// <summary>
        /// Handles the SelectionChanged event of FontNameComboBox object.
        /// </summary>
        private void fontNameComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_updateUI)
                return;

            if (e.AddedItems.Count > 0)
            {
                string text;
                if (e.AddedItems[0] is ComboBoxItem)
                    text = ((ComboBoxItem)e.AddedItems[0]).Content.ToString();
                else
                    text = e.AddedItems[0].ToString();
                SetFontName(text);
                SpreadsheetEditor.Focus();
            }
        }

        private void fontSizeComboBox_Leave(object sender, RoutedEventArgs e)
        {
            SetFontSize(fontSizeComboBox.Text);
        }

        /// <summary>
        /// Handles the SelectionChanged event of FontSizeComboBox object.
        /// </summary>
        private void fontSizeComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_updateUI)
                return;

            if (e.AddedItems.Count > 0)
            {
                string text;
                if (e.AddedItems[0] is ComboBoxItem)
                    text = ((ComboBoxItem)e.AddedItems[0]).Content.ToString();
                else
                    text = e.AddedItems[0].ToString();
                SetFontSize(text);
                SpreadsheetEditor.Focus();
            }
        }

        /// <summary>
        /// Handles the KeyDown event of FontSizeComboBox object.
        /// </summary>
        private void fontSizeComboBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
                SetFontSize(fontSizeComboBox.Text);
        }

        /// <summary>
        /// Handles the Click event of IncFontSizeButton object.
        /// </summary>
        private void incFontSizeButton_Click(object sender, RoutedEventArgs e)
        {
            int index = GetPredefinedSizeIndex();
            int predefinedFontSize = int.Parse(((ComboBoxItem)fontSizeComboBox.Items[index]).Content.ToString());
            if (predefinedFontSize > VisualEditor.FontSize)
                fontSizeComboBox.SelectedIndex = index;
            else if (index < fontSizeComboBox.Items.Count - 1)
                fontSizeComboBox.SelectedIndex = index + 1;
        }

        /// <summary>
        /// Handles the Click event of DecFontSizeButton object.
        /// </summary>
        private void decFontSizeButton_Click(object sender, RoutedEventArgs e)
        {
            int index = GetPredefinedSizeIndex();
            int predefinedFontSize = int.Parse(((ComboBoxItem)fontSizeComboBox.Items[index]).Content.ToString());
            if (predefinedFontSize < VisualEditor.FontSize)
                fontSizeComboBox.SelectedIndex = index;
            else if (index > 0)
                fontSizeComboBox.SelectedIndex = index - 1;
        }

        /// <summary>
        /// Handles the Click event of BoldFontButton object.
        /// </summary>
        private void boldFontButton_Click(object sender, RoutedEventArgs e)
        {
            VisualEditor.IsFontBold = !boldFontButton.IsChecked;
            UpdateUI();
            SpreadsheetEditor.Focus();
        }

        /// <summary>
        /// Handles the Click event of ItalicFontButton object.
        /// </summary>
        private void italicFontButton_Click(object sender, RoutedEventArgs e)
        {
            VisualEditor.IsFontItalic = !italicFontButton.IsChecked;
            UpdateUI();
            SpreadsheetEditor.Focus();
        }

        /// <summary>
        /// Handles the Click event of UnderlineButton object.
        /// </summary>
        private void underlineButton_Click(object sender, RoutedEventArgs e)
        {
            VisualEditor.IsFontUnderline = !underlineButton.IsChecked;
            UpdateUI();
            SpreadsheetEditor.Focus();
        }

        /// <summary>
        /// Handles the Click event of StrikeoutButton object.
        /// </summary>
        private void strikeoutButton_Click(object sender, RoutedEventArgs e)
        {
            VisualEditor.IsFontStrikeout = !strikeoutButton.IsChecked;
            UpdateUI();
            SpreadsheetEditor.Focus();
        }

        /// <summary>
        /// Handles the Click event of NoFillMenuItem object.
        /// </summary>
        private void noFillMenuItem_Click(object sender, RoutedEventArgs e)
        {
            _fillColor = new Color();
            VisualEditor.FillColor = WpfObjectConverter.Convert(_fillColor);
            fillColorButtonRectangle.Fill = new SolidColorBrush(_fillColor);
            SpreadsheetEditor.Focus();
        }

        /// <summary>
        /// Handles the Click event of SelectFillColorMenuItem object.
        /// </summary>
        private void selectFillColorMenuItem_Click(object sender, RoutedEventArgs e)
        {
            ColorPickerDialog colorDialog1 = new ColorPickerDialog();
            colorDialog1.WindowStartupLocation = WindowStartupLocation.CenterOwner;
            colorDialog1.Owner = Application.Current.MainWindow;
            colorDialog1.StartingColor = _fillColor;
            colorDialog1.CanEditAlphaChannel = false;
            if (colorDialog1.ShowDialog() == true)
            {
                _fillColor = colorDialog1.SelectedColor;
                VisualEditor.FillColor = WpfObjectConverter.Convert(_fillColor);
                fillColorButtonRectangle.Fill = new SolidColorBrush(_fillColor);
            }
            SpreadsheetEditor.Focus();
        }

        private void fontColorButton_ButtonClick(object sender, RoutedEventArgs e)
        {
            VisualEditor.FontColor = WpfObjectConverter.Convert(_fontColor);
            SpreadsheetEditor.Focus();
        }

        /// <summary>
        /// Handles the Click event of SelectFontColorMenuItem object.
        /// </summary>
        private void selectFontColorMenuItem_Click(object sender, RoutedEventArgs e)
        {
            ColorPickerDialog colorDialog1 = new ColorPickerDialog();
            colorDialog1.WindowStartupLocation = WindowStartupLocation.CenterOwner;
            colorDialog1.Owner = Application.Current.MainWindow;
            colorDialog1.StartingColor = _fontColor;
            colorDialog1.CanEditAlphaChannel = false;
            if (colorDialog1.ShowDialog() == true)
            {
                _fontColor = colorDialog1.SelectedColor;
                VisualEditor.FontColor = WpfObjectConverter.Convert(_fontColor);
                fontColorButtonRectangle.Fill = new SolidColorBrush(_fontColor);
            }
            SpreadsheetEditor.Focus();
        }

        /// <summary>
        /// Handles the Click event of FontPropertiesButton object.
        /// </summary>
        private void fontPropertiesButton_Click(object sender, RoutedEventArgs e)
        {
            CellsStyleWindow.ShowFontDialog(VisualEditor);
            UpdateUI();
        }

        #endregion


        #region Fill

        private void fillColorButton_ButtonClick(object sender, RoutedEventArgs e)
        {
            VisualEditor.FillColor = WpfObjectConverter.Convert(_fillColor);
            SpreadsheetEditor.Focus();
        }

        #endregion


        #region Borders

        /// <summary>
        /// Initializes border styles for panel button and it's drop down menu items.
        /// </summary>
        private void InitBordersDropDownMenu()
        {
            CellBorder thinBorder = new CellBorder(CellBorderStyle.Thin, WpfObjectConverter.Convert(_borderColor));
            CellBorder mediumBorder = new CellBorder(CellBorderStyle.Medium, WpfObjectConverter.Convert(_borderColor));
            CellBorder doubleBorder = new CellBorder(CellBorderStyle.Double, WpfObjectConverter.Convert(_borderColor));

            // default border for button is "all borders"
            CellsBorders borders = new CellsBorders(new CellBorders(thinBorder), thinBorder, thinBorder);
            _buttonsToBorders.Add(bordersButton, borders);

            // bottom border
            borders = new CellsBorders(new CellBorders(null, null, null, thinBorder), null, null);
            _buttonsToBorders.Add(bottomBorderMenuItem, borders);

            // top border
            borders = new CellsBorders(new CellBorders(null, null, thinBorder, null), null, null);
            _buttonsToBorders.Add(topBorderMenuItem, borders);

            // left border
            borders = new CellsBorders(new CellBorders(thinBorder, null, null, null), null, null);
            _buttonsToBorders.Add(leftBorderMenuItem, borders);

            // right border
            borders = new CellsBorders(new CellBorders(null, thinBorder, null, null), null, null);
            _buttonsToBorders.Add(rightBorderMenuItem, borders);

            // no border
            _buttonsToBorders.Add(noBorderMenuItem, null);

            // all borders
            borders = new CellsBorders(new CellBorders(thinBorder), thinBorder, thinBorder);
            _buttonsToBorders.Add(allBordersMenuItem, borders);

            // outside borders
            borders = new CellsBorders(new CellBorders(thinBorder), null, null);
            _buttonsToBorders.Add(outsideBordersMenuItem, borders);

            // thick outside borders
            borders = new CellsBorders(new CellBorders(mediumBorder), null, null);
            _buttonsToBorders.Add(thickOutsideBordersMenuItem, borders);

            // bottom double border
            borders = new CellsBorders(new CellBorders(null, null, null, doubleBorder), null, null);
            _buttonsToBorders.Add(bottomDoubleBorderMenuItem, borders);

            // thick bottom border
            borders = new CellsBorders(new CellBorders(null, null, null, mediumBorder), null, null);
            _buttonsToBorders.Add(thickBottomBorderMenuItem, borders);

            // top and bottom borders
            borders = new CellsBorders(new CellBorders(null, null, thinBorder, thinBorder), null, null);
            _buttonsToBorders.Add(topAndBottomBorderMenuItem, borders);

            // top and thick bottom border
            borders = new CellsBorders(new CellBorders(null, null, thinBorder, mediumBorder), null, null);
            _buttonsToBorders.Add(topAndThickBottomBorderMenuItem, borders);

            // top and double bottom border
            borders = new CellsBorders(new CellBorders(null, null, thinBorder, doubleBorder), null, null);
            _buttonsToBorders.Add(topAndDoubleBottomBorderMenuItem, borders);
        }

        /// <summary>
        /// Handles the Click event of BordersDropDownButton object.
        /// </summary>
        private void bordersDropDownButton_Click(object sender, RoutedEventArgs e)
        {
            Button bordersButton = sender as Button;
            // if the borders button is clicked
            if (bordersButton != null)
            {
                // if border is set
                if (_buttonsToBorders[this.bordersButton] != null)
                {
                    // set the selected borders style with current border color
                    VisualEditor.CellsBorders = ChangeBordersColor(_buttonsToBorders[this.bordersButton], _borderColor);
                }
                else
                {
                    // set invisible border
                    VisualEditor.CellsBorders = new CellsBorders(new CellBorders(CellBorder.Invisible), CellBorder.Invisible, CellBorder.Invisible);
                }
                return;
            }

            MenuItem borderMenuButton = sender as MenuItem;
            // if the menu item is clicked
            if (borderMenuButton != null)
            {
                // if border is set
                if (_buttonsToBorders[borderMenuButton] != null)
                {
                    // get the selected borders with current border color
                    CellsBorders borders = ChangeBordersColor(_buttonsToBorders[borderMenuButton], _borderColor);
                    // set the selected borders to the cells
                    VisualEditor.CellsBorders = borders;
                    // set the selected borders for borders button
                    _buttonsToBorders[this.bordersButton] = borders;
                }
                else
                {
                    // set invisible borders to the cells
                    VisualEditor.CellsBorders = new CellsBorders(new CellBorders(CellBorder.Invisible), CellBorder.Invisible, CellBorder.Invisible);
                    // set null for borders button
                    _buttonsToBorders[this.bordersButton] = null;
                }
                
                // update image of borders button
                bordersButtonImage.Source = ((Image)borderMenuButton.Icon).Source;
            }
        }

        /// <summary>
        /// Handles the Click event of BorderColorMenuItem object.
        /// </summary>
        private void borderColorMenuItem_Click(object sender, RoutedEventArgs e)
        {
            ColorPickerDialog colorDialog1 = new ColorPickerDialog();
            colorDialog1.WindowStartupLocation = WindowStartupLocation.CenterOwner;
            colorDialog1.Owner = Application.Current.MainWindow;
            colorDialog1.StartingColor = _borderColor;
            colorDialog1.CanEditAlphaChannel = false;
            if (colorDialog1.ShowDialog() == true)
            {
                _borderColor = colorDialog1.SelectedColor;
                VisualEditor.BordersColor = WpfObjectConverter.Convert(_borderColor);
            }
        }

        /// <summary>
        /// Handles the Click event of MoreBordersMenuItem object.
        /// </summary>
        private void moreBordersMenuItem_Click(object sender, RoutedEventArgs e)
        {
            CellsStyleWindow.ShowBordersDialog(VisualEditor);
        }

        /// <summary>
        /// Returns a copy of <see cref="CellsBorders"/> with a specified color.
        /// </summary>
        /// <param name="borders">Cells borders.</param>
        /// <param name="color">Color.</param>
        private CellsBorders ChangeBordersColor(CellsBorders borders, Color borderColor)
        {
            VintasoftColor color = WpfObjectConverter.Convert(borderColor);

            CellBorder leftBorder = null;
            CellBorder rightBorder = null;
            CellBorder topBorder = null;
            CellBorder bottomBorder = null;

            if (borders.OutsideBorders != null)
            {
                if (borders.OutsideBorders.Left != null)
                    leftBorder = new CellBorder(borders.OutsideBorders.Left.Style, color);
                if (borders.OutsideBorders.Right != null)
                    rightBorder = new CellBorder(borders.OutsideBorders.Right.Style, color);
                if (borders.OutsideBorders.Top != null)
                    topBorder = new CellBorder(borders.OutsideBorders.Top.Style, color);
                if (borders.OutsideBorders.Bottom != null)
                    bottomBorder = new CellBorder(borders.OutsideBorders.Bottom.Style, color);
            }

            CellBorder horizontalBorder = null;
            if (borders.HorizontalBorder != null)
                horizontalBorder = new CellBorder(borders.HorizontalBorder.Style, color);

            CellBorder verticalBorder = null;
            if (borders.VerticalBorder != null)
                verticalBorder = new CellBorder(borders.VerticalBorder.Style, color);

            CellBorders outsideBorders = new CellBorders(leftBorder, rightBorder, topBorder, bottomBorder);

            return new CellsBorders(outsideBorders, horizontalBorder, verticalBorder);
        }

        #endregion


        #region Style paint

        /// <summary>
        /// Handles the Click event of CopyStyleButton object.
        /// </summary>
        private void copyStyleButton_Click(object sender, RoutedEventArgs e)
        {
            if (VisualEditor.StylePainterSource != null)
                FinishStylePaint();
            else
                StartStylePaint();
        }

        #endregion

        #endregion


        #region Visual Editor

        private void VisualEditor_CellsStylePropertiesChanged(object sender, EventArgs e)
        {
            UpdateUI();
        }

        private void VisualEditor_FocusedWorksheetChanged(object sender, PropertyChangedEventArgs<Worksheet> e)
        {
            UpdateUI();
        }

        private void VisualEditor_FocusedCellsChanged(object sender, PropertyChangedEventArgs<CellReferences> e)
        {
            // if source style cell is selected and new selection is complete
            if (VisualEditor.StylePainterSource != null && Equals(e.NewValue, e.OldValue))
            {
                bool isControlKeyPressed = Keyboard.IsKeyDown(Key.LeftCtrl) || Keyboard.IsKeyDown(Key.RightCtrl);

                // if control key is pressed and style painting is executing
                if (isControlKeyPressed || !_setStyleCellIsExecuted)
                    VisualEditor.PerformStylePaint();

                // if control key is not pressed
                if (!isControlKeyPressed)
                    FinishStylePaint();

                _setStyleCellIsExecuted = true;
            }            
        }

        private void VisualEditor_FocusedCellChanged(object sender, PropertyChangedEventArgs<CellReference> e)
        {
            UpdateUI();
        }

        private void VisualEditor_FocusedCommentChanged(object sender, PropertyChangedEventArgs<CellComment> e)
        {
            UpdateUI();
        }

        #endregion

        #endregion


        #region UI state

        /// <summary>
        /// Updates the user interface of this form.
        /// </summary>
        public void UpdateUI()
        {
            if (VisualEditor.IsFocusedWorksheetChanging)
                return;

            _updateUI = true;
            try
            {
                if (VisualEditor.FocusedCell == null && VisualEditor.FocusedComment == null)
                {
                    IsEnabled = false;
                }
                else
                {
                    IsEnabled = true;
                    // update panel state from focused cell properties
                    FontProperties fontProperties = VisualEditor.FocusedFontProperties;
                    fontNameComboBox.Text = fontProperties.Name;
                    fontSizeComboBox.Text = fontProperties.Size.ToString(UICulture);
                    boldFontButton.IsChecked = fontProperties.IsBold;
                    italicFontButton.IsChecked = fontProperties.IsItalic;
                    underlineButton.IsChecked = fontProperties.IsUnderline;
                    strikeoutButton.IsChecked = fontProperties.IsStrikeout;

                    bool commentIsFocused = VisualEditor.FocusedComment != null;
                    bordersButton.IsEnabled = !commentIsFocused;
                    fontPropertiesButton.IsEnabled = !commentIsFocused;
                    copyStyleButton.IsEnabled = !commentIsFocused;
                }
            }
            finally
            {
                _updateUI = false;
            }
        }

        #endregion


        #region Font size

        /// <summary>
        /// Sets the font size.
        /// </summary>
        /// <param name="fontSize">The font size.</param>
        private void SetFontSize(string fontSize)
        {
            if (!_updateUI)
            {
                double size;
                if (double.TryParse(fontSize, NumberStyles.Number, UICulture, out size))
                {
                    if (VisualEditor.FontSize != size)
                    {
                        VisualEditor.FontSize = size;
                        UpdateUI();
                    }
                }
            }
        }

        /// <summary>
        /// Returns the index of the font size value in the font size combo box.
        /// </summary>
        private int GetPredefinedSizeIndex()
        {
            double size;
            if (!double.TryParse(fontSizeComboBox.Text, NumberStyles.Number, UICulture, out size))
                size = VisualEditor.Document.Styles[0].FontProperties.Size;
            for (int i = 0; i < fontSizeComboBox.Items.Count; i++)
                if (int.Parse(((ComboBoxItem)fontSizeComboBox.Items[i]).Content.ToString()) >= size)
                    return i;
            return fontSizeComboBox.Items.Count - 1;
        }

        #endregion


        #region Style painting

        /// <summary>
        /// Activates style painting mode.
        /// </summary>
        private void StartStylePaint()
        {
            copyStyleButton.IsChecked = true;
            _setStyleCellIsExecuted = false;
            VisualEditor.StylePainterSource = VisualEditor.FocusedCells;
        }

        /// <summary>
        /// Deactivates style painting mode.
        /// </summary>
        private void FinishStylePaint()
        {
            VisualEditor.StylePainterSource = null;
            copyStyleButton.IsChecked = false;
        }

        #endregion

        #endregion

        #endregion

    }
}

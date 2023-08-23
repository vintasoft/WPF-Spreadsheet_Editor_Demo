using System;
using System.Windows.Controls;
using System.Windows.Media;

using Vintasoft.Imaging.Office.Spreadsheet.UI;
using Vintasoft.Imaging.Wpf;
using Vintasoft.Primitives;

namespace WpfSpreadsheetEditorDemo.CustomControls
{
    /// <summary>
    /// A control that allows to show and change the appearance settings of cells.
    /// </summary>
    public partial class CellReferencesAppearanceEditorControl : UserControl
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="CellReferencesAppearanceEditorControl"/> class.
        /// </summary>
        public CellReferencesAppearanceEditorControl()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Gets or sets the appearance of current cell.
        /// </summary>
        public CellReferencesAppearance CellsAppearance
        {
            get
            {
                VintasoftColor fillColor = WpfObjectConverter.Convert(fillColorPanelControl.Color);
                VintasoftColor borderColor = WpfObjectConverter.Convert(borderColorPanelControl.Color);
                int borderWidth = (int)borderWidthNumericUpDown.Value;

                return new CellReferencesAppearance(fillColor, borderColor, borderWidth);
            }
            set
            {
                if (value != null)
                {
                    fillColorPanelControl.Color = WpfObjectConverter.Convert(value.FillColor);
                    borderColorPanelControl.Color = WpfObjectConverter.Convert(value.BorderColor);
                    borderWidthNumericUpDown.Value = (int)Math.Round(value.BorderWidth, 0);
                }
                else
                {
                    fillColorPanelControl.Color = Colors.Transparent;
                    borderColorPanelControl.Color = Colors.Transparent;
                    borderWidthNumericUpDown.Value = 0;
                }
            }
        }

    }
}

using System;
using System.Windows.Controls;
using System.Windows.Media;

using Vintasoft.Imaging.Office.Spreadsheet.Document;
using Vintasoft.Imaging.Wpf;
using Vintasoft.Primitives;

namespace WpfSpreadsheetEditorDemo.CustomControls
{
    /// <summary>
    /// A control that allows to show and change the <see cref="Vintasoft.Imaging.Office.Spreadsheet.Document.ShapeAppearance"/>.
    /// </summary>
    public partial class ShapeAppearanceEditorControl : UserControl
    {
        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="ShapeAppearanceEditorControl"/> class.
        /// </summary>
        public ShapeAppearanceEditorControl()
        {
            InitializeComponent();
        }

        #endregion



        #region Properties

        /// <summary>
        /// Gets or sets the shape appearance.
        /// </summary>
        public ShapeAppearance ShapeAppearance
        {
            get
            {
                VintasoftColor fillColor = WpfObjectConverter.Convert(fillColorPanelControl.Color);
                VintasoftColor borderColor = WpfObjectConverter.Convert(outlineColorPanelControl.Color);
                int outlineWidth = (int)outlineWidthNumericUpDown.Value;

                return new ShapeAppearance(fillColor, borderColor, outlineWidth);
            }
            set
            {
                if (value != null)
                {
                    fillColorPanelControl.Color = WpfObjectConverter.Convert(value.FillColor);
                    outlineColorPanelControl.Color = WpfObjectConverter.Convert(value.OutlineColor);
                    outlineWidthNumericUpDown.Value = (int)Math.Round(value.OutlineWidth, 0);
                }
                else
                {
                    fillColorPanelControl.Color = Colors.Transparent;
                    outlineColorPanelControl.Color = Colors.Transparent;
                    outlineWidthNumericUpDown.Value = 0;
                }

                OnShapeAppearanceChanged();
            }
        }

        #endregion



        #region Methods

        /// <summary>
        /// Raises the <see cref="ShapeAppearanceChanged" /> event.
        /// </summary>
        public void OnShapeAppearanceChanged()
        {
            if (ShapeAppearanceChanged != null)
                ShapeAppearanceChanged(this, null);
        }

        /// <summary>
        /// Handles the ColorChanged event of FillColorPanelControl object.
        /// </summary>
        private void fillColorPanelControl_ColorChanged(object sender, EventArgs e)
        {
            OnShapeAppearanceChanged();
        }

        /// <summary>
        /// Handles the ColorChanged event of OutlineColorPanelControl object.
        /// </summary>
        private void outlineColorPanelControl_ColorChanged(object sender, EventArgs e)
        {
            OnShapeAppearanceChanged();
        }

        /// <summary>
        /// Handles the ValueChanged event of OutlineWidthNumericUpDown object.
        /// </summary>
        private void outlineWidthNumericUpDown_ValueChanged(object sender, EventArgs e)
        {
            OnShapeAppearanceChanged();
        }

        #endregion



        #region Events

        /// <summary>
        /// Occurs when <see cref="ShapeAppearance"/> property is changed.
        /// </summary>
        public event EventHandler ShapeAppearanceChanged;


        #endregion
    }
}

using System.Windows.Controls;

using Vintasoft.Imaging.Office.Spreadsheet.Document;
using Vintasoft.Primitives;

namespace WpfSpreadsheetEditorDemo.CustomControls
{
    /// <summary>
    /// A control that allows to show and change the location of sheet drawing.
    /// </summary>
    public partial class SheetDrawingLocationEditorControl : UserControl
    {

        #region Fields

        /// <summary>
        /// The source location.
        /// </summary>
        SheetDrawingLocation _sourceLocation;

        #endregion



        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="SheetDrawingLocationEditorControl"/> class.
        /// </summary>
        public SheetDrawingLocationEditorControl()
        {
            InitializeComponent();
            UpdateUI();
        }

        #endregion



        #region Properties

        /// <summary>
        /// Gets or sets the location of sheet drawing.
        /// </summary>
        public SheetDrawingLocation SheetDrawingLocation
        {
            get
            {
                if (Worksheet == null || _sourceLocation == null)
                    return _sourceLocation;

                // get location type
                SheetDrawingLocationType locationType;
                if (dontMoveOrSizeWithCellsRadioButton.IsChecked == true)
                    locationType = SheetDrawingLocationType.Absolute;
                else if (moveButDontSizeWithCellsRadioButton.IsChecked == true)
                    locationType = SheetDrawingLocationType.RelativeToCell;
                else
                    locationType = SheetDrawingLocationType.RelativeToTwoCell;

                // get location bounding box
                VintasoftRect boundingBox = new VintasoftRect(
                    (double)xNumericUpDown.Value, (double)yNumericUpDown.Value,
                    (double)widthNumericUpDown.Value, (double)heightNumericUpDown.Value);

                // create new drawing location
                return SheetDrawingLocation.Create(locationType, Worksheet, boundingBox);
            }
            set
            {
                _sourceLocation = value;

                UpdateUI();
            }
        }

        Worksheet _worksheet;
        /// <summary>
        /// Gets or sets the worksheet.
        /// </summary>
        public Worksheet Worksheet
        {
            get
            {
                return _worksheet;
            }
            set
            {
                _worksheet = value;

                UpdateUI();
            }
        }

        #endregion



        #region Methods

        /// <summary>
        /// Updates the user interface of this control.
        /// </summary>
        private void UpdateUI()
        {
            // the location bounding box
            VintasoftRect bounds;
            // the location type
            SheetDrawingLocationType locationType;

            // if location or type of bounding box cannot be calculated
            if (_sourceLocation == null || Worksheet == null)
            {
                bounds = VintasoftRect.Empty;
                locationType = SheetDrawingLocationType.Absolute;
            }
            else
            {
                bounds = _sourceLocation.GetBoundingBox(Worksheet);
                locationType = _sourceLocation.LocationType;
            }

            // show the location and size of bounding box

            xNumericUpDown.Value = bounds.X;
            yNumericUpDown.Value = bounds.Y;
            widthNumericUpDown.Value = bounds.Width;
            heightNumericUpDown.Value = bounds.Height;


            // show location type

            switch (locationType)
            {
                case SheetDrawingLocationType.Absolute:
                    dontMoveOrSizeWithCellsRadioButton.IsChecked = true;
                    break;

                case SheetDrawingLocationType.RelativeToCell:
                    moveButDontSizeWithCellsRadioButton.IsChecked = true;
                    break;

                case SheetDrawingLocationType.RelativeToTwoCell:
                    moveAndSizeWithCellsRadioButton.IsChecked = true;
                    break;

                default:
                    dontMoveOrSizeWithCellsRadioButton.IsChecked = true;
                    break;
            }
        }

        #endregion

    }
}

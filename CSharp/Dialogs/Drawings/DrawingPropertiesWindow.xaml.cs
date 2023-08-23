using System;
using System.Windows;

using Vintasoft.Imaging.Office.Spreadsheet.Document;
using Vintasoft.Imaging.Office.Spreadsheet.Document.Editors;
using Vintasoft.Imaging.Office.Spreadsheet.UI;

using WpfDemosCommonCode;

namespace WpfSpreadsheetEditorDemo
{
    /// <summary>
    /// A dialog that allows to edit the drawing properties.
    /// </summary>
    public partial class DrawingPropertiesWindow : Window
    {

        #region Fields

        /// <summary>
        /// The visual editor.
        /// </summary>
        SpreadsheetVisualEditor _visualEditor;

        /// <summary>
        /// The drawing.
        /// </summary>
        SheetDrawing _drawing;

        #endregion



        #region Constructors
        
        /// <summary>
        /// Initializes a new instance of the <see cref="DrawingPropertiesWindow"/> class.
        /// </summary>
        public DrawingPropertiesWindow()
        {
            InitializeComponent();
        } 

        /// <summary>
        /// Initializes a new instance of the <see cref="DrawingPropertiesWindow"/> class.
        /// </summary>
        /// <param name="visualEditor">The visual editor.</param>
        /// <param name="drawing">The drawing.</param>
        public DrawingPropertiesWindow(SpreadsheetVisualEditor visualEditor, SheetDrawing drawing)
            : this()
        {
            _visualEditor = visualEditor;
            _drawing = drawing;

            // drawing name
            nameTextBox.Text = drawing.Name;
            // drawing description
            descriptionTextBox.Text = drawing.Description;
            // drawing rotation
            rotationAngleNumericUpDown.Value = drawing.Rotation;
            rotationAngleNumericUpDown.IsEnabled = drawing.Type != DrawingType.Chart;

            // drawing location
            sheetDrawingLocationEditorControl.Worksheet = visualEditor.FocusedWorksheet;
            sheetDrawingLocationEditorControl.SheetDrawingLocation = drawing.Location;
        }

        #endregion



        #region Methods

        /// <summary>
        /// When overridden in a derived class, is invoked whenever application code or internal 
        /// processes call <see cref="System.Windows.FrameworkElement.ApplyTemplate" />.
        /// </summary>
        public override void OnApplyTemplate()
        {
            base.OnApplyTemplate();

            descriptionTextBox.Focus();
        }

        /// <summary>
        /// Handles the Click event of OkButton object.
        /// </summary>
        private void okButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                CheckDrawingName();

                // create worksheet editor
                WorksheetEditor worksheetEditor = _visualEditor.Editor.StartEditing(_visualEditor.FocusedWorksheet);
                try
                {
                    // create drawing editor
                    SheetDrawingEditor drawingEditor = worksheetEditor.CreateDrawingEditor(_drawing);

                    // set drawing name
                    drawingEditor.SetName(nameTextBox.Text.Trim());

                    // set drawing description
                    drawingEditor.SetDescription(descriptionTextBox.Text);

                    // set drawing description
                    if (rotationAngleNumericUpDown.IsEnabled)
                        drawingEditor.SetRotation(rotationAngleNumericUpDown.Value);

                    // if drawing location is changed
                    if (!Equals(_drawing.Location, sheetDrawingLocationEditorControl.SheetDrawingLocation))
                        // set the drawing location
                        drawingEditor.SetLocation(sheetDrawingLocationEditorControl.SheetDrawingLocation);
                }
                finally
                {
                    _visualEditor.Editor.FinishEditing();
                }
            }
            catch (Exception ex)
            {
                DemosTools.ShowWarningMessage("Spreadsheet Editor Demo", ex.Message);
                return;
            }

            DialogResult = true;
        }

        /// <summary>
        /// Checks the drawing name.
        /// </summary>
        private void CheckDrawingName()
        {
            // get the drawing name
            string name = nameTextBox.Text;

            if (name != null)
                name = name.Trim();

            // if drawing name is empty
            if (string.IsNullOrEmpty(name))
                throw new InvalidOperationException("The drawing name cannot be empty.");

            foreach (SheetDrawing drawing in _visualEditor.FocusedWorksheet.Drawings)
            {
                if (drawing == _drawing)
                    continue;

                // if drawing name exists already
                if (string.Equals(drawing.Name, name, StringComparison.InvariantCultureIgnoreCase))
                    throw new InvalidOperationException(string.Format("The drawing with '{0}' exists already.", name));
            }
        }

        #endregion

    }
}

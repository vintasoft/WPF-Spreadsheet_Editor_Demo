using System;
using System.ComponentModel;
using System.Globalization;
using System.Windows;
using Vintasoft.Imaging.Office.Spreadsheet.Document;
using Vintasoft.Imaging.Office.Spreadsheet.UI;
using Vintasoft.Imaging.Wpf;
using Vintasoft.Primitives;
using WpfDemosCommonCode;

namespace WpfSpreadsheetEditorDemo
{
    /// <summary>
    /// A window that allows to add or edit the comment.
    /// </summary>
    public partial class EditCommentWindow : Window
    {


        #region Fields

        /// <summary>
        /// Spreadsheet visual editor.
        /// </summary>
        SpreadsheetVisualEditor _visualEditor;

        /// <summary>
        /// The last added comment.
        /// </summary>
        static Comment _lastAddedComment;

        /// <summary>
        /// Indicates whether a new comment is being added.
        /// </summary>
        bool _addNewComment = false;

        #endregion



        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="EditCommentForm"/> class.
        /// </summary>
        public EditCommentWindow()
        {
            InitializeComponent();

            // text align

            textHorizontalAlignmentComboBox.Items.Add(TextHorizontalAlign.Left);
            textHorizontalAlignmentComboBox.Items.Add(TextHorizontalAlign.Center);
            textHorizontalAlignmentComboBox.Items.Add(TextHorizontalAlign.Right);

            textVerticalAlignmentComboBox.Items.Add(TextVerticalAlign.Top);
            textVerticalAlignmentComboBox.Items.Add(TextVerticalAlign.Middle);
            textVerticalAlignmentComboBox.Items.Add(TextVerticalAlign.Bottom);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="EditCommentForm"/> class.
        /// </summary>
        /// <param name="visualEditor">The spreadsheet visual editor.</param>
        public EditCommentWindow(SpreadsheetVisualEditor visualEditor)
            : this()
        {
            _visualEditor = visualEditor;

            sheetDrawingLocationEditorControl.Worksheet = _visualEditor.FocusedWorksheet;

            // create comment properties
            TextProperties textProperties = new TextProperties(TextHorizontalAlign.Left, TextVerticalAlign.Top, true, 0);
            FontProperties fontProperties = _visualEditor.Document.Styles[0].FontProperties;

            // create comment
            if (_lastAddedComment != null)
            {
                _comment = new Comment(
                    _lastAddedComment.Author,
                    _lastAddedComment.ShowAuthor,
                    _lastAddedComment.Color,
                    "",
                    textProperties,
                    _lastAddedComment.FontProperties);
            }
            else
            {
                _comment = new Comment(
                    Environment.UserName,
                    true,
                    VintasoftColor.FromRgb(255, 255, 225),
                    "",
                    textProperties,
                    fontProperties);
            }

            // create new comment location
            CellReference anchorCell = _visualEditor.FocusedWorksheet.GetVisibleCellOnRight(_visualEditor.FocusedCell); 
            double offset = 10;
            double xOffset = offset;
            double yOffset = anchorCell.RowIndex == 0 ? offset : -offset;
            VintasoftPoint topLeftOffset = new VintasoftPoint(xOffset, yOffset);
            VintasoftPoint bottomRightOffset = new VintasoftPoint(170, 70);
            _commentLocation = new SheetDrawingLocation(anchorCell, topLeftOffset, null, bottomRightOffset);

            _addNewComment = true;
            this.Title = "Add Comment";

            UpdateUI();
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="EditCommentForm"/> class.
        /// </summary>
        /// <param name="visualEditor">The spreadsheet visual editor.</param>
        /// <param name="comment">The comment that should be edited.</param>
        /// <param name="commentLocation">The location of comment.</param>
        public EditCommentWindow(SpreadsheetVisualEditor visualEditor, Comment comment, SheetDrawingLocation commentLocation)
            : this()
        {
            _visualEditor = visualEditor;

            sheetDrawingLocationEditorControl.Worksheet = _visualEditor.FocusedWorksheet;
            _comment = comment;
            _commentLocation = commentLocation;

            UpdateUI();
        }

        #endregion



        #region Properties

        Comment _comment;
        /// <summary>
        /// Gets the comment to edit.
        /// </summary>
        public Comment Comment
        {
            get
            {
                return _comment;
            }
        }

        SheetDrawingLocation _commentLocation;
        /// <summary>
        /// Gets the comment location.
        /// </summary>
        public SheetDrawingLocation CommentLocation
        {
            get
            {
                return _commentLocation;
            }
        }

        /// <summary>
        /// Gets the current culture.
        /// </summary>
        [Browsable(false)]
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
        /// Handles the Loaded event of Window object.
        /// </summary>
        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            commentTextBox.Focus();
        }

        /// <summary>
        /// Updates the user interface of this form.
        /// </summary>
        private void UpdateUI()
        {
            // common properties
            authorTextBox.Text = _comment.Author;
            showAuthorCheckBox.IsChecked = _comment.ShowAuthor;
            commentColorPanelControl.Color = WpfObjectConverter.Convert(_comment.Color);
            commentTextBox.Text = _comment.Text;

            // text align
            textHorizontalAlignmentComboBox.SelectedItem = _comment.TextProperties.HorizontalAlign;
            textVerticalAlignmentComboBox.SelectedItem = _comment.TextProperties.VerticalAlign;

            // font properties
            fontPropertiesEditorControl.Culture = Culture;
            fontPropertiesEditorControl.FontProperties = _comment.FontProperties;
            fontPropertiesEditorControl.NormalFontProperties = _visualEditor.Document.Styles[0].FontProperties;

            // comment location
            sheetDrawingLocationEditorControl.SheetDrawingLocation = _commentLocation;
        }

        /// <summary>
        /// Handles the Click event of okButton object.
        /// </summary>
        private void okButton_Click(object sender, EventArgs e)
        {
            try
            {
                // get font properties
                FontProperties fontProperties = fontPropertiesEditorControl.FontProperties;
                // get alignment properties
                TextProperties textProperties = new TextProperties(
                        (TextHorizontalAlign)textHorizontalAlignmentComboBox.SelectedItem,
                        (TextVerticalAlign)textVerticalAlignmentComboBox.SelectedItem,
                        true, 0);

                VintasoftColor color = WpfObjectConverter.Convert(commentColorPanelControl.Color);

                // create comment
                bool showAuthor = showAuthorCheckBox.IsChecked != null && showAuthorCheckBox.IsChecked.Value == true;
                _comment = new Comment(
                    authorTextBox.Text,
                    showAuthor,
                    color,
                    commentTextBox.Text,
                    textProperties,
                    fontProperties);

                // save the comment
                if (_addNewComment)
                    _lastAddedComment = _comment;

                // get comment location
                _commentLocation = sheetDrawingLocationEditorControl.SheetDrawingLocation;
            }
            catch (Exception ex)
            {
                DemosTools.ShowWarningMessage("Spreadsheet Editor Demo", ex.Message);
                return;
            }

            DialogResult = true;
        }

        /// <summary>
        /// Handles the Click event of cancelButton object.
        /// </summary>
        private void cancelButton_Click(object sender, EventArgs e)
        {
            DialogResult = false;
        }

        #endregion

    }
}

using System;
using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;

using Vintasoft.Imaging.Office.Spreadsheet;
using Vintasoft.Imaging.Office.Spreadsheet.Document;
using Vintasoft.Imaging.Office.Spreadsheet.UI;

namespace WpfSpreadsheetEditorDemo
{
    /// <summary>
    /// "Find and Replace" window.
    /// </summary>
    public partial class FindReplaceWindow : Window
    {

        #region Fields

        /// <summary>
        /// A value indicating whether <see cref="_textFindReplace"/> must be reset.
        /// </summary>
        bool _needReset = true;

        /// <summary>
        /// A value indicating whether the window must be closed.
        /// </summary>
        bool _needClose = false;

        /// <summary>
        /// The text find and replace engine.
        /// </summary>
        SpreadsheetFindReplace _textFindReplace;

        /// <summary>
        /// The current results.
        /// </summary>
        SpreadsheetCellReference[] _currentResults;

        bool _updateFocusedCell = false;

        #endregion



        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="FindReplaceWindow"/> class.
        /// </summary>
        public FindReplaceWindow()
        {
            InitializeComponent();

            findWithinComboBox.SelectedIndex = 0;
            lookInComboBox.SelectedIndex = 0;
            searchComboBox.SelectedIndex = 0;

            Height = MinHeight;

            UpdateUI();
        }

        #endregion



        #region Properties

        SpreadsheetVisualEditor _visualEditor;
        /// <summary>
        /// Gets or sets the visual editor.
        /// </summary>
        public SpreadsheetVisualEditor VisualEditor
        {
            get
            {
                return _visualEditor;
            }
            set
            {
                if (_visualEditor != null)
                {
                    _visualEditor.SelectedCells.Changed -= SelectedCells_Changed;
                    _visualEditor.FocusedWorksheetChanged -= visualEditor_FocusedWorksheetChanged;
                }

                _visualEditor = value;

                if (_visualEditor != null)
                {
                    _visualEditor.SelectedCells.Changed += SelectedCells_Changed;
                    _visualEditor.FocusedWorksheetChanged += visualEditor_FocusedWorksheetChanged;
                }
                _textFindReplace = null;
                UpdateUI();
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether replace mode is used.
        /// </summary>
        public bool ReplaceMode
        {
            get
            {
                return replaceCheckBox.IsChecked.Value == true;
            }
            set
            {
                replaceCheckBox.IsChecked = value;
            }
        }

        #endregion



        #region Methods

        /// <summary>
        /// Resets the dialog.
        /// </summary>
        public void Reset()
        {
            _needReset = true;
            findWhatComboBox.Focus();
        }

        /// <summary>
        /// Manually closes the <see cref="System.Windows.Window" />.
        /// </summary>
        public new void Close()
        {
            _needClose = true;
            base.Close();
        }

        /// <summary>
        /// Updates the user interface.
        /// </summary>
        private void UpdateUI()
        {
            IsEnabled = VisualEditor != null;
            Visibility visibility = ReplaceMode ? Visibility.Visible : Visibility.Hidden;

            replaceAllButton.Visibility = visibility;
            replaceButton.Visibility = visibility;
            replaceLabel.Visibility = visibility;
            replaceWithComboBox.Visibility = visibility;
            addToSelectionButton.IsEnabled = _currentResults != null && _currentResults.Length > 0;
            if (ReplaceMode)
            {
                lookInComboBox.IsEnabled = false;
                lookInComboBox.SelectedIndex = 0;
            }
            else
            {
                lookInComboBox.IsEnabled = true;
            }
            UpdateStatus();

            bool hasWorksheet = VisualEditor != null && VisualEditor.FocusedWorksheet != null;
            replaceAllButton.IsEnabled = hasWorksheet;
            replaceButton.IsEnabled = hasWorksheet;
            findAllButton.IsEnabled = hasWorksheet;
            findNextButton.IsEnabled = hasWorksheet;
        }

        /// <summary>
        /// Handles the Click event of closeButton object.
        /// </summary>
        private void closeButton_Click(object sender, RoutedEventArgs e)
        {
            Visibility = Visibility.Hidden;
            Height = MinHeight;
        }

        protected override void OnClosing(CancelEventArgs e)
        {
            base.OnClosing(e);

            if (!_needClose)
            {
                e.Cancel = true;
                Visibility = Visibility.Hidden;
                Height = MinHeight;
            }
        }

        /// <summary>
        /// Handles the CheckedChanged event of replaceCheckBox object.
        /// </summary>
        private void replaceCheckBox_CheckedChanged(object sender, RoutedEventArgs e)
        {
            UpdateUI();
        }

        /// <summary>
        /// Handles the Click event of findNextButton object.
        /// </summary>
        private void findNextButton_Click(object sender, RoutedEventArgs e)
        {
            FindNext();
        }

        /// <summary>
        /// Handles the Click event of addToSelectionButton object.
        /// </summary>
        private void addToSelectionButton_Click(object sender, RoutedEventArgs e)
        {
            VisualEditor.SetFocusedAndSelectedCells(_currentResults);
        }

        /// <summary>
        /// Handles the Click event of findAllButton object.
        /// </summary>
        private void findAllButton_Click(object sender, RoutedEventArgs e)
        {
            SpreadsheetFindReplace textReplacer = GetTextFindReplace();
            if (textReplacer != null)
            {
                if (Height < MinHeight + 50)
                    Height = MinHeight + 150;

                ShowResults(textReplacer.FindAll());
                UpdateStatus();
                if (textReplacer.FoundCellCount == 0)
                    MessageBox.Show(string.Format("Text '{0}' is not found.", textReplacer.Text));
            }
        }

        /// <summary>
        /// Handles the Click event of replaceButton object.
        /// </summary>
        private void replaceButton_Click(object sender, RoutedEventArgs e)
        {
            SpreadsheetFindReplace textReplacer = GetTextFindReplace();
            if (textReplacer != null)
            {
                bool result;
                VisualEditor.StartEditing("Replace");
                try
                {
                    result = textReplacer.Replace(replaceWithComboBox.Text);
                }
                finally
                {
                    VisualEditor.FinishEditing();
                }
                if (result)
                    FindNext();
                else
                    MessageBox.Show("Cannot find a match.");
            }
        }

        /// <summary>
        /// Handles the Click event of replaceAllButton object.
        /// </summary>
        private void replaceAllButton_Click(object sender, RoutedEventArgs e)
        {
            SpreadsheetFindReplace textReplacer = GetTextFindReplace();
            if (textReplacer != null)
            {
                if (Height < MinHeight + 50)
                    Height = MinHeight + 150;
                
                SpreadsheetCellReference[] result;
                VisualEditor.StartEditing("Replace all");
                try
                {
                    result = textReplacer.ReplaceAll(replaceWithComboBox.Text);
                }
                finally
                {
                    VisualEditor.FinishEditing();
                }

                ShowResults(result);

                UpdateStatus();
                if (textReplacer.FoundCellCount == 0)
                    MessageBox.Show("Could not find anything to replace.", "Replace", MessageBoxButton.OK, MessageBoxImage.Warning);
                else
                    MessageBox.Show(string.Format("Made '{0}' replacements.", textReplacer.FoundCellCount), "Replace", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        /// <summary>
        /// Handles the TextChanged event of findWhatComboBox object.
        /// </summary>
        private void findWhatComboBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            UpdateUI();
        }

        /// <summary>
        /// Handles the SelectionChanged event of resultListView object.
        /// </summary>
        private void resultListView_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (resultListView.SelectedItem != null)
            {
                SpreadsheetCellReference workbookCellReference = (SpreadsheetCellReference)resultListView.SelectedItem;
                if (_textFindReplace.SelectedCells != null)
                    VisualEditor.FocusedSpreadsheetCell = workbookCellReference;
                else
                    VisualEditor.SetFocusedAndSelectedCells(workbookCellReference);
            }
        }

        /// <summary>
        /// Handles the SelectionChanged event of findWithinComboBox object.
        /// </summary>
        private void findWithinComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            _needReset = true;
        }

        /// <summary>
        /// Handles the SelectionChanged event of searchComboBox object.
        /// </summary>
        private void searchComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            _needReset = true;
        }

        /// <summary>
        /// Handles the SelectionChanged event of lookInComboBox object.
        /// </summary>
        private void lookInComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            _needReset = true;
        }

        /// <summary>
        /// Handles the Changed event of SelectedCells object.
        /// </summary>
        private void SelectedCells_Changed(object sender, EventArgs e)
        {
            if (!_updateFocusedCell && _textFindReplace != null)
                _needReset = true;
        }

        /// <summary>
        /// Handles the FocusedWorksheetChanged event of visualEditor object.
        /// </summary>
        private void visualEditor_FocusedWorksheetChanged(object sender, Vintasoft.Imaging.PropertyChangedEventArgs<Worksheet> e)
        {
            if (e.NewValue == null)
            {
                ShowResults(null);
                Height = MinHeight;
            }

            UpdateUI();
        }

        /// <summary>
        /// Finds the next cell.
        /// </summary>
        private void FindNext()
        {
            ShowResults(null);
            SpreadsheetFindReplace textReplacer = GetTextFindReplace();
            bool found = false;
            if (textReplacer != null)
            {
                found = textReplacer.FindNext();
                UpdateStatus();
                if (!found)
                {
                    if (textReplacer.FoundCellCount == 0)
                    {
                        MessageBox.Show(string.Format("Text '{0}' is not found.", textReplacer.Text));
                    }
                    else
                    {
                        // find again
                        found = textReplacer.FindNext();
                    }
                }
            }
            if (found)
            {
                _updateFocusedCell = true;
                VisualEditor.FocusedSpreadsheetCell = textReplacer.GetCurrentSpreadsheetCell();
                _updateFocusedCell = false;
            }
        }

        /// <summary>
        /// Updates the status.
        /// </summary>
        private void UpdateStatus()
        {
            if (_textFindReplace == null || _textFindReplace.FoundCellCount == 0)
                statusLabel.Content = "";
            else
                statusLabel.Content = string.Format("{0} cells found.", _textFindReplace.FoundCellCount);
        }

        /// <summary>
        /// Shows the results.
        /// </summary>
        /// <param name="cellReferences">The cell references.</param>
        private void ShowResults(SpreadsheetCellReference[] cellReferences)
        {
            resultListView.Items.Clear();
            _currentResults = cellReferences;
            if (cellReferences != null)
            {
                addToSelectionButton.IsEnabled = cellReferences.Length > 0;

                if (_textFindReplace.LookInValues)
                    valueGridViewColumn.DisplayMemberBinding = new Binding("FormattedValue");
                else
                    valueGridViewColumn.DisplayMemberBinding = new Binding("Value");

                foreach (SpreadsheetCellReference cellReference in cellReferences)
                    resultListView.Items.Add(cellReference);
            }
            else
            {
                addToSelectionButton.IsEnabled = false;
            }
        }

        /// <summary>
        /// Returns the text find replace engine.
        /// </summary>
        private SpreadsheetFindReplace GetTextFindReplace()
        {
            VisualEditor.FinishEditCellValue();

            if (_textFindReplace == null)
                _textFindReplace = new SpreadsheetFindReplace();

            SetTextFindReplaceProperties(_textFindReplace);

            if (_textFindReplace.SearchInWorkbook)
            {
                _updateFocusedCell = true;
                _visualEditor.SetFocusedAndSelectedCells(new CellReferences(_visualEditor.FocusedCell));
                _updateFocusedCell = false;
            }

            string text = findWhatComboBox.Text;
            if (_needReset || _textFindReplace.Editor != VisualEditor.Editor || _textFindReplace.Text != text)
            {
                VisualEditor.InitializeFindReplace(_textFindReplace, text);
                _needReset = false;
            }

            if (!string.IsNullOrEmpty(findWhatComboBox.Text) && !findWhatComboBox.Items.Contains(findWhatComboBox.Text))
                findWhatComboBox.Items.Insert(0, findWhatComboBox.Text);

            if (!string.IsNullOrEmpty(replaceWithComboBox.Text) && !replaceWithComboBox.Items.Contains(replaceWithComboBox.Text))
                replaceWithComboBox.Items.Insert(0, replaceWithComboBox.Text);

            return _textFindReplace;
        }

        /// <summary>
        /// Sets the properties of text find and replace.
        /// </summary>
        /// <param name="textFindReplace">The text replacer.</param>
        private void SetTextFindReplaceProperties(SpreadsheetFindReplace textFindReplace)
        {
            textFindReplace.MatchCase = matchCaseCheckBox.IsChecked.Value == true;
            textFindReplace.MatchEntireContent = matchContentsCheckBox.IsChecked.Value == true;
            switch (findWithinComboBox.SelectedIndex)
            {
                case 0:
                    textFindReplace.SearchInWorkbook = false;
                    break;
                case 1:
                    textFindReplace.SearchInWorkbook = true;
                    break;
                default:
                    throw new NotImplementedException();
            }
            switch (searchComboBox.SelectedIndex)
            {
                case 0:
                    textFindReplace.SearchByRows = true;
                    break;
                case 1:
                    textFindReplace.SearchByRows = false;
                    break;
                default:
                    throw new NotImplementedException();
            }
            switch (lookInComboBox.SelectedIndex)
            {
                case 0:
                    textFindReplace.LookInFormulas = true;
                    textFindReplace.LookInValues = false;
                    textFindReplace.LookInComments = false;
                    break;
                case 1:
                    textFindReplace.LookInFormulas = false;
                    textFindReplace.LookInValues = true;
                    textFindReplace.LookInComments = false;
                    break;
                case 2:
                    textFindReplace.LookInFormulas = false;
                    textFindReplace.LookInValues = false;
                    textFindReplace.LookInComments = true;
                    break;
                default:
                    throw new NotImplementedException();
            }
        }

        #endregion

    }
}

using System.Windows;
using System.Windows.Controls;

using Vintasoft.Imaging;
using Vintasoft.Imaging.Office.Spreadsheet;
using Vintasoft.Imaging.Office.Spreadsheet.Wpf.UI;

namespace WpfSpreadsheetEditorDemo
{
    /// <summary>
    /// Provides an "Undo" panel.
    /// </summary>
    public partial class UndoPanel : SpreadsheetVisualEditorPanel
    {

        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="UndoPanel"/> class.
        /// </summary>
        public UndoPanel()
        {
            InitializeComponent();
        }

        #endregion



        #region Methods

        /// <summary>
        /// Raises the <see cref="E:SpreadsheetEditorChanged" /> event.
        /// </summary>
        /// <param name="args">The <see cref="PropertyChangedEventArgs{SpreadsheetEditorControl}"/> instance containing the event data.</param>
        protected override void OnSpreadsheetEditorChanged(PropertyChangedEventArgs<WpfSpreadsheetEditorControl> args)
        {
            base.OnSpreadsheetEditorChanged(args);

            if (args.OldValue != null)
            {
                args.OldValue.VisualEditor.UndoManagerStateChanged -= VisualEditor_UndoManagerStateChanged;
            }

            if (args.NewValue != null)
            {
                args.NewValue.VisualEditor.UndoManagerStateChanged += VisualEditor_UndoManagerStateChanged;
            }

            UpdateUI();
        }

        /// <summary>
        /// Handles the UndoManagerStateChanged event of the VisualEditor.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="System.EventArgs"/> instance containing the event data.</param>
        private void VisualEditor_UndoManagerStateChanged(object sender, System.EventArgs e)
        {
            UpdateUI();
        }

        /// <summary>
        /// Updates the user interface.
        /// </summary>
        protected override void UpdateCoreUI()
        {
            base.UpdateCoreUI();
            if (IsEnabled)
                UpdateUI();
        }

        /// <summary>
        /// Updates the user interface.
        /// </summary>
        private void UpdateUI()
        {
            if (VisualEditor.IsFocusedWorksheetChanging)
                return;

            if (VisualEditor.FocusedWorksheet == null)
            {
                IsEnabled = false;
            }
            else
            {
                IsEnabled = true;

                if (VisualEditor.CanUndo)
                {
                    undoButton.IsEnabled = true;
                    undoButton.ToolTip = VisualEditor.UndoManager.CurrentUndoAction.Description;
                }
                else
                {
                    undoButton.IsEnabled = false;
                }
                if (VisualEditor.CanRedo)
                {
                    redoButton.IsEnabled = true;
                    redoButton.ToolTip = VisualEditor.UndoManager.CurrentRedoAction.Description;
                }
                else
                {
                    redoButton.IsEnabled = false;
                }
            }
        }

        private void undoButton_ButtonClick(object sender, RoutedEventArgs e)
        {
            VisualEditor.UndoManager.Undo();
            SpreadsheetEditor.Focus();
        }

        private void redoButton_ButtonClick(object sender, RoutedEventArgs e)
        {
            VisualEditor.UndoManager.Redo();
            SpreadsheetEditor.Focus();
        }

        /// <summary>
        /// Handles the Click event of UndoItem object.
        /// </summary>
        private void UndoItem_Click(object sender, RoutedEventArgs e)
        {
            VisualEditor.UndoManager.Undo((int)((MenuItem)sender).Tag);
        }


        /// <summary>
        /// Handles the Click event of RedoItem object.
        /// </summary>
        private void RedoItem_Click(object sender, RoutedEventArgs e)
        {
            VisualEditor.UndoManager.Redo((int)((MenuItem)sender).Tag);
        }

        /// <summary>
        /// Handles the SubmenuOpened event of UndoButton object.
        /// </summary>
        private void undoButton_SubmenuOpened(object sender, RoutedEventArgs e)
        {
            SpreadsheetEditorUndoAction[] availableUndoActions = VisualEditor.UndoManager.GetAvailableUndoActions();
            undoButton.Items.Clear();
            for (int i = 0; i < availableUndoActions.Length; i++)
            {
                MenuItem item = new MenuItem();
                undoButton.Items.Add(item);
                item.Header = availableUndoActions[i].Description;
                item.Tag = i + 1;
                item.Click += UndoItem_Click;
            }
        }

        /// <summary>
        /// Handles the SubmenuOpened event of RedoButton object.
        /// </summary>
        private void redoButton_SubmenuOpened(object sender, RoutedEventArgs e)
        {
            SpreadsheetEditorUndoAction[] availableRedoActions = VisualEditor.UndoManager.GetAvailableRedoActions();
            redoButton.Items.Clear();
            for (int i = 0; i < availableRedoActions.Length; i++)
            {
                MenuItem item = new MenuItem();
                redoButton.Items.Add(item);
                item.Header = availableRedoActions[i].Description;
                item.Tag = i + 1;
                item.Click += RedoItem_Click;
            }
        }

        #endregion

    }
}

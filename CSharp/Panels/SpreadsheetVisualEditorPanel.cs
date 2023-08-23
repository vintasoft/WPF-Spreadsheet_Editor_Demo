using System;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;

using Vintasoft.Imaging;
using Vintasoft.Imaging.Office.Spreadsheet.UI;
using Vintasoft.Imaging.Office.Spreadsheet.Wpf.UI;

namespace WpfSpreadsheetEditorDemo
{
    /// <summary> 
    /// Represents a base UI panel for <see cref="SpreadsheetVisualEditor"/>.
    /// </summary>
    public class SpreadsheetVisualEditorPanel : UserControl
    {

        #region Fields

        /// <summary> 
        /// Identifies the <see cref="SpreadsheetEditor"/> dependency property.
        /// </summary> 
        public static readonly DependencyProperty SpreadsheetEditorProperty =
            DependencyProperty.Register("SpreadsheetEditor", typeof(WpfSpreadsheetEditorControl), typeof(SpreadsheetVisualEditorPanel), new FrameworkPropertyMetadata(null, new PropertyChangedCallback(OnSpreadsheetEditorPropertyValueChanged), null));

        #endregion



        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="SpreadsheetVisualEditorPanel"/> class.
        /// </summary>
        public SpreadsheetVisualEditorPanel()
        {
            IsEnabled = false;
        }

        #endregion



        #region Properties

        /// <summary> 
        /// Gets or sets the value assigned to the control. 
        /// </summary> 
        public WpfSpreadsheetEditorControl SpreadsheetEditor
        {
            get
            {
                return (WpfSpreadsheetEditorControl)GetValue(SpreadsheetEditorProperty);
            }
            set
            {
                SetValue(SpreadsheetEditorProperty, value);
            }
        }

        /// <summary>
        /// Gets a value indicating whether this panel is disabled without editor.
        /// </summary>
        protected virtual bool IsDisabledWithoutEditor
        {
            get
            {
                return true;
            }
        }

        /// <summary>
        /// Gets the spreadsheet visual editor.
        /// </summary>    
        public SpreadsheetVisualEditor VisualEditor
        {
            get
            {
                if (SpreadsheetEditor != null)
                    return SpreadsheetEditor.VisualEditor;
                return null;
            }
        }

        /// <summary>
        /// Gets the document current culture.
        /// </summary>
        public CultureInfo Culture
        {
            get
            {
                if (VisualEditor != null)
                {
                    try
                    {
                        return CultureInfo.GetCultureInfo(VisualEditor.DocumentCulture);
                    }
                    catch
                    {
                    }
                }
                return CultureInfo.CurrentCulture;
            }
        }

        /// <summary>
        /// Gets the document UI current culture.
        /// </summary>      
        public CultureInfo UICulture
        {
            get
            {
                if (VisualEditor != null)
                {
                    try
                    {
                        return CultureInfo.GetCultureInfo(VisualEditor.DocumentUICulture);
                    }
                    catch
                    {
                    }
                }
                return CultureInfo.CurrentUICulture;
            }
        }

        #endregion



        #region Methods

        /// <summary>
        /// Occurs when <see cref="SpreadsheetEditor"/> is changed.
        /// </summary>
        /// <param name="args">The <see cref="PropertyChangedEventArgs{WpfSpreadsheetEditorControl}"/> instance containing the event data.</param>
        protected virtual void OnSpreadsheetEditorChanged(PropertyChangedEventArgs<WpfSpreadsheetEditorControl> args)
        {
        }

        /// <summary>
        /// Called when <see cref="SpreadsheetEditor"/> property is changed.
        /// </summary>
        /// <param name="obj">The object.</param>
        /// <param name="args">The <see cref="DependencyPropertyChangedEventArgs"/> instance containing the event data.</param>
        private static void OnSpreadsheetEditorPropertyValueChanged(DependencyObject obj, DependencyPropertyChangedEventArgs args)
        {
            SpreadsheetVisualEditorPanel panel = (SpreadsheetVisualEditorPanel)obj;

            PropertyChangedEventArgs<WpfSpreadsheetEditorControl> changedArgs = new PropertyChangedEventArgs<WpfSpreadsheetEditorControl>(
                (WpfSpreadsheetEditorControl)args.OldValue,
                (WpfSpreadsheetEditorControl)args.NewValue);

            if (changedArgs.OldValue != null)
            {
                changedArgs.OldValue.VisualEditor.EditorChanged -= panel.VisualEditor_EditorChanged;
                changedArgs.OldValue.VisualEditor.InitializationStarted -= panel.VisualEditor_InitializationStarted;
                changedArgs.OldValue.VisualEditor.InitializationFinished -= panel.VisualEditor_InitializationFinished;
            }

            if (changedArgs.NewValue != null)
            {
                changedArgs.NewValue.VisualEditor.EditorChanged += panel.VisualEditor_EditorChanged;
                changedArgs.NewValue.VisualEditor.InitializationStarted += panel.VisualEditor_InitializationStarted;
                changedArgs.NewValue.VisualEditor.InitializationFinished += panel.VisualEditor_InitializationFinished;
            }

            panel.OnSpreadsheetEditorChanged(changedArgs);
            panel.UpdateCoreUI();
        }

        private void VisualEditor_EditorChanged(object sender, PropertyChangedEventArgs<Vintasoft.Imaging.Office.Spreadsheet.SpreadsheetEditor> e)
        {
            UpdateCoreUI();
        }

        /// <summary>
        /// Updates the User Interface.
        /// </summary>
        protected virtual void UpdateCoreUI()
        {
            if (SpreadsheetEditor == null || VisualEditor.IsInitializing)
            {
                IsEnabled = false;
            }
            else
            {
                if (IsDisabledWithoutEditor)
                    IsEnabled = VisualEditor.Editor != null;
                else
                    IsEnabled = true;
            }
        }

        /// <summary>
        /// Handles the InitializationFinished event of the VisualEditor.
        /// </summary>
        private void VisualEditor_InitializationFinished(object sender, EventArgs e)
        {
            UpdateCoreUI();
        }

        /// <summary>
        /// Handles the InitializationStarted event of the VisualEditor.
        /// </summary>
        private void VisualEditor_InitializationStarted(object sender, EventArgs e)
        {
            UpdateCoreUI();
        }

        #endregion

    }
}

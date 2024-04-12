using System;
using System.ComponentModel;
using System.IO;
using System.Windows;
using System.Windows.Controls;

using Vintasoft.Imaging;
using Vintasoft.Imaging.Wpf.Print;

using WpfDemosCommonCode.Imaging;

namespace WpfSpreadsheetEditorDemo
{
    /// <summary>
    /// A window that allows to preview and print XLSX document.
    /// </summary>
    public partial class PrintPreviewWindow : Window
    {

        #region Fields
        
        /// <summary>
        /// Print manager.
        /// </summary>
        WpfImagePrintManager _printManager;

        #endregion



        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="PrintPreviewWindow"/> class.
        /// </summary>
        /// <param name="fileStream">A stream that contains XLSX file.</param>
        public PrintPreviewWindow(Stream fileStream)
        {
            if (fileStream == null)
                throw new ArgumentNullException("fileStream");

            InitializeComponent();

            _printManager = new WpfImagePrintManager();
            _printManager.Images = new ImageCollection();
            _printManager.Images.Add(fileStream);
            _printManager.Preview = printPreviewControl1;
            _printManager.Preview.InputBindings.Clear();

            // initialize the page selector
            previewPageIndexNumericUpDown.Minimum = 1;
            previewPageIndexNumericUpDown.Maximum = _printManager.Images.Count;
            previewPageCountLabel.Content = string.Format("from {0} pages", _printManager.Images.Count);

            // set 100% zoom in preview
            previewZoomComboBox.SelectedIndex = 3;
        }

        #endregion



        #region Methods

        /// <summary>
        /// Raises the <see cref="System.Windows.Window.Closing" /> event.
        /// </summary>
        /// <param name="e">A <see cref="System.ComponentModel.CancelEventArgs" /> that contains the event data.</param>
        protected override void OnClosing(CancelEventArgs e)
        {
            base.OnClosing(e);

            _printManager.Images.ClearAndDisposeItems();

            _printManager.Dispose();
            _printManager = null;
        }


        /// <summary>
        /// "Print" button is clicked.
        /// </summary>
        private void printButton_Click(object sender, RoutedEventArgs e)
        {
            using (WpfImagePrintManager imagePrintManager = new WpfImagePrintManager())
            {
                imagePrintManager.Images = _printManager.Images;
                imagePrintManager.PagePadding = _printManager.PagePadding;
                imagePrintManager.ImagePadding = _printManager.ImagePadding;
                imagePrintManager.PrintDialog.MinPage = 1;
                imagePrintManager.PrintDialog.MaxPage = (uint)_printManager.Images.Count;

                if (imagePrintManager.PrintDialog.ShowDialog() == true)
                    // print XLSX document
                    imagePrintManager.Print("XLSX document");
            }
        }

        /// <summary>
        /// "Page settings" button is clicked.
        /// </summary>
        private void pageSettingsButton_Click(object sender, RoutedEventArgs e)
        {
            PageSettingsWindow pageSettingsWindow = new PageSettingsWindow(
                _printManager, _printManager.PagePadding, _printManager.ImagePadding);
            pageSettingsWindow.WindowStartupLocation = WindowStartupLocation.CenterOwner;
            pageSettingsWindow.Owner = Application.Current.MainWindow;
            if (pageSettingsWindow.ShowDialog() == true)
            {
                _printManager.PagePadding = pageSettingsWindow.PagePadding;
                _printManager.ImagePadding = pageSettingsWindow.ImagePadding;
            }
        }

        /// <summary>
        /// Preview page is changed.
        /// </summary>
        private void previewPageIndexNumericUpDown_ValueChanged(object sender, EventArgs e)
        {
            _printManager.PreviewFirstPageIndex = (int)previewPageIndexNumericUpDown.Value - 1;
        }

        /// <summary>
        /// Preview zoom is changed.
        /// </summary>
        private void previewZoomComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            double zoom = _printManager.PreviewZoom;

            switch (previewZoomComboBox.SelectedIndex)
            {
                case 0:
                    zoom = 0.2;
                    break;
                case 1:
                    zoom = 0.5;
                    break;
                case 2:
                    zoom = 0.75;
                    break;
                case 3:
                    zoom = 1;
                    break;
            }

            _printManager.PreviewZoom = zoom;
        }

        #endregion

    }
}

using System;
using System.Globalization;
using System.IO;
using System.Windows;
using System.Windows.Controls;

using Microsoft.Win32;

using Vintasoft.Imaging;
using Vintasoft.Imaging.Codecs.Decoders;
using Vintasoft.Imaging.Office.OpenXml;
using Vintasoft.Imaging.Office.Spreadsheet;
using Vintasoft.Imaging.Office.Spreadsheet.Document;
using Vintasoft.Imaging.Office.Spreadsheet.UI;
using Vintasoft.Imaging.Office.Spreadsheet.Wpf.UI;
using Vintasoft.Imaging.Wpf.Print;

using WpfDemosCommonCode;
using WpfDemosCommonCode.Imaging;
using WpfDemosCommonCode.Imaging.Codecs;

namespace WpfSpreadsheetEditorDemo
{
    /// <summary>
    /// Provides the "File" panel.
    /// </summary>
    public partial class FilePanel : SpreadsheetVisualEditorPanel
    {

        #region Fields

        /// <summary>
        /// A value indicating whether document is new.
        /// </summary>
        bool _isNewDocument = true;

        /// <summary>
        /// The document converter.
        /// </summary>
        DocumentConverter _converter;

        /// <summary>
        /// The layout settings manager.
        /// </summary>
        ImageCollectionXlsxLayoutSettingsManager _layoutSettingsManager;

        /// <summary>
        /// A value indicating whether the layout settings are initialized.
        /// </summary>
        bool _isLayoutSettingsInitialized = false;

        /// <summary>
        /// The export file dialog.
        /// </summary>
        SaveFileDialog _exportFileDialog = new SaveFileDialog();

        /// <summary>
        /// The open worksheet file dialog.
        /// </summary>
        OpenFileDialog _openWorksheetFileDialog = new OpenFileDialog();

        /// <summary>
        /// The save worksheet file dialog.
        /// </summary>
        SaveFileDialog _saveWorksheetFileDialog = new SaveFileDialog();

        /// <summary>
        /// The print manager.
        /// </summary>
        WpfImagePrintManager _imagePrintManager = new WpfImagePrintManager();

        #endregion



        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="FilePanel"/> class.
        /// </summary>
        public FilePanel()
        {
            InitializeComponent();

            _imagePrintManager.PrintScaleMode = Vintasoft.Imaging.Print.PrintScaleMode.BestFit;
            _converter = new DocumentConverter();

            _layoutSettingsManager = new ImageCollectionXlsxLayoutSettingsManager(_converter.Images);
            XlsxDocumentLayoutSettings layoutSettings = new XlsxDocumentLayoutSettings();
            layoutSettings.PageLayoutSettingsType = XlsxPageLayoutSettingsType.UseWorksheetWidth;
            _layoutSettingsManager.LayoutSettings = layoutSettings;


            _openWorksheetFileDialog.Filter = "XLSX files|*.xlsx|XLS files|*.xls|TSV files|*.tsv;*.tab|CSV files|*.csv|All supported Workbooks|*.xlsx;*.xls;*.tsv;*.tab;*.csv";
            _openWorksheetFileDialog.FilterIndex = 5;

            DemosTools.SetTestXlsxFolder(_openWorksheetFileDialog);

            _saveWorksheetFileDialog.Filter = "XLSX files|*.xlsx";

            CodecsFileFilters.SetFilters(_exportFileDialog, false);

            _exportFileDialog.Filter += "|TSV files|*.tsv|CSV files|*.csv";
            // set default filter index to PDF
            string[] filters = _exportFileDialog.Filter.Split('|');
            for (int i = 1; i < filters.Length; i++)
            {
                if (filters[i].ToUpperInvariant().Contains("PDF"))
                    _exportFileDialog.FilterIndex = i / 2 + 1;
            }

        }

        #endregion



        #region Properties

        string _filename;
        /// <summary>
        /// Gets the filename.
        /// </summary>
        public string Filename
        {
            get
            {
                return _filename;
            }
        }

        /// <summary>
        /// Gets a value indicating whether this panel is disabled without editor.
        /// </summary>
        protected override bool IsDisabledWithoutEditor
        {
            get
            {
                return false;
            }
        }

        #endregion



        #region Methods        

        #region PUBLIC

        /// <summary>
        /// Creates new document.
        /// </summary>
        public bool NewDocument()
        {
            if (CheckChanges())
            {
                SetFilename("NewWorksheet1.xlsx");
                VisualEditor.NewDocument();
                _isNewDocument = true;
            }
            return true;
        }

        /// <summary>
        /// Opens the spreadsheet document.
        /// </summary>
        public bool OpenDocument()
        {
            if (CheckChanges())
            {
                // close previously opened XLSX document
                VisualEditor.CloseDocument();

                // show dialog for opening the XLSX file
                if (_openWorksheetFileDialog.ShowDialog() == true)
                {
                    // get file path from open dialog
                    string filename = _openWorksheetFileDialog.FileName;
                    try
                    {
                        // if file is XLS file
                        if (XlsxDecoder.IsXlsDocument(filename))
                        {
                            if (MessageBox.Show("The loaded file is XLS file. To open XLS file application needs to convert XLS file to the XLSX file. Do you want to create XLSX file from XLS file?", "Convert XLS to XLSX", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.No)
                                return true;

                            // create path to an XLSX file
                            filename = Path.Combine(Path.GetDirectoryName(filename), Path.GetFileNameWithoutExtension(filename) + ".xlsx");
                            // set file to the save dialog
                            _saveWorksheetFileDialog.FileName = filename;
                            // show the save dialog
                            if (_saveWorksheetFileDialog.ShowDialog() != true)
                                return true;
                            // get file path from save dialog
                            filename = _saveWorksheetFileDialog.FileName;
                            // convert XLS file to the XLSX file
                            OpenXmlDocumentConverter.ConvertXlsToXlsx(_openWorksheetFileDialog.FileName, filename);
                        }
                        // if file is CSV file
                        else if (XlsxDecoder.IsCsvFile(filename))
                        {
                            if (MessageBox.Show("The loaded file is CSV file. To open CSV file application needs to convert CSV file to the XLSX file. Do you want to create XLSX file from CSV file?", "Convert CSV to XLSX", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.No)
                                return true;

                            // create path to an XLSX file
                            filename = Path.Combine(Path.GetDirectoryName(filename), Path.GetFileNameWithoutExtension(filename) + ".xlsx");
                            // set file to the save dialog
                            _saveWorksheetFileDialog.FileName = filename;
                            // show the save dialog
                            if (_saveWorksheetFileDialog.ShowDialog() != true)
                                return true;
                            // get file path from save dialog
                            filename = _saveWorksheetFileDialog.FileName;
                            // convert XLS file to the XLSX file
                            OpenXmlDocumentConverter.ConvertCsvToXlsx(_openWorksheetFileDialog.FileName, filename);
                        }
                        // if file is TSV file
                        else if (XlsxDecoder.IsTsvFile(filename))
                        {
                            if (MessageBox.Show("The loaded file is TSV file. To open TSV file application needs to convert TSV file to the XLSX file. Do you want to create XLSX file from TSV file?", "Convert TSV to XLSX", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.No)
                                return true;

                            // create path to an XLSX file
                            filename = Path.Combine(Path.GetDirectoryName(filename), Path.GetFileNameWithoutExtension(filename) + ".xlsx");
                            // set file to the save dialog
                            _saveWorksheetFileDialog.FileName = filename;
                            // show the save dialog
                            if (_saveWorksheetFileDialog.ShowDialog() != true)
                                return true;
                            // get file path from save dialog
                            filename = _saveWorksheetFileDialog.FileName;
                            // convert XLS file to the XLSX file
                            OpenXmlDocumentConverter.ConvertTsvToXlsx(_openWorksheetFileDialog.FileName, filename);
                        }

                        // save information about path to XLSX file
                        SetFilename(filename);
                        // open XLSX file
                        VisualEditor.OpenDocument(filename);
                        _isNewDocument = false;
                    }
                    catch (Exception ex)
                    {
                        DemosTools.ShowErrorMessage(ex);
                        SetFilename(null);
                    }
                }
            }

            return true;
        }

        /// <summary>
        /// Saves the document changes.
        /// </summary>
        public bool SaveDocumentChanges()
        {
            if (VisualEditor.IsDocumentSourceChanged)
            {
                try
                {
                    VisualEditor.SaveDocumentChanges();
                }
                catch (Exception ex)
                {
                    DemosTools.ShowErrorMessage(ex);
                }
                return true;
            }
            return false;
        }

        /// <summary>
        /// Closes the document.
        /// </summary>
        public bool CloseDocument(bool checkChanges)
        {
            if (VisualEditor.Document != null)
            {
                if (!checkChanges || CheckChanges())
                {
                    VisualEditor.CloseDocument();
                    SetFilename(null);
                    UpdateUI();
                }
                return true;
            }
            return false;
        }

        /// <summary>
        /// Saves the document as new file.
        /// </summary>
        public bool SaveDocumentAs()
        {
            if (VisualEditor.Document != null)
            {
                try
                {
                    SaveAs();
                }
                catch (Exception ex)
                {
                    DemosTools.ShowErrorMessage(ex);
                }
                return true;
            }
            return false;
        }

        /// <summary>
        /// Prints the document.
        /// </summary>
        public bool PrintDocument()
        {
            if (VisualEditor.Document == null)
                return false;
            try
            {
                // if layout settings are not initialized
                if (!_isLayoutSettingsInitialized)
                {
                    // set layout settings
                    if (_layoutSettingsManager.EditLayoutSettingsUseDialog(Application.Current.MainWindow))
                        _isLayoutSettingsInitialized = true;
                    else
                        return true;
                }

                // create a temporary stream
                using (MemoryStream tempStream = new MemoryStream())
                {
                    // save XLSX file to a temporary stream
                    VisualEditor.SaveDocumentTo(tempStream);

                    using (ImageCollection images = new ImageCollection())
                    {
                        images.Add(tempStream);
                        try
                        {
                            _imagePrintManager.Images = images;
                            _imagePrintManager.PrintDialog.MinPage = 1;
                            _imagePrintManager.PrintDialog.MaxPage = (uint)images.Count;

                            if (_imagePrintManager.PrintDialog.ShowDialog() == true)
                                // print XLSX document
                                _imagePrintManager.Print("XLSX document");
                        }
                        finally
                        {
                            _imagePrintManager.Images = null;
                            images.ClearAndDisposeItems();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // show error message
                DemosTools.ShowErrorMessage(ex);
            }
            return true;
        }

        #endregion


        #region PROTECTED

        /// <summary>
        /// Raises the <see cref="E:SpreadsheetEditorChanged" /> event.
        /// </summary>
        /// <param name="args">The <see cref="PropertyChangedEventArgs{SpreadsheetEditorControl}"/> instance containing the event data.</param>
        protected override void OnSpreadsheetEditorChanged(PropertyChangedEventArgs<WpfSpreadsheetEditorControl> args)
        {
            base.OnSpreadsheetEditorChanged(args);

            if (args.OldValue != null)
            {
                SpreadsheetVisualEditor visualEditor = args.OldValue.VisualEditor;
                visualEditor.DocumentSourceChanged -= VisualEditor_DocumentSourceChanged;
                visualEditor.EditorChanged -= VisualEditor_EditorChanged;
                visualEditor.DocumentSavingStarted -= VisualEditor_DocumentSavingStarted;
                visualEditor.DocumentSavingFinished -= VisualEditor_DocumentSavingFinished;
            }

            if (args.NewValue != null)
            {
                SpreadsheetVisualEditor visualEditor = args.NewValue.VisualEditor;
                visualEditor.DocumentSourceChanged += VisualEditor_DocumentSourceChanged;
                visualEditor.EditorChanged += VisualEditor_EditorChanged;
                visualEditor.DocumentSavingStarted += VisualEditor_DocumentSavingStarted;
                visualEditor.DocumentSavingFinished += VisualEditor_DocumentSavingFinished;
            }

            UpdateUI();
        }




        /// <summary>
        /// Raises the <see cref="FilenameChanged" /> event.
        /// </summary>
        /// <param name="args">The <see cref="EventArgs"/> instance containing the event data.</param>
        protected virtual void OnFilenameChanged(EventArgs args)
        {
            if (FilenameChanged != null)
                FilenameChanged(this, args);
        }

        #endregion


        #region PRIVATE

        #region UI

        /// <summary>
        /// "New" button is clicked.
        /// </summary>
        private void newButton_Click(object sender, RoutedEventArgs e)
        {
            NewDocument();
        }

        /// <summary>
        /// "Open" button is clicked.
        /// </summary>
        private void openButton_Click(object sender, RoutedEventArgs e)
        {
            OpenDocument();
        }

        /// <summary>
        /// "Info" button is clicked.
        /// </summary>
        private void infoButton_Click(object sender, RoutedEventArgs e)
        {
            DocumentInfoWindow.ShowDialog(VisualEditor);
        }

        /// <summary>
        /// "Save" button is clicked.
        /// </summary>
        private void saveButton_Click(object sender, RoutedEventArgs e)
        {
            SaveDocumentChanges();
        }

        /// <summary>
        /// "Save As" button is clicked.
        /// </summary>
        private void saveAsButton_Click(object sender, RoutedEventArgs e)
        {
            SaveDocumentAs();
        }


        #region Export

        /// <summary>
        /// "Export" button is clicked.
        /// </summary>
        private void exportButton_Click(object sender, RoutedEventArgs e)
        {
            _exportFileDialog.FileName = Path.GetFileNameWithoutExtension(Filename);
            if (_exportFileDialog.ShowDialog() == true)
            {
                try
                {
                    string extension = Path.GetExtension(_exportFileDialog.FileName);

                    if (string.Equals(extension, ".tsv", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(extension, ".csv", StringComparison.OrdinalIgnoreCase))
                    {
                        if (MessageBox.Show(
                            "The selected file type does not support workbooks that contain multiple sheets.\r\nTo save only the active sheet, click OK.\r\nTo save all sheets, save them individually using a different file name for each, or choose a file type that supports multiple sheets.",
                            "Export document", MessageBoxButton.OKCancel, MessageBoxImage.Warning) == MessageBoxResult.OK)
                        {
                            // create a temporary stream
                            using (MemoryStream tempStream = new MemoryStream())
                            {
                                // save XLSX file to a temporary stream
                                VisualEditor.SaveDocumentTo(tempStream);

                                tempStream.Position = 0;

                                using (Stream stream = File.Create(_exportFileDialog.FileName))
                                {
                                    DocumentEnvironmentProperties environmentProperties = DocumentEnvironmentProperties.Default;
                                    environmentProperties.Culture = CultureInfo.CurrentCulture;
                                    environmentProperties.UICulture = CultureInfo.CurrentUICulture;

                                    if (string.Equals(extension, ".tsv", StringComparison.OrdinalIgnoreCase))
                                        OpenXmlDocumentConverter.ConvertXlsxToTsv(environmentProperties, tempStream, VisualEditor.FocusedWorksheetIndex, stream);
                                    else
                                        OpenXmlDocumentConverter.ConvertXlsxToCsv(environmentProperties, tempStream, VisualEditor.FocusedWorksheetIndex, stream, System.Text.Encoding.UTF8);
                                }
                            }
                        }
                    }
                    else
                    {
                        if (!_isLayoutSettingsInitialized)
                        {
                            // set layout settings
                            if (_layoutSettingsManager.EditLayoutSettingsUseDialog(Application.Current.MainWindow))
                                _isLayoutSettingsInitialized = true;
                            else
                                return;
                        }

                        // create a temporary stream
                        using (MemoryStream tempStream = new MemoryStream())
                        {
                            // save XLSX file to a temporary stream
                            VisualEditor.SaveDocumentTo(tempStream);

                            // add XLSX file to the image collection of document converter
                            _converter.Images.Add(tempStream);

                            // create dialog that displays progress for document conversion process
                            ActionProgressWindow dlg = new ActionProgressWindow(ExportDocument, 1, "Export document");
                            // specify that dialog should be closed when conversion is finished
                            dlg.CloseAfterComplete = true;
                            // show dialog and run conversion process
                            dlg.RunAndShowDialog(Application.Current.MainWindow);

                            // clear image collection of document converter
                            _converter.Images.ClearAndDisposeItems();
                        }
                    }
                }
                catch (Exception ex)
                {
                    DemosTools.ShowErrorMessage(ex);
                }
            }
        }

        /// <summary>
        /// Exports the XLSX document.
        /// </summary>
        /// <param name="progressController">Progress controller.</param>
        private void ExportDocument(Vintasoft.Imaging.Utils.IActionProgressController progressController)
        {
            // set progress controller for document converter
            _converter.ProgressController = progressController;

            // convert XLSX to the selected format
            _converter.Convert(_exportFileDialog.FileName);
        }

        #endregion


        #region Print

        /// <summary>
        /// "Print" button is clicked.
        /// </summary>
        private void printButton_ButtonClick(object sender, RoutedEventArgs e)
        {
            PrintDocument();
        }


        /// <summary>
        /// "Layout settings" menu is selected.
        /// </summary>
        private void layoutSettingsMenuItem_Click(object sender, RoutedEventArgs e)
        {
            if (_layoutSettingsManager.EditLayoutSettingsUseDialog(Application.Current.MainWindow))
                _isLayoutSettingsInitialized = true;
        }

        /// <summary>
        /// "Page settings" menu is selected.
        /// </summary>
        private void pageSettingsMenuItem_Click(object sender, RoutedEventArgs e)
        {
            // show dialog with page setup setings

            PageSettingsWindow pageSettingsWindow = new PageSettingsWindow(
                _imagePrintManager, _imagePrintManager.PagePadding, _imagePrintManager.ImagePadding);
            pageSettingsWindow.WindowStartupLocation = WindowStartupLocation.CenterOwner;
            pageSettingsWindow.Owner = Application.Current.MainWindow;
            if (pageSettingsWindow.ShowDialog() == true)
            {
                _imagePrintManager.PagePadding = pageSettingsWindow.PagePadding;
                _imagePrintManager.ImagePadding = pageSettingsWindow.ImagePadding;
            }
        }

        /// <summary>
        /// "Print preview" menu is selected.
        /// </summary>
        private void printPreviewMenuItem_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // if layout settings are not initialized
                if (!_isLayoutSettingsInitialized)
                {
                    // set layout settings
                    if (_layoutSettingsManager.EditLayoutSettingsUseDialog(Application.Current.MainWindow))
                        _isLayoutSettingsInitialized = true;
                    else
                        return;
                }

                // create a temporary stream
                using (MemoryStream tempStream = new MemoryStream())
                {
                    // save XLSX file to a temporary stream
                    VisualEditor.SaveDocumentTo(tempStream);

                    // create a dialog that allows to preview and print XLSX document
                    PrintPreviewWindow window = new PrintPreviewWindow(tempStream);
                    window.WindowStartupLocation = WindowStartupLocation.CenterOwner;
                    window.Owner = Application.Current.MainWindow;
                    // show the dialog
                    window.ShowDialog();
                }
            }
            catch (Exception ex)
            {
                // show error message
                DemosTools.ShowErrorMessage(ex);
            }
        }

        #endregion


        /// <summary>
        /// "Close" button is clicked.
        /// </summary>
        private void closeButton_Click(object sender, RoutedEventArgs e)
        {
            CloseDocument(true);
        }

        #endregion


        /// <summary>
        /// Checks changes in document and saves document if necessary.
        /// </summary>
        public bool CheckChanges()
        {
            if (VisualEditor.Document == null)
                return true;

            VisualEditor.FinishEditCellValue();

            if (VisualEditor.Editor.IsVirtual)
                return true;

            if (VisualEditor.IsDocumentSourceChanged)
            {
                MessageBoxResult result = MessageBox.Show("Current workbook is changed. Do you want to save changes?", "New Workbook", MessageBoxButton.YesNoCancel, MessageBoxImage.Question, MessageBoxResult.Cancel);
                if (result == MessageBoxResult.Cancel)
                    return false;
                if (result == MessageBoxResult.Yes)
                {
                    try
                    {
                        if (_isNewDocument)
                        {
                            if (!SaveAs())
                                return false;
                        }
                        else
                        {
                            VisualEditor.SaveDocumentChanges();
                        }
                    }
                    catch (Exception ex)
                    {
                        DemosTools.ShowErrorMessage(ex);
                        return false;
                    }
                }
            }
            return true;
        }

        /// <summary>
        /// "Options" button is clicked.
        /// </summary>
        private void optionsButton_Click(object sender, RoutedEventArgs e)
        {
            OptionsWindow.ShowDialog(VisualEditor);
        }

        /// <summary>
        /// Handles the EditorChanged event of the <see cref="SpreadsheetVisualEditorPanel.VisualEditor"/>.
        /// </summary>
        private void VisualEditor_EditorChanged(object sender, PropertyChangedEventArgs<Vintasoft.Imaging.Office.Spreadsheet.SpreadsheetEditor> e)
        {
            UpdateUI();
        }

        /// <summary>
        /// Handles the DocumentSourceChanged event of the <see cref="SpreadsheetVisualEditorPanel.VisualEditor"/>.
        /// </summary>
        private void VisualEditor_DocumentSourceChanged(object sender, EventArgs e)
        {
            saveButton.IsEnabled = !_isNewDocument;
        }

        /// <summary>
        /// Updates the User Interface.
        /// </summary>
        private void UpdateUI()
        {
            if (VisualEditor.Document == null)
            {
                infoButton.IsEnabled = false;
                saveButton.IsEnabled = false;
                saveAsButton.IsEnabled = false;
                exportButton.IsEnabled = false;
                printSplitButton.IsEnabled = false;
                closeButton.IsEnabled = false;
                SpreadsheetEditor.Visibility = Visibility.Hidden;
            }
            else
            {
                SpreadsheetEditor.Visibility = Visibility.Visible;
                infoButton.IsEnabled = true;
                saveButton.IsEnabled = VisualEditor.IsDocumentSourceChanged;
                saveAsButton.IsEnabled = true;
                exportButton.IsEnabled = true;
                printSplitButton.IsEnabled = true;
                closeButton.IsEnabled = true;
            }
        }

        /// <summary>
        /// Saves document as.
        /// </summary>
        private bool SaveAs()
        {
            _saveWorksheetFileDialog.FileName = Filename;
            if (_saveWorksheetFileDialog.ShowDialog() == true)
            {
                VisualEditor.SaveDocumentAs(_saveWorksheetFileDialog.FileName);
                SetFilename(_saveWorksheetFileDialog.FileName);
                _isNewDocument = false;
                return true;
            }
            return false;
        }

        /// <summary>
        /// Sets the filename.
        /// </summary>
        /// <param name="filename">The filename.</param>
        private void SetFilename(string filename)
        {
            if (_filename != filename)
            {
                _filename = filename;
                OnFilenameChanged(EventArgs.Empty);
            }
        }

        /// <summary>
        /// Handles the DocumentSavingStarted event of the VisualEditor.
        /// </summary>
        private void VisualEditor_DocumentSavingStarted(object sender, EventArgs e)
        {
            SpreadsheetEditor editor = VisualEditor.Editor;
            DocumentInformation info = new DocumentInformation(editor.DocumentInformation);

            // set the user name that last modified this document
            info.LastModifiedBy = Environment.UserName;

            // set the modified date
            info.ModifiedDate = DateTime.Now.ToString(CultureInfo.InvariantCulture);

            // set DocumentInformation
            editor.DocumentInformation = info;
        }


        /// <summary>
        /// Handles the DocumentSavingFinished event of the VisualEditor.
        /// </summary>
        private void VisualEditor_DocumentSavingFinished(object sender, EventArgs e)
        {
            saveButton.IsEnabled = false;
        }


        #endregion

        #endregion



        #region Events

        /// <summary>
        /// Occurs when <see cref="Filename"/> is changed.
        /// </summary>
        public event EventHandler FilenameChanged;

        #endregion

    }
}

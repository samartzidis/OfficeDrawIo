using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;

namespace OfficeDrawIo
{
    public partial class ThisAddIn
    {
        public SynchronizationContext TheWindowsFormsSynchronizationContext { get; private set; }

        private Microsoft.Office.Tools.Word.PictureContentControl _selectedCtrl;
        private static ThisAddIn _addin;    
        private string _userTmpFilesDir;
        private SettingsAdapter _settings;
        private SettingsForm _sf;
        private string _drawioExportDir;
        private Ribbon _ribbon;
        private FileSystemWatcher _watcher;
        private object _addInNotifyChangedLock = new object();

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            TheWindowsFormsSynchronizationContext = SynchronizationContext.Current ?? new WindowsFormsSynchronizationContext();

            _addin = this;
            _drawioExportDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "drawio-export");
            _userTmpFilesDir = Path.Combine(Path.GetTempPath(), "OfficeDrawIo");
            _settings = new SettingsAdapter();
            _sf = new SettingsForm(_settings, () =>
            {
                Properties.Settings.Default.Reset();
            });

            Application.DocumentBeforeSave += Application_DocumentBeforeSave;
            Application.DocumentOpen += Application_DocumentOpen;
            Application.DocumentBeforeClose += Application_DocumentBeforeClose;
            
            if (!Directory.Exists(_userTmpFilesDir))
                Directory.CreateDirectory(_userTmpFilesDir);

            CreateFileWatcher(_userTmpFilesDir);
        }

        public void SetRibbon(Ribbon ribbon)
        {
            _ribbon = ribbon;            
        }

        public void AddDrawIoDiagramOnDocument()
        {
            if (!ValidateDependencies())
                return;

            var blankDrawIoXml = Helpers.LoadStringResource("Resources.blank.drawio");
            var part = AddDrawIoDataPart(blankDrawIoXml);

            var ctrl = ActiveVstoDocument.Controls.AddPictureContentControl(part.Id);
            //pictureControl.Title = pictureControl1XMLPartID;
            ctrl.Title = $"Draw.io Diagram";
            ctrl.Image = DrawFilledRectangle(128, 128);
            ctrl.Tag = MakeDrawioTag(part.Id);

            ctrl.LockContents = true;

            ctrl.Entering += PictureControl_Entering;
            ctrl.Exiting += PictureControl_Exiting;
            ctrl.Deleting += PictureControl_Deleting;
            
        }
        public void EditDrawIoDiagramOnDocument()
        {
            if (!ValidateDependencies())
                return;

            if (_selectedCtrl == null)
                return;
          
            var id = GetDrawioTagGuidPart(_selectedCtrl.Tag);
            var wnd = NativeWindowHelper.FindWindowsWithText(id).FirstOrDefault();
            if (wnd != IntPtr.Zero)
            {
                NativeWindowHelper.RestoreFromMinimized(wnd);
                NativeWindowHelper.SetForegroundWindow(wnd);
                return;
            }

            var drawioFilePath = Path.Combine(_userTmpFilesDir, $"{id}.drawio");
            var drawioData = GetDrawIoDataPart(id);
            if (drawioData == null)
                return;            
            File.WriteAllText(drawioFilePath, drawioData);

            var process = new Process();
            process.StartInfo.FileName = _settings.DrawIoExePath;
            process.StartInfo.Arguments = drawioFilePath;
            process.StartInfo.WindowStyle = ProcessWindowStyle.Maximized;

            try
            {
                process.Start();
            }
            catch (Exception m)
            {
                MessageBox.Show($"Failed to start Draw.io Desktop application: {m.Message}.",
                    Application.ActiveWindow.Caption, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }
        public void ExportDrawIoDiagram()
        {
            if (!ValidateDependencies())
                return;

            if (_selectedCtrl == null)
                return;

            var dlg = new SaveFileDialog();
            dlg.Filter = "Draw.io files (*.drawio)|*.drawio|All files (*.*)|*.*";
            dlg.DefaultExt = ".drawio";

            if (dlg.ShowDialog() == DialogResult.OK)
            {
                var data = GetDrawIoDataPart(GetDrawioTagGuidPart(_selectedCtrl.Tag));
                File.WriteAllText(dlg.FileName, data);
            }
        }
        public void Settings()
        {
            if (_sf.ShowDialog() == DialogResult.OK)
                _settings.Save();
        }
        public void AddInNotifyChanged(string partId)
        {
            bool _lockTaken = false;            
            try
            {
                Monitor.TryEnter(_addInNotifyChangedLock, ref _lockTaken);
                if (_lockTaken)
                {
                    if (!File.Exists(_settings.NodeJsExePath))
                        return;

                    if (!ExistsDrawIoDataPart(partId))
                        return;

                    var drawioFilePath = Path.Combine(_userTmpFilesDir, $"{partId}.drawio");
                    var pngFilePath = Path.Combine(_userTmpFilesDir, $"{partId}.png");

                    string drawioData = null;
                    using (var fs = new FileStream(drawioFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    using (var sr = new StreamReader(fs, Encoding.Default))
                        drawioData = sr.ReadToEnd();

                    //Check if the XMLPart payload has really changed
                    var currentXmlPartData = GetDrawIoDataPart(partId);
                    if (currentXmlPartData == drawioData)
                        return; // No need to change XMLPart

                    using (new ScopedCursor(Cursors.WaitCursor))
                    {
                        if (!UpdateDrawIoDataPart(partId, drawioData))
                            return;

                        using (var process = new Process())
                        {
                            string stdErrData = string.Empty;

                            process.StartInfo.FileName = _settings.NodeJsExePath;
                            process.StartInfo.WorkingDirectory = _drawioExportDir;
                            process.StartInfo.Arguments = $"index.js \"{drawioFilePath}\" \"{pngFilePath}\"";
                            process.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                            process.StartInfo.RedirectStandardError = true;
                            process.StartInfo.UseShellExecute = false;
                            process.StartInfo.CreateNoWindow = true;
                            process.ErrorDataReceived += (o, e) => { stdErrData += e.Data; };

                            try
                            {
                                var res = process.Start();
                                if (res == false)
                                    throw new ApplicationException("process failed to start");

                                process.BeginErrorReadLine();

                                res = process.WaitForExit(10 * 1000);
                                if (res == false)
                                    throw new ApplicationException("process timed out");
                            }
                            catch (Exception m)
                            {
                                MessageBox.Show($"Node.js: {m.Message}. Please check the Add-In settings.",
                                    Application.ActiveWindow.Caption, MessageBoxButtons.OK, MessageBoxIcon.Error);
                                return;
                            }

                            if (process.ExitCode != 0 || !string.IsNullOrEmpty(stdErrData))
                            {
                                MessageBox.Show($"Node.js: rendering failed (node.js exit code: {process.ExitCode}).\r\n{stdErrData}",
                                    Application.ActiveWindow.Caption, MessageBoxButtons.OK, MessageBoxIcon.Error);
                                return;
                            }
                        }

                        if (File.Exists(pngFilePath))
                        {
                            using (var bitmap1 = new System.Drawing.Bitmap(pngFilePath, true))
                            {
                                foreach (var cc in ActiveVstoDocument.Controls)
                                {
                                    if (cc is Microsoft.Office.Tools.Word.PictureContentControl ctrl && GetDrawioTagGuidPart(ctrl.Tag) == partId)
                                    {
                                        ctrl.LockContents = false;
                                        ctrl.Image = bitmap1;
                                        ctrl.LockContents = true;

                                        break;
                                    }
                                }
                            }
                        }
                    }
                }
            }
            finally
            {
                if (_lockTaken)
                    Monitor.Exit(_addInNotifyChangedLock);
            }
        }
        public void About()
        {
            var dlg = new AboutBox();
            dlg.ShowDialog();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            _watcher?.Dispose();
        }

        private void Application_DocumentOpen(Microsoft.Office.Interop.Word.Document doc)
        {
            // See: https://docs.microsoft.com/en-us/visualstudio/vsto/persisting-dynamic-controls-in-office-documents?view=vs-2017
            foreach (Microsoft.Office.Interop.Word.ContentControl nativeControl in doc.ContentControls)
            {
                if (nativeControl.Type == Microsoft.Office.Interop.Word.WdContentControlType.wdContentControlPicture)
                {
                    if (!IsValidDrawioTag(nativeControl.Tag))
                        continue;

                    var ctrl = ActiveVstoDocument.Controls.AddPictureContentControl(nativeControl, nativeControl.Tag);

                    ctrl.LockContents = true;

                    ctrl.Entering += PictureControl_Entering;
                    ctrl.Exiting += PictureControl_Exiting;
                    ctrl.Deleting += PictureControl_Deleting;

                }
            }
            
        }

        private bool IsValidDrawioTag(string tag)
        {
            return tag != null && tag.StartsWith("OfficeDrawIo:");
        }

        private string GetDrawioTagGuidPart(string tag)
        {
            if (!IsValidDrawioTag(tag))
                return null;
            return tag.Split(':')[1];
        }

        private string MakeDrawioTag(string id)
        {
            return $"OfficeDrawIo:{id}";
        }

        private void Application_DocumentBeforeClose(Microsoft.Office.Interop.Word.Document doc, ref bool cancel)
        {
            foreach (var cc in ActiveVstoDocument.Controls)
            {
                if (cc is Microsoft.Office.Tools.Word.PictureContentControl ctrl)
                {
                    var partId = GetDrawioTagGuidPart(ctrl.Tag);
                    if (partId == null)
                        continue;

                    try
                    {
                        var drawioFilePath = Path.Combine(_userTmpFilesDir, $"{partId}.drawio");
                        if (File.Exists(drawioFilePath))
                            File.Delete(drawioFilePath);
                    }
                    catch { }

                    try
                    {
                        var imageFilePath = Path.Combine(_userTmpFilesDir, $"{partId}.png");
                        if (File.Exists(imageFilePath))
                            File.Delete(imageFilePath);
                    }
                    catch { }
                }
            }
        }

        private Microsoft.Office.Tools.Word.Document ActiveVstoDocument
        {
            get
            {
                if (Application.ActiveDocument == null)
                    return null;

                return Globals.Factory.GetVstoObject(Application.ActiveDocument);
            }
        }        

        private Bitmap DrawFilledRectangle(int x, int y)
        {
            var bmp = new Bitmap(x, y);
            using (Graphics graph = Graphics.FromImage(bmp))
            {
                var imageSize = new Rectangle(0, 0, x, y);
                graph.FillRectangle(Brushes.LightBlue, imageSize);
            }
            return bmp;
        }

        private void PictureControl_Exiting(object sender, Microsoft.Office.Tools.Word.ContentControlExitingEventArgs e)
        {
            _selectedCtrl = null;

            _ribbon.btnEditDiagram.Enabled = false;
            _ribbon.btnExport.Enabled = false;
            _ribbon.btnAddDiagram.Enabled = true;

        }

        private void PictureControl_Entering(object sender, Microsoft.Office.Tools.Word.ContentControlEnteringEventArgs e)
        {
            _selectedCtrl = sender as Microsoft.Office.Tools.Word.PictureContentControl;

            _ribbon.btnEditDiagram.Enabled = true;
            _ribbon.btnExport.Enabled = true;
            _ribbon.btnAddDiagram.Enabled = false;
        }

        private void PictureControl_Deleting(object sender, Microsoft.Office.Tools.Word.ContentControlDeletingEventArgs e)
        {
            // Commented out to support undo operations. Deleting in Application_DocumentBeforeSave instead.
            //var ctrl = (Microsoft.Office.Tools.Word.PictureContentControl)sender;
            //DeleteDrawIoDataPart(GetDrawioTagGuidPart(ctrl.Tag));
        }

        private bool UpdateDrawIoDataPart(string id, string data)
        {
            var xmlPart = Application.ActiveDocument.CustomXMLParts.SelectByID(id);
            if (xmlPart == null)
                return false;

            var base64 = System.Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(data));
            xmlPart.DocumentElement.FirstChild.NodeValue = base64;

            return true;
        }

        private bool ExistsDrawIoDataPart(string id)
        {
            var xmlPart = Application.ActiveDocument.CustomXMLParts.SelectByID(id);
            return xmlPart != null;
        }

        private Microsoft.Office.Core.CustomXMLPart AddDrawIoDataPart(string data)
        {
            var base64 = System.Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(data));
            var xmlPart = Application.ActiveDocument.CustomXMLParts.Add($"<OfficeDrawIo v=\"1\">{base64}</OfficeDrawIo>");
            return xmlPart;
        }

        private string GetDrawIoDataPart(string id)
        {
            var xmlPart = Application.ActiveDocument.CustomXMLParts.SelectByID(id);
            if (xmlPart == null)
                return null;

            var xdoc = new XmlDocument();
            xdoc.LoadXml(xmlPart.XML);

            var base64EncodedBytes = System.Convert.FromBase64String(xdoc.InnerText);
            return System.Text.Encoding.UTF8.GetString(base64EncodedBytes);
        }

        private void DeleteDrawIoDataPart(string id)
        {
            foreach (Microsoft.Office.Core.CustomXMLPart part in ActiveVstoDocument.CustomXMLParts)
            {
                if (part.Id == id)
                {
                    part.Delete();
                    break;
                }
            }
        }  

        private FileSystemWatcher CreateFileWatcher(string path)
        {
            if (_watcher != null)
                return _watcher;

            _watcher = new FileSystemWatcher(path);
            _watcher.NotifyFilter = NotifyFilters.LastAccess | NotifyFilters.LastWrite | NotifyFilters.FileName | NotifyFilters.DirectoryName;
            _watcher.Filter = "*.drawio";
            _watcher.IncludeSubdirectories = false;
            _watcher.Changed += OnFileChanged;
            //watcher.Created += OnFileChanged;
            //watcher.Deleted += OnFileChanged;
            _watcher.EnableRaisingEvents = true;

            return _watcher;
        }        

        private static void OnFileChanged(object source, FileSystemEventArgs e)
        {
            var xmlPartId = Path.GetFileNameWithoutExtension(e.FullPath);

            Globals.ThisAddIn.TheWindowsFormsSynchronizationContext.Send(d =>
            {
                _addin.AddInNotifyChanged(xmlPartId);
            }, null);

        }  

        private void Application_DocumentBeforeSave(Microsoft.Office.Interop.Word.Document doc, ref bool saveAsUi, ref bool cancel)
        {
            // Get list of all Draw.io PictureContentControls in document
            var ids = new HashSet<string>();
            foreach (Microsoft.Office.Interop.Word.ContentControl nativeControl in doc.ContentControls)
            {
                if (nativeControl.Type == Microsoft.Office.Interop.Word.WdContentControlType.wdContentControlPicture 
                    && IsValidDrawioTag(nativeControl.Tag))
                    ids.Add(GetDrawioTagGuidPart(nativeControl.Tag));
            }

            // Clean up unreferenced OfficeDrawIo data parts
            foreach (Microsoft.Office.Core.CustomXMLPart part in Application.ActiveDocument.CustomXMLParts)
            {
                if(!ids.Contains(part.Id) && part.XML != null && part.XML.TrimStart().StartsWith("<OfficeDrawIo"))
                    part.Delete();
            }
        }

        private bool ValidateDependencies()
        {
            if (!File.Exists(_settings.DrawIoExePath))
            {
                MessageBox.Show($"Draw.io Desktop not found. Please download and install it from {Properties.Settings.Default.DrawIoUrl}. If you believe it is installed, then please check the Add-In settings.",
                    Application.ActiveWindow.Caption, MessageBoxButtons.OK, MessageBoxIcon.Error);

                System.Diagnostics.Process.Start(Properties.Settings.Default.DrawIoUrl);

                return false;
            }

            if (!File.Exists(_settings.NodeJsExePath))
            {
                MessageBox.Show($"Node.js not found. Please download and install it from {Properties.Settings.Default.NodeJsUrl}. If you believe it is installed, then please check the Add-In settings.",
                    Application.ActiveWindow.Caption, MessageBoxButtons.OK, MessageBoxIcon.Error);

                System.Diagnostics.Process.Start(Properties.Settings.Default.NodeJsUrl);

                return false;
            }

            return true;
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += ThisAddIn_Startup;
            this.Shutdown += ThisAddIn_Shutdown;
        }
        
        #endregion
    }
}

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;
using OfficeDrawIo;
using OfficeDrawIoUtil;
using Word = Microsoft.Office.Interop.Word;

namespace OfficeDrawIoWord
{
    public partial class ThisAddIn
    {
        public EventHandler SelectionChanged;
        public SynchronizationContext TheWindowsFormsSynchronizationContext { get; private set; }
        public ShapeHolder SelectedShape => GetCurrentSelection().FirstOrDefault();
        public bool SuppressFileWatcherNotifications { get; private set; }

        private static ThisAddIn _addin;    
        private string _userTmpFilesDir;
        private SettingsAdapter _settings;
        private SettingsForm _sf;
        private FileSystemWatcher _watcher;
        private readonly BiDictionary<Guid, int> _editMap = new BiDictionary<Guid, int>();

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Trace.Listeners.Clear();
            Trace.Listeners.Add(new OfficeDrawIo.TraceListener());
            Trace.WriteLine("ThisAddIn_Startup()");

            _addin = this;
            _userTmpFilesDir = Path.Combine(Path.GetTempPath(), $"OfficeDrawIo.{Process.GetCurrentProcess().Id}");
            Trace.WriteLine(_userTmpFilesDir);
            _settings = new SettingsAdapter();
            _sf = new SettingsForm(_settings, () => Properties.Settings.Default.Reset());

            TheWindowsFormsSynchronizationContext = SynchronizationContext.Current ?? new WindowsFormsSynchronizationContext();            

            if (!Directory.Exists(_userTmpFilesDir))
                Directory.CreateDirectory(_userTmpFilesDir);

            Application.WindowBeforeDoubleClick += Application_WindowBeforeDoubleClick;
            Application.WindowSelectionChange += Application_WindowSelectionChange;

            EnsureStartFileWatcher();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            EnsureStopFileWatcher();

            try 
            {
                Trace.WriteLine($"ThisAddIn_Shutdown: Deleting {_userTmpFilesDir}");
                Directory.Delete(_userTmpFilesDir, true); 
            } catch { }
        }

        private void Application_WindowSelectionChange(Word.Selection sel)
        {
            var handler = SelectionChanged;
            handler?.Invoke(this, EventArgs.Empty);
        }

        private IEnumerable<ShapeHolder> GetCurrentSelection()
        {
            foreach (Word.InlineShape shape in ActiveDocument.InlineShapes)
            {
                if (shape.Range.InRange(Application.Selection.Range) 
                    && shape.Title.IndexOf(Util.OfficeDrawIoPayloadPrexix, StringComparison.Ordinal) == 0)
                {
                    yield return new ShapeHolder(shape);
                }
            }

            foreach (Word.Shape shape in Application.Selection.ShapeRange)
            {
                if(shape.Title.IndexOf(Util.OfficeDrawIoPayloadPrexix, StringComparison.Ordinal) == 0)
                    yield return new ShapeHolder(shape);
            }
        }

        private void Application_WindowBeforeDoubleClick(Word.Selection sel, ref bool cancel)
        {
            if (SelectedShape != null)
                EditDiagramShape(SelectedShape);

            cancel = true;
        }

        private ShapeHolder FindShape(Guid editId)
        {
            if (!_editMap.TryGetByFirst(editId, out var anchorId))
                return null;

            return FindShapeByAnchorId(anchorId);
        }

        private ShapeHolder FindShapeByAnchorId(int anchorId)
        {
            foreach (Word.InlineShape shape in ActiveDocument.InlineShapes)
                if (shape.AnchorID == anchorId)
                    return new ShapeHolder(shape);

            foreach (Word.Shape shape in ActiveDocument.Shapes)
                if (shape.AnchorID == anchorId)
                    return new ShapeHolder(shape);

            return null;
        }

        private void HousekeepEditMap()
        {
            foreach (var anchorId in _editMap.SecondKeys.ToList())
            {
                if (FindShapeByAnchorId(anchorId) == null)
                {
                    _editMap.TryRemoveBySecond(anchorId, out var editFileId);

                    var editFilePath = Path.Combine(_userTmpFilesDir, $"{editFileId}.png");
                    try 
                    {
                        Trace.WriteLine($"HousekeepEditMap: Deleting {editFilePath}");
                        File.Delete(editFilePath);
                    } catch { }
                }
            }
        }


        public void EditDiagramShape(ShapeHolder shape)
        {
            Trace.WriteLine($"EditDiagramShape: AnchorID: {shape.AnchorID}");
            if (shape.AnchorID == 0)
                return;

            if (!_editMap.TryGetBySecond(shape.AnchorID, out var editFileId))
            {
                editFileId = Guid.NewGuid();
                _editMap.AddOrUpdate(editFileId, shape.AnchorID);
            }

            var wnd = NativeWindowHelper.FindWindow($"{editFileId}.png - draw.io");
            if (wnd != IntPtr.Zero)
            {
                NativeWindowHelper.RestoreFromMinimized(wnd);
                NativeWindowHelper.SetForegroundWindow(wnd);

                return;
            }

            var editFilePath = Path.Combine(_userTmpFilesDir, $"{editFileId}.png");

            try
            {
                var pngBytes = Util.DecodePngFile(shape.Title);
                File.WriteAllBytes(editFilePath, pngBytes);

                var process = new Process();
                process.StartInfo.FileName = _settings.DrawIoExePath;
                process.StartInfo.Arguments = editFilePath;
                process.StartInfo.WindowStyle = ProcessWindowStyle.Maximized;

                process.Start();
            }
            catch (Exception m)
            {
                var msg = $"Failed to start Draw.io Desktop application for file {editFilePath}. Error: {m}.";
                MessageBox.Show(msg, Application.ActiveWindow.Caption, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void FileNotifyChanged(string id)
        {
            Trace.WriteLine($"FileNotifyChanged({id})");
            if (!Guid.TryParse(id, out var editId))
                return;

            try
            {
                var oldShape = FindShape(editId);
                if (oldShape == null)
                    return;

                Trace.WriteLine($"FileNotifyChanged:oldShape.AnchorID: {oldShape.AnchorID}");

                var editFilePath = Path.Combine(_userTmpFilesDir, $"{id}.png");
                ShapeHolder newShape;
                if (oldShape.IsInlineShape)
                {
                    var range = oldShape.InlineShape.Range;
                    oldShape.Delete();
                    newShape = AddDiagramInlineShape(editFilePath, range);
                }
                else
                {
                    var rect = new RectangleF(oldShape.Shape.Left, oldShape.Shape.Top, oldShape.Shape.Width, oldShape.Shape.Height);
                    var rotation = oldShape.Shape.Rotation;

                    oldShape.Delete();
                    newShape = AddDiagramShape(editFilePath, rect);
                    newShape.Shape.Rotation = rotation;
                }

                Trace.WriteLine($"FileNotifyChanged:newShape.AnchorID: {newShape.AnchorID}");

                _editMap.AddOrUpdate(editId, newShape.AnchorID);
            }
            catch (Exception m)
            {
                MessageBox.Show($"Something went wrong while updating the the Draw.io diagram.\r\nError details:\r\n{m}",
                            Application.ActiveWindow.Caption, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private ShapeHolder AddDiagramInlineShape(string newPngFilePath, Word.Range range)
        {
            if (range == null)
                range = ActiveDocument.Application.Selection.Range;
            
            var shape = ActiveDocument.InlineShapes.AddPicture(FileName: newPngFilePath, Range: range);
            shape.Title = Util.EncodePngFile(newPngFilePath);
            shape.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue;

            return new ShapeHolder(shape);
        }

        private ShapeHolder AddDiagramShape(string newPngFilePath, RectangleF rect)
        {
            var shape = ActiveDocument.Shapes.AddPicture(FileName: newPngFilePath, Left: rect.Left, Top: rect.Top);
            shape.Title = Util.EncodePngFile(newPngFilePath);
            shape.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue;

            return new ShapeHolder(shape);
        }

        public void AddDiagramOnDocument()
        {            
            if (!ValidateDependencies())
                return;

            HousekeepEditMap();

            try
            {
                var newPngData = Helpers.LoadBinaryResource(Assembly.GetExecutingAssembly(), "Resources.new.png");
                var newPngFileName = "new.png";
                var newPngFilePath = Path.Combine(_userTmpFilesDir, newPngFileName);
                File.WriteAllBytes(newPngFilePath, newPngData);

                var shape = AddDiagramInlineShape(newPngFilePath, null);

                try { File.Delete(newPngFilePath); } catch { }
            }
            catch (Exception m)
            {
                var msg = $"Failed to add Draw.io diagram. Error: {m}.";
                MessageBox.Show(msg, Application.ActiveWindow.Caption, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void ExportDrawIoDiagram()
        {
            Trace.WriteLine("ExportDrawIoDiagram()");

            var sels = GetCurrentSelection().ToList();
            if (sels.Count == 0)
                return; // Nothing selected

            var dlg = new FolderBrowserDialog();
            if (dlg.ShowDialog() != DialogResult.OK) 
                return;

            try
            {
                var index = 1;
                foreach (var shape in sels)
                {
                    var filePath = Path.Combine(dlg.SelectedPath, $"{index++}.drawio.png");
                    var pngBytes = Util.DecodePngFile(shape.Title);
                    File.WriteAllBytes(filePath, pngBytes);
                }
            }
            catch (Exception m)
            {
                MessageBox.Show($@"Failed to export Draw.io document: {m.Message}.",
                    Application.ActiveWindow.Caption, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void Settings()
        {
            if (_sf.ShowDialog() == DialogResult.OK)
                _settings.Save();
        }

        public void About()
        {
            var dlg = new AboutBox();
            dlg.ShowDialog();
        }


        private Microsoft.Office.Tools.Word.Document ActiveDocument
        {
            get
            {
                if (Application.ActiveDocument == null)
                    return null;
                try
                {
                    return Globals.Factory.GetVstoObject(Application.ActiveDocument);
                }
                catch
                {
                }

                return null;
            }
        }        

        private void EnsureStartFileWatcher()
        {
            if (_watcher != null)
                return;

            _watcher = new FileSystemWatcher(_userTmpFilesDir);
            _watcher.NotifyFilter = NotifyFilters.LastAccess | NotifyFilters.LastWrite | NotifyFilters.FileName | NotifyFilters.DirectoryName;
            _watcher.Filter = "*.png";
            _watcher.IncludeSubdirectories = false;

            _watcher.Changed += OnFileChanged;

            _watcher.EnableRaisingEvents = true;
        }

        private void EnsureStopFileWatcher()
        {
            _watcher?.Dispose();
            _watcher = null;
        }

        private static void OnFileChanged(object source, FileSystemEventArgs e)
        {
            if (_addin.SuppressFileWatcherNotifications)
                return;

            var id = Path.GetFileNameWithoutExtension(e.FullPath);

            Globals.ThisAddIn.TheWindowsFormsSynchronizationContext.Send(d =>
            {
                using (new ScopedCursor(Cursors.WaitCursor))
                using (new ScopedLambda(() => _addin.SuppressFileWatcherNotifications = true, () => _addin.SuppressFileWatcherNotifications = false))
                    _addin.FileNotifyChanged(id);

            }, null);
        }  

        private bool ValidateDependencies()
        {
            if (!File.Exists(_settings.DrawIoExePath))
            {
                MessageBox.Show($"Draw.io Desktop not found. Please download and install it from {Properties.Settings.Default.DrawIoUrl}. If you believe it is installed, then please check the Add-In settings.",
                    Application.ActiveWindow.Caption, MessageBoxButtons.OK, MessageBoxIcon.Error);

                Process.Start(Properties.Settings.Default.DrawIoUrl);

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

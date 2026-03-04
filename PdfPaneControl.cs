using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

using Outlook = Microsoft.Office.Interop.Outlook;

// iText7
using iText.Kernel.Pdf;

namespace OutlookAddIn1
{
    public partial class PdfPaneControl : UserControl
    {
        private readonly Outlook.Application _app;

        private FlowLayoutPanel pdfListPanel;
        private FlowLayoutPanel embListPanel;

        private Panel header;
        private Label lblTitle;
        private Button btnClose;

        private Label lblPdfsTitle;
        private Label lblEmbTitle;

        private ToolTip tip;

        private List<PdfAttachmentItem> _pdfs = new List<PdfAttachmentItem>();
        private Dictionary<string, byte[]> _embedded = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);

        private string _currentPdfPath;

        private CancellationTokenSource _cts;
        private System.Windows.Forms.Timer _debounceTimer;

        public PdfPaneControl(Outlook.Application app)
        {
            _app = app;

            try { InitializeComponent(); } catch { }

            BuildUi();
            SetupDebounce();
        }

        private static string TempDir
        {
            get { return Path.Combine(Path.GetTempPath(), "PdfEmbeddedViewer"); }
        }

        private void SetupDebounce()
        {
            _debounceTimer = new System.Windows.Forms.Timer();
            _debounceTimer.Interval = 250;
            _debounceTimer.Tick += async (s, e) =>
            {
                _debounceTimer.Stop();
                await RefreshFromContextAsync();
            };
        }

        public void TriggerContextRefresh()
        {
            if (_debounceTimer == null) return;
            _debounceTimer.Stop();
            _debounceTimer.Start();
        }

        private void CancelWork()
        {
            try
            {
                if (_cts != null)
                {
                    _cts.Cancel();
                    _cts.Dispose();
                }
            }
            catch { }

            _cts = new CancellationTokenSource();
        }

        private void BuildUi()
        {
            this.Dock = DockStyle.Fill;
            this.BackColor = Color.White;
            tip = new ToolTip();

            // Header
            header = new Panel
            {
                Dock = DockStyle.Top,
                Height = 54,
                BackColor = Color.White
            };
            header.Paint += (s, e) =>
            {
                using (var p = new Pen(Color.FromArgb(230, 230, 230), 1))
                    e.Graphics.DrawLine(p, 0, header.Height - 1, header.Width, header.Height - 1);
            };

            lblTitle = new Label
            {
                Text = "PDF Embedded Viewer",
                Left = 14,
                Top = 16,
                Width = 340,
                Height = 22,
                Font = new Font("Segoe UI", 12f, FontStyle.Bold),
                ForeColor = Color.FromArgb(20, 20, 20)
            };

            btnClose = new Button
            {
                Text = "✕",
                Width = 36,
                Height = 28,
                Top = 12,
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.White
            };
            btnClose.FlatAppearance.BorderSize = 0;
            btnClose.ForeColor = Color.FromArgb(80, 80, 80);
            btnClose.Click += (s, e) => this.Visible = false;

            header.Controls.Add(lblTitle);
            header.Controls.Add(btnClose);

            this.Resize += (s, e) => btnClose.Left = this.Width - btnClose.Width - 10;
            btnClose.Left = this.Width - btnClose.Width - 10;

            // One scroll container so nothing is "far apart"
            var content = new Panel
            {
                Dock = DockStyle.Fill,
                AutoScroll = true,
                BackColor = Color.White,
                Padding = new Padding(12, 10, 12, 12)
            };

            // PDFs section
            var pdfSection = new Panel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                BackColor = Color.White
            };

            var pdfHeader = new Panel { Dock = DockStyle.Top, Height = 28, BackColor = Color.White };

            lblPdfsTitle = new Label
            {
                Text = "Add documents",
                AutoSize = true,
                Left = 0,
                Top = 4,
                Font = new Font("Segoe UI", 11f, FontStyle.Regular),
                ForeColor = Color.FromArgb(30, 30, 30)
            };
            pdfHeader.Controls.Add(lblPdfsTitle);

            pdfListPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                BackColor = Color.White,
                Margin = new Padding(0),
                Padding = new Padding(0)
            };

            pdfSection.Controls.Add(pdfListPanel);
            pdfSection.Controls.Add(pdfHeader);

            // small spacer
            var spacer = new Panel { Dock = DockStyle.Top, Height = 12, BackColor = Color.White };

            // Embedded section
            var embSection = new Panel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                BackColor = Color.White
            };

            var embHeader = new Panel { Dock = DockStyle.Top, Height = 28, BackColor = Color.White };

            lblEmbTitle = new Label
            {
                Text = "Embedded",
                AutoSize = true,
                Left = 0,
                Top = 4,
                Font = new Font("Segoe UI", 11f, FontStyle.Regular),
                ForeColor = Color.FromArgb(30, 30, 30)
            };
            embHeader.Controls.Add(lblEmbTitle);

            embListPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                BackColor = Color.White,
                Margin = new Padding(0),
                Padding = new Padding(0)
            };

            embSection.Controls.Add(embListPanel);
            embSection.Controls.Add(embHeader);

            // order
            content.Controls.Add(embSection);
            content.Controls.Add(spacer);
            content.Controls.Add(pdfSection);

            this.Controls.Clear();
            this.Controls.Add(content);
            this.Controls.Add(header);

            ClearUiOnly();
        }

        private void ClearUiOnly()
        {
            try { pdfListPanel.Controls.Clear(); } catch { }
            try { embListPanel.Controls.Clear(); } catch { }

            _pdfs.Clear();
            _embedded.Clear();
            _currentPdfPath = null;
        }

        private Outlook.MailItem GetCurrentMailItem()
        {
            try
            {
                var explorer = _app.ActiveExplorer();
                if (explorer != null && explorer.Selection != null && explorer.Selection.Count > 0)
                {
                    object item = explorer.Selection[1];
                    var mail = item as Outlook.MailItem;
                    if (mail != null) return mail;
                }

                var insp = _app.ActiveInspector();
                if (insp != null)
                {
                    var mail2 = insp.CurrentItem as Outlook.MailItem;
                    if (mail2 != null) return mail2;
                }
            }
            catch { }

            return null;
        }

        private async Task RefreshFromContextAsync()
        {
            CancelWork();
            var token = _cts.Token;

            Outlook.MailItem mail = null;
            try { mail = GetCurrentMailItem(); } catch { }

            ClearUiOnly();

            if (mail == null) return;

            List<PdfAttachmentItem> found = new List<PdfAttachmentItem>();

            try
            {
                var atts = mail.Attachments;
                if (atts == null || atts.Count == 0)
                    return;

                for (int i = 1; i <= atts.Count; i++)
                {
                    if (token.IsCancellationRequested) return;

                    Outlook.Attachment a = null;
                    try { a = atts[i]; } catch { continue; }

                    string name = "";
                    try { name = a.FileName ?? ""; } catch { name = ""; }

                    if (name.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase))
                        found.Add(new PdfAttachmentItem { AttachmentIndex = i, Name = name });
                }
            }
            catch
            {
                return; // never crash Outlook
            }

            _pdfs = found;

            foreach (var pdf in _pdfs)
            {
                if (token.IsCancellationRequested) return;

                var row = CreateLinkRow(
                    pdf.Name,
                    async () =>
                    {
                        await OpenPdfAndReadEmbeddedAsync(pdf);

                        // open the PDF itself in Edge/default
                        try
                        {
                            if (!string.IsNullOrWhiteSpace(_currentPdfPath) && File.Exists(_currentPdfPath))
                                Process.Start(new ProcessStartInfo { FileName = _currentPdfPath, UseShellExecute = true });
                        }
                        catch { }
                    },
                    () => DownloadMailAttachment(pdf)
                );

                pdfListPanel.Controls.Add(row);
            }

            // if no PDFs -> done (no crash)
            if (_pdfs.Count == 0) return;

            // auto-load first PDF to fill embedded list
            try { await OpenPdfAndReadEmbeddedAsync(_pdfs[0]); } catch { }
        }

        private void DownloadMailAttachment(PdfAttachmentItem pdf)
        {
            try
            {
                var mail = GetCurrentMailItem();
                if (mail == null) return;
                if (pdf == null || pdf.AttachmentIndex <= 0) return;

                using (var sfd = new SaveFileDialog
                {
                    FileName = pdf.Name,
                    Filter = "PDF (*.pdf)|*.pdf|Alle Dateien (*.*)|*.*"
                })
                {
                    if (sfd.ShowDialog() != DialogResult.OK) return;

                    var att = mail.Attachments[pdf.AttachmentIndex];
                    att.SaveAsFile(sfd.FileName);
                }
            }
            catch { }
        }

        private async Task OpenPdfAndReadEmbeddedAsync(PdfAttachmentItem pdf)
        {
            CancelWork();
            var token = _cts.Token;

            try
            {
                embListPanel.Controls.Clear();
                _embedded.Clear();
            }
            catch { }

            if (pdf == null || pdf.AttachmentIndex <= 0) return;

            var mail = GetCurrentMailItem();
            if (mail == null) return;

            try
            {
                Directory.CreateDirectory(TempDir);
                _currentPdfPath = Path.Combine(TempDir, Guid.NewGuid().ToString("N") + "_" + Sanitize(pdf.Name));

                await Task.Run(() =>
                {
                    if (token.IsCancellationRequested) return;
                    var att = mail.Attachments[pdf.AttachmentIndex];
                    att.SaveAsFile(_currentPdfPath);
                }, token);

                if (token.IsCancellationRequested) return;

                Dictionary<string, byte[]> embedded = await Task.Run(() =>
                {
                    if (token.IsCancellationRequested) return new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);
                    return ExtractEmbeddedFilesFromPdf(_currentPdfPath);
                }, token);

                if (token.IsCancellationRequested) return;

                _embedded = embedded ?? new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);

                foreach (var kv in _embedded)
                {
                    if (token.IsCancellationRequested) return;

                    string name = kv.Key;
                    byte[] bytes = kv.Value;

                    var row = CreateLinkRow(
                        name,
                        () =>
                        {
                            OpenBytesInBrowser(bytes, name);
                            return Task.CompletedTask;
                        },
                        () => DownloadBytes(bytes, name)
                    );

                    embListPanel.Controls.Add(row);
                }
            }
            catch
            {
                // swallow exceptions to protect Outlook
            }
        }

        private static Dictionary<string, byte[]> ExtractEmbeddedFilesFromPdf(string pdfPath)
        {
            var result = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);

            try
            {
                using (var reader = new PdfReader(pdfPath))
                using (var pdf = new PdfDocument(reader))
                {
                    var catalog = pdf.GetCatalog();
                    var nameTree = catalog.GetNameTree(iText.Kernel.Pdf.PdfName.EmbeddedFiles);
                    var names = nameTree.GetNames();

                    if (names == null || names.Count == 0)
                        return result;

                    foreach (var kv in names)
                    {
                        var keyName = (kv.Key != null) ? kv.Key.ToString() : null;
                        var fileSpecDict = kv.Value as iText.Kernel.Pdf.PdfDictionary;
                        if (fileSpecDict == null) continue;

                        string fileName = null;

                        var ufStr = fileSpecDict.GetAsString(iText.Kernel.Pdf.PdfName.UF);
                        if (ufStr != null) fileName = ufStr.ToUnicodeString();

                        if (string.IsNullOrWhiteSpace(fileName))
                        {
                            var fStr = fileSpecDict.GetAsString(iText.Kernel.Pdf.PdfName.F);
                            if (fStr != null) fileName = fStr.ToUnicodeString();
                        }

                        if (string.IsNullOrWhiteSpace(fileName))
                            fileName = string.IsNullOrWhiteSpace(keyName) ? "embedded.bin" : keyName;

                        var ef = fileSpecDict.GetAsDictionary(iText.Kernel.Pdf.PdfName.EF);
                        if (ef == null) continue;

                        iText.Kernel.Pdf.PdfStream stream = null;

                        var sUf = ef.GetAsStream(iText.Kernel.Pdf.PdfName.UF);
                        if (sUf != null) stream = sUf;

                        if (stream == null)
                        {
                            var sF = ef.GetAsStream(iText.Kernel.Pdf.PdfName.F);
                            if (sF != null) stream = sF;
                        }

                        if (stream == null) continue;

                        var bytes = stream.GetBytes();
                        if (bytes == null || bytes.Length == 0) continue;

                        result[fileName] = bytes;
                    }
                }
            }
            catch { }

            return result;
        }

        private void OpenBytesInBrowser(byte[] bytes, string filename)
        {
            try
            {
                Directory.CreateDirectory(TempDir);

                var safeName = Sanitize(filename ?? "file.bin");
                var tmp = Path.Combine(TempDir, Guid.NewGuid().ToString("N") + "_" + safeName);

                File.WriteAllBytes(tmp, bytes);

                Process.Start(new ProcessStartInfo
                {
                    FileName = tmp,
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                MessageBox.Show("Konnte Datei nicht öffnen: " + ex.Message);
            }
        }

        private void DownloadBytes(byte[] bytes, string filename)
        {
            try
            {
                using (var sfd = new SaveFileDialog
                {
                    FileName = filename ?? "download.bin",
                    Filter = "Alle Dateien (*.*)|*.*"
                })
                {
                    if (sfd.ShowDialog() != DialogResult.OK) return;
                    File.WriteAllBytes(sfd.FileName, bytes);
                }
            }
            catch { }
        }

        private Control CreateLinkRow(string text, Func<Task> onOpen, Action onDownload)
        {
            var row = new Panel
            {
                Width = 480,
                Height = 30,
                Margin = new Padding(0, 2, 0, 8),
                BackColor = Color.White
            };

            var link = new LinkLabel
            {
                Text = text,
                AutoSize = false,
                Left = 0,
                Top = 6,
                Width = 360,
                Height = 18,
                LinkColor = Color.FromArgb(20, 88, 170),
                ActiveLinkColor = Color.FromArgb(15, 70, 135),
                VisitedLinkColor = Color.FromArgb(20, 88, 170),
                Font = new Font("Segoe UI", 10f, FontStyle.Regular),
                LinkBehavior = LinkBehavior.HoverUnderline
            };
            tip.SetToolTip(link, text);

            link.LinkClicked += async (s, e) =>
            {
                try
                {
                    Debug.WriteLine("OPEN CLICK: " + text);
                    if (onOpen != null) await onOpen();
                }
                catch (Exception ex)
                {
                    Debug.WriteLine("OPEN ERROR: " + ex);
                    MessageBox.Show("Fehler beim Öffnen: " + ex.Message);
                }
            };

            var btn = new Button
            {
                Text = "Download",
                Width = 82,
                Height = 26,
                Top = 2,
                Left = 380,
                BackColor = Color.FromArgb(245, 247, 250),
                ForeColor = Color.FromArgb(45, 45, 45),
                FlatStyle = FlatStyle.Flat
            };
            btn.FlatAppearance.BorderColor = Color.FromArgb(220, 220, 220);
            btn.FlatAppearance.BorderSize = 1;
            btn.FlatAppearance.MouseOverBackColor = Color.FromArgb(235, 238, 243);
            btn.FlatAppearance.MouseDownBackColor = Color.FromArgb(225, 230, 238);

            btn.Click += (s, e) =>
            {
                try { if (onDownload != null) onDownload(); }
                catch (Exception ex)
                {
                    Debug.WriteLine("DL ERROR: " + ex);
                    MessageBox.Show("Download-Fehler: " + ex.Message);
                }
            };

            row.Controls.Add(link);
            row.Controls.Add(btn);

            row.Resize += (s, e) =>
            {
                btn.Left = row.Width - btn.Width;
                link.Width = Math.Max(120, btn.Left - 10);
            };

            return row;
        }

        private static string Sanitize(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return "file.bin";
            foreach (var c in Path.GetInvalidFileNameChars())
                s = s.Replace(c, '_');
            return s;
        }

        private sealed class PdfAttachmentItem
        {
            public int AttachmentIndex { get; set; } // 1-based in Outlook
            public string Name { get; set; }
        }
    }
}
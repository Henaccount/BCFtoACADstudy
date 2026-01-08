using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.GraphicsSystem;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Windows;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using System.Reflection.Emit;
using System.Reflection.Metadata;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml.Linq;
using static System.Net.Mime.MediaTypeNames;

[assembly: CommandClass(typeof(BcfAutoCAD.BcfClient))]

namespace BcfAutoCAD
{
    // Simple BCF client for AutoCAD
    // - Command BCFPANEL opens a resizable PaletteSet (panel)
    // - Panel contains a button to choose a .bcf (zip) file
    // - After loading, each issue folder is shown horizontally (one row / FlowLayoutPanel)
    // - Each issue shows stripped bcf text and a "Go to View" button
    // - Clicking the button sets the AutoCAD camera according to the .bcfv and opens a separate resizable image window showing the PNG
    public class BcfClient
    {
        private const string PanelTitle = "BCF Client";
        private PaletteSet _palette;
        private FlowLayoutPanel _issuesFlow;
        private System.Windows.Forms.OpenFileDialog _openDlg;
        private ImageWindow _imageWindow;

        [CommandMethod("BCFPANEL")]
        public void OpenPanel()
        {
            // Create palette if not exists
            if (_palette == null)
            {
                _palette = new PaletteSet(PanelTitle)
                {
                    DockEnabled = (DockSides)((int)DockSides.Left + (int)DockSides.Right + (int)DockSides.Top + (int)DockSides.Bottom),
                    Size = new System.Drawing.Size(800, 300)
                };

                // Create top-level user control
                var host = new System.Windows.Forms.UserControl() { Dock = DockStyle.Fill };

                var openButton = new Button() { Text = "Open .bcf file...", AutoSize = true, Anchor = AnchorStyles.Top | AnchorStyles.Left };
                openButton.Click += OpenButton_Click;

                _issuesFlow = new FlowLayoutPanel()
                {
                    Dock = DockStyle.Fill,
                    AutoScroll = true,
                    FlowDirection = System.Windows.Forms.FlowDirection.LeftToRight,
                    WrapContents = false,
                    Padding = new Padding(6),
                };

                var topPanel = new Panel() { Dock = DockStyle.Top, Height = 40 };
                topPanel.Controls.Add(openButton);

                host.Controls.Add(_issuesFlow);
                host.Controls.Add(topPanel);

                _palette.Add("BCF", host);
            }

            _palette.Visible = true;
            // Ensure image window exists (modeless resizable form)
            if (_imageWindow == null || _imageWindow.IsDisposed)
            {
                _imageWindow = new ImageWindow();
            }
        }

        private void OpenButton_Click(object sender, EventArgs e)
        {
            if (_openDlg == null)
            {
                _openDlg = new System.Windows.Forms.OpenFileDialog()
                {
                    Filter = "BCF files (*.bcf;*.zip)|*.bcf;*.zip|All files (*.*)|*.*",
                    Multiselect = false,
                    Title = "Select BCF (zip) file"
                };
            }

            if (_openDlg.ShowDialog() != DialogResult.OK)
                return;

            try
            {
                LoadBcfArchive(_openDlg.FileName);
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("Failed to load BCF: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void LoadBcfArchive(string path)
        {
            _issuesFlow.SuspendLayout();
            _issuesFlow.Controls.Clear();

            using (var fs = File.OpenRead(path))
            using (var za = new ZipArchive(fs, ZipArchiveMode.Read))
            {
                // Group entries by top-level folder name (issue)
                var groups = new Dictionary<string, List<ZipArchiveEntry>>(StringComparer.OrdinalIgnoreCase);

                foreach (var entry in za.Entries)
                {
                    if (string.IsNullOrEmpty(entry.FullName))
                        continue;

                    // Ignore root-level files
                    var parts = entry.FullName.Split(new[] { '/', '\\' }, StringSplitOptions.RemoveEmptyEntries);
                    if (parts.Length < 2)
                        continue;

                    var top = parts[0];
                    if (!groups.ContainsKey(top))
                        groups[top] = new List<ZipArchiveEntry>();
                    groups[top].Add(entry);
                }

                foreach (var kv in groups)
                {
                    var issueName = kv.Key;
                    var entries = kv.Value;

                    var bcfEntry = entries.FirstOrDefault(e => e.Name.EndsWith(".bcf", StringComparison.OrdinalIgnoreCase));
                    var bcfvEntry = entries.FirstOrDefault(e => e.Name.EndsWith(".bcfv", StringComparison.OrdinalIgnoreCase));
                    var pngEntry = entries.FirstOrDefault(e => e.Name.EndsWith(".png", StringComparison.OrdinalIgnoreCase) ||
                                                               e.Name.EndsWith(".jpg", StringComparison.OrdinalIgnoreCase) ||
                                                               e.Name.EndsWith(".jpeg", StringComparison.OrdinalIgnoreCase));

                    string bcfText = bcfEntry != null ? ReadAndStripXml(za, bcfEntry) : "(no .bcf)";
                    var bcfvXml = bcfvEntry != null ? ReadXDocument(za, bcfvEntry) : null;
                    byte[] imageBytes = null;
                    if (pngEntry != null)
                    {
                        using (var s = pngEntry.Open())
                        using (var ms = new MemoryStream())
                        {
                            s.CopyTo(ms);
                            imageBytes = ms.ToArray();
                        }
                    }

                    // Create a panel for this issue
                    var issuePanel = CreateIssuePanel(issueName, bcfText, bcfvXml, imageBytes);
                    _issuesFlow.Controls.Add(issuePanel);
                }
            }

            _issuesFlow.ResumeLayout();
        }

        private Panel CreateIssuePanel(string title, string text, XDocument bcfvXml, byte[] imageBytes)
        {
            var panel = new Panel()
            {
                Width = 320,
                Height = 240,
                Margin = new Padding(6),
                BorderStyle = BorderStyle.FixedSingle,
            };

            var lbl = new System.Windows.Forms.Label()
            {
                Text = title,
                Dock = DockStyle.Top,
                Height = 20,
                TextAlign = ContentAlignment.MiddleLeft
            };

            var txt = new TextBox()
            {
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Vertical,
                Dock = DockStyle.Fill,
                Text = text
            };

            var bottom = new Panel() { Dock = DockStyle.Bottom, Height = 34 };

            var btn = new Button() { Text = "Go to View", AutoSize = true, Anchor = AnchorStyles.Right | AnchorStyles.Top };
            btn.Click += (s, e) =>
            {
                // 1) set AutoCAD camera from bcfv
                if (bcfvXml != null)
                {
                    try
                    {
                        SetAutoCadViewFromBcfv(bcfvXml);
                    }
                    catch (System.Exception ex)
                    {
                        MessageBox.Show("Failed to set view: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                    MessageBox.Show("No bcfv (view) found for this issue.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

                // 2) show image in separate window
                if (imageBytes != null)
                {
                    try
                    {
                        if (_imageWindow == null || _imageWindow.IsDisposed)
                            _imageWindow = new ImageWindow();
                        _imageWindow.SetImage(imageBytes);
                        _imageWindow.Show();
                        _imageWindow.BringToFront();
                    }
                    catch (System.Exception ex)
                    {
                        MessageBox.Show("Failed to show image: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                    MessageBox.Show("No image found for this issue.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            };

            bottom.Controls.Add(btn);
            btn.Left = bottom.Width - btn.Width - 8;
            btn.Anchor = AnchorStyles.Top | AnchorStyles.Right;

            panel.Controls.Add(txt);
            panel.Controls.Add(bottom);
            panel.Controls.Add(lbl);
            return panel;
        }

        private static string ReadAndStripXml(ZipArchive za, ZipArchiveEntry entry)
        {
            try
            {
                using (var s = entry.Open())
                using (var sr = new StreamReader(s))
                {
                    var xml = sr.ReadToEnd();
                    // Load and extract text nodes
                    var doc = XDocument.Parse(xml);
                    var texts = doc.DescendantNodes().OfType<XText>().Select(t => t.Value.Trim()).Where(t => !string.IsNullOrEmpty(t));
                    return string.Join(" ", texts);
                }
            }
            catch
            {
                // fallback: return raw content as string without tags by regex
                using (var s = entry.Open())
                using (var sr = new StreamReader(s))
                {
                    var raw = sr.ReadToEnd();
                    var stripped = Regex.Replace(raw, "<.*?>", " ");
                    return Regex.Replace(stripped, "\\s+", " ").Trim();
                }
            }
        }

        private static XDocument ReadXDocument(ZipArchive za, ZipArchiveEntry entry)
        {
            try
            {
                using (var s = entry.Open())
                {
                    return XDocument.Load(s);
                }
            }
            catch
            {
                return null;
            }
        }

        // Signed angle from a to b around axis (all should be normalized for best results)
        private static double SignedAngle(Vector3d a, Vector3d b, Vector3d axis)
        {
            double sin = axis.DotProduct(a.CrossProduct(b));
            double cos = a.DotProduct(b);
            return Math.Atan2(sin, cos);
        }

        // Updated SetAutoCadViewFromBcfv method: parses IfcGuid from Selection/Component and exposes it
        // as the string variable `ifcguid` which is available where the view is applied.
        public void SetAutoCadViewFromBcfv(XDocument bcfvDoc)
        {
            if (bcfvDoc == null) return;

            // --- Apply the view in AutoCAD ---
            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            if (doc == null)
                throw new System.Exception("No active AutoCAD document");

            var ed = doc.Editor;

            // --- parse camera element (your existing logic) ---
            var cameraElement = bcfvDoc.Descendants()
                .FirstOrDefault(e => string.Equals(e.Name.LocalName, "PerspectiveCamera", StringComparison.OrdinalIgnoreCase) || string.Equals(e.Name.LocalName, "OrthogonalCamera", StringComparison.OrdinalIgnoreCase));


            if (cameraElement == null)
                throw new System.Exception("Camera element not found in .bcfv");

            var pos = TryParseTriple(cameraElement, "CameraViewPoint");
            var dir = TryParseTriple(cameraElement, "CameraDirection");
            var up = TryParseTriple(cameraElement, "CameraUpVector");
            double? fov = TryParseScalar(cameraElement, "FieldOfView");

            if (pos == null || dir == null)
                throw new System.Exception("CameraViewPoint or CameraDirection missing in .bcfv");

            // --- parse IfcGuid from Selection/Component (first Selected="true" or first Component) ---
            string ifcguid = null;
            try
            {
                var selectionElement = bcfvDoc.Descendants()
                    .FirstOrDefault(e => string.Equals(e.Name.LocalName, "Selection", StringComparison.OrdinalIgnoreCase));

                if (selectionElement != null)
                {
                    // Prefer Component elements marked Selected="true"; fall back to first Component
                    var component = selectionElement.Elements()
                        .FirstOrDefault(c => string.Equals(c.Name.LocalName, "Component", StringComparison.OrdinalIgnoreCase) &&
                                             string.Equals((string)c.Attribute("Selected"), "true", StringComparison.OrdinalIgnoreCase))
                        ?? selectionElement.Elements()
                        .FirstOrDefault(c => string.Equals(c.Name.LocalName, "Component", StringComparison.OrdinalIgnoreCase));

                    if (component != null)
                    {
                        var guidAttr = component.Attribute("IfcGuid") ?? component.Attribute("ifcGuid") ?? component.Attribute("IfcGUID");
                        if (guidAttr != null)
                            ifcguid = guidAttr.Value;
                    }
                }
            }
            catch
            {
                // swallow parsing errors for selection; ifcguid remains null
                ed.WriteMessage("\nerror parsing ifcguid");
            }

            // Now ifcguid contains the parsed IfcGuid (or null if none found).
            // You can use `ifcguid` below where the view gets updated.
            /*
            // --- Build AutoCAD types ---
            var camPos = new Point3d(pos.Item1, pos.Item2, pos.Item3);
            var camDir = new Vector3d(dir.Item1, dir.Item2, dir.Item3);
            var camUp = (up != null) ? new Vector3d(up.Item1, up.Item2, up.Item3) : Vector3d.YAxis;

            // target = eye + direction (BCF's CameraDirection points from eye toward the target)
            var target = camPos + camDir;

            // AutoCAD ViewDirection should point from target TO camera: camera - target
            var viewDirVec = (camPos - target);
            if (viewDirVec.Length == 0)
            {
                // fallback: invert camDir
                viewDirVec = camDir.Negate();
                if (viewDirVec.Length == 0)
                    viewDirVec = Vector3d.ZAxis.Negate(); // last resort
            }
            var viewDir = viewDirVec.GetNormal();

            // Compute distance 
            double distance = (camPos - target).Length;
            if (distance <= 1e-9) distance = camDir.Length;
            if (distance <= 1e-9) distance = 1.0;
            */

            /*using (doc.LockDocument())
            {
                var curView = ed.GetCurrentView();

                // Apply direction/target
                curView.ViewDirection = camDir.Negate(); // view direction = camera - target (negated camDir)
                                                         // Choose a reasonable target position in AutoCAD (BCF target = eye + direction; we used target variable)
                                                         // AutoCAD expects Target to be the point the camera looks at; set it to target
                curView.Target = target;

                // Optionally set center point / twist / height based on FOV
                // Example: set Height from FieldOfView (if provided)
                if (fov.HasValue)
                {
                    double fovVal = fov.Value;
                    // Convert degrees to radians heuristically if needed
                    if (fovVal > 2.0 * Math.PI)
                        fovVal = fovVal * Math.PI / 180.0;

                    fovVal = Math.Min(Math.Max(fovVal, 1e-6), Math.PI - 1e-6);
                    double viewHeight = 2.0 * distance * Math.Tan(fovVal / 2.0);
                    curView.Height = Math.Max(1e-9, viewHeight);

                    try
                    {
                        curView.LensLength = distance / Math.Tan(fovVal / 2.0);
                    }
                    catch { }
                }

                // If you want to use the parsed ifcguid to select or highlight something, `ifcguid` is available here:
                // Example (no-op): if (!string.IsNullOrEmpty(ifcguid)) {  }
                
                
                ed.SetCurrentView(curView);
            }*/

            //ed.WriteMessage("\nhandle: " +  ifcguid);
            ZoomToAndSelectByHandle(ifcguid);

            //ed.Command("_zoom", "a", "");
            //ed.Command("_zoom", "o", "(handent \"" + ifcguid + "\")", "");
            //ed.Command("_select", "(handent '" + ifcguid + "')", "");
            // For debugging you may show the ifcguid to the user (optional)
            // if (!string.IsNullOrEmpty(ifcguid))
            //     MessageBox.Show($"IfcGuid: {ifcguid}", "BCF Selection", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public static void ZoomToAndSelectByHandle(string handle)
        {
            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            if (doc == null) throw new InvalidOperationException("No active document.");
            var ed = doc.Editor;
            var db = doc.Database;

            // Resolve handle string to a Handle (try hex first, then decimal)
            Autodesk.AutoCAD.DatabaseServices.Handle acadHandle;
            try
            {
                long val = Convert.ToInt64(handle, 16); // try hex
                acadHandle = new Autodesk.AutoCAD.DatabaseServices.Handle(val);
            }
            catch
            {
                try
                {
                    long valDec = Convert.ToInt64(handle, CultureInfo.InvariantCulture); // decimal fallback
                    acadHandle = new Autodesk.AutoCAD.DatabaseServices.Handle(valDec);
                }
                catch
                {
                    MessageBox.Show($"Invalid handle: {handle}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }

            using (doc.LockDocument())
            using (var tr = db.TransactionManager.StartTransaction())
            {
                ObjectId id;
                try
                {
                    // GetObjectId(false, Handle, 0) resolves a handle to an ObjectId in the drawing
                    id = db.GetObjectId(false, acadHandle, 0);
                }
                catch
                {
                    MessageBox.Show($"Object with handle {handle} not found.", "Not found", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                if (id.IsNull)
                {
                    MessageBox.Show($"Object with handle {handle} not found.", "Not found", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // Try to open the entity and read its geometric extents
                Entity ent = null;
                Extents3d extents;
                bool gotExtents = false;
                try
                {
                    ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
                    if (ent != null)
                    {
                        // GeometricExtents might throw for some entities; wrap in try
                        extents = ent.GeometricExtents;
                        gotExtents = true;
                    }
                    else
                    {
                        MessageBox.Show("Object is not an entity.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
                catch
                {
                    // Failed to obtain geometric extents
                    ed.WriteMessage("\nFailed to obtain geometric extents");
                    gotExtents = false;
                    extents = new Extents3d();
                }

                // Get current view and prepare to update
                var curView = ed.GetCurrentView();

                if (gotExtents)
                {
                    // Center of extents:
                    var min = extents.MinPoint;
                    var max = extents.MaxPoint;
                    var center = new Point3d((min.X + max.X) * 0.5, (min.Y + max.Y) * 0.5, (min.Z + max.Z) * 0.5);

                    // Compute a simple view height (max dimension) and apply a margin
                    double sizeX = Math.Abs(max.X - min.X);
                    double sizeY = Math.Abs(max.Y - min.Y);
                    double sizeZ = Math.Abs(max.Z - min.Z);
                    double maxDim = Math.Max(sizeX, Math.Max(sizeY, sizeZ));
                    double marginFactor = 1.15; // leave some border
                    double newHeight = Math.Max(1e-6, maxDim * marginFactor);

                    // If current view is perspective, keep it; else use orthographic with current ViewDirection.
                    // Set the Target to the object's center and set Height.
                    curView.Target = center;
                    curView.Height = newHeight;

                    // Optionally: If you want to set ViewDirection to look from the camera viewpoint,
                    // compute it here and set curView.ViewDirection = ...
                    // For many cases keeping the existing ViewDirection is acceptable.
                }
                else
                {
                    // Fallback: couldn't obtain extents. Use the entity's position if available (like for DBPoint, Circle, etc.)
                    try
                    {
                        // many entities expose a Bounds or Position-like property — we attempt to get a bounding box
                        var box = ent.GeometricExtents; // sometimes this throws again; wrapped by outer try
                        var center = new Point3d((box.MinPoint.X + box.MaxPoint.X) * 0.5, (box.MinPoint.Y + box.MaxPoint.Y) * 0.5, (box.MinPoint.Z + box.MaxPoint.Z) * 0.5);
                        curView.Target = center;
                        curView.Height = Math.Max(1.0, (box.MaxPoint - box.MinPoint).Length * 1.2);
                    }
                    catch
                    {
                        // Last resort: do nothing to the view
                        ed.WriteMessage("\nLast resort: do nothing to the view");
                    }
                }

                // Commit the view update
                ed.SetCurrentView(curView);

                // Select the entity programmatically without ed.Command
                try
                {
                    ed.SetImpliedSelection(new ObjectId[] { id });
                }
                catch
                {
                    // Fallback: try to build a selection set if SetImpliedSelection fails
                    try
                    {
                        var ssFilter = new TypedValue[] { new TypedValue((int)DxfCode.Handle, handle) };
                        var ss = ed.SelectAll(new SelectionFilter(ssFilter));
                        if (ss.Status == PromptStatus.OK)
                        {
                            var ids = ss.Value.GetObjectIds();
                            ed.SetImpliedSelection(ids);
                        }
                    }
                    catch
                    {
                        // ignore
                    }
                }

                tr.Commit();
            }
        }
        private static double? TryParseScalar(XElement parent, string name)
        {
            var el = parent.Descendants().FirstOrDefault(e => string.Equals(e.Name.LocalName, name, StringComparison.OrdinalIgnoreCase));
            if (el == null) return null;
            double v;
            if (double.TryParse(el.Value, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out v))
                return v;

            // try attribute
            var attr = el.Attributes().FirstOrDefault();
            if (attr != null && double.TryParse(attr.Value, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out v))
                return v;

            return null;
        }

        private static Tuple<double, double, double> TryParseTriple(XElement parent, string name)
        {
            var el = parent.Descendants().FirstOrDefault(e => string.Equals(e.Name.LocalName, name, StringComparison.OrdinalIgnoreCase));
            if (el == null) return null;

            // Check for children X,Y,Z
            var xEl = el.Elements().FirstOrDefault(e => string.Equals(e.Name.LocalName, "X", StringComparison.OrdinalIgnoreCase));
            var yEl = el.Elements().FirstOrDefault(e => string.Equals(e.Name.LocalName, "Y", StringComparison.OrdinalIgnoreCase));
            var zEl = el.Elements().FirstOrDefault(e => string.Equals(e.Name.LocalName, "Z", StringComparison.OrdinalIgnoreCase));
            double x, y, z;
            if (xEl != null && yEl != null && zEl != null &&
                double.TryParse(xEl.Value, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out x) &&
                double.TryParse(yEl.Value, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out y) &&
                double.TryParse(zEl.Value, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out z))
            {
                return Tuple.Create(x, y, z);
            }

            // Check attributes x,y,z
            var ax = el.Attribute("x") ?? el.Attribute("X");
            var ay = el.Attribute("y") ?? el.Attribute("Y");
            var az = el.Attribute("z") ?? el.Attribute("Z");
            if (ax != null && ay != null && az != null &&
                double.TryParse(ax.Value, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out x) &&
                double.TryParse(ay.Value, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out y) &&
                double.TryParse(az.Value, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out z))
            {
                return Tuple.Create(x, y, z);
            }

            // If inner text contains three numbers
            var numbers = Regex.Matches(el.Value, @"-?\d+(\.\d+)?(?:[eE][+-]?\d+)?")
                               .Cast<Match>()
                               .Select(m => m.Value)
                               .ToArray();
            if (numbers.Length >= 3 &&
                double.TryParse(numbers[0], System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out x) &&
                double.TryParse(numbers[1], System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out y) &&
                double.TryParse(numbers[2], System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out z))
            {
                return Tuple.Create(x, y, z);
            }

            // Try elements named with coordinates directly under parent (e.g., Position/X etc handled above), else look for any descendant with three numbers
            var allText = string.Join(" ", el.DescendantNodes().OfType<XText>().Select(t => t.Value.Trim()));
            numbers = Regex.Matches(allText, @"-?\d+(\.\d+)?(?:[eE][+-]?\d+)?")
                           .Cast<Match>()
                           .Select(m => m.Value)
                           .ToArray();
            if (numbers.Length >= 3 &&
                double.TryParse(numbers[0], System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out x) &&
                double.TryParse(numbers[1], System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out y) &&
                double.TryParse(numbers[2], System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out z))
            {
                return Tuple.Create(x, y, z);
            }

            return null;
        }
    }

    // Simple separate resizable image window with PictureBox that scales
    public class ImageWindow : Form
    {
        private PictureBox _pic;

        public ImageWindow()
        {
            this.Text = "BCF Image";
            this.Size = new Size(800, 600);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.Sizable;
            _pic = new PictureBox() { Dock = DockStyle.Fill, SizeMode = PictureBoxSizeMode.Zoom, BackColor = Color.Gray };
            this.Controls.Add(_pic);
        }

        public void SetImage(byte[] data)
        {
            if (data == null) return;
            using (var ms = new MemoryStream(data))
            {
                var img = System.Drawing.Image.FromStream(ms);
                _pic.Image?.Dispose();
                _pic.Image = new Bitmap(img);
            }
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            base.OnFormClosed(e);
            _pic.Image?.Dispose();
        }

        
    }
}

using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace CertPhotoSorter
{
    internal sealed class ToggleSwitch : Control
    {
        private bool _checked;

        public ToggleSwitch()
        {
            SetStyle(ControlStyles.AllPaintingInWmPaint |
                     ControlStyles.OptimizedDoubleBuffer |
                     ControlStyles.ResizeRedraw |
                     ControlStyles.UserPaint, true);

            Cursor = Cursors.Hand;
            Size = new Size(46, 24);
        }

        public event EventHandler CheckedChanged;

        public bool Checked
        {
            get { return _checked; }
            set
            {
                if (_checked == value) return;
                _checked = value;
                Invalidate();
                var handler = CheckedChanged;
                if (handler != null) handler(this, EventArgs.Empty);
            }
        }

        public Color OnBackColor = Color.FromArgb(59, 130, 246);
        public Color OffBackColor = Color.FromArgb(209, 213, 219);
        public Color ThumbColor = Color.White;

        protected override void OnMouseUp(MouseEventArgs e)
        {
            base.OnMouseUp(e);
            if (e.Button == MouseButtons.Left)
            {
                Checked = !Checked;
            }
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;

            var rect = new Rectangle(0, 0, Width - 1, Height - 1);
            var radius = rect.Height;

            using (var path = CreateRoundRect(rect, radius))
            using (var backBrush = new SolidBrush(Checked ? OnBackColor : OffBackColor))
            using (var borderPen = new Pen(Color.FromArgb(190, 190, 190)))
            {
                e.Graphics.FillPath(backBrush, path);
                e.Graphics.DrawPath(borderPen, path);
            }

            var padding = 2;
            var thumbSize = rect.Height - padding * 2;
            var x = Checked ? (rect.Width - padding - thumbSize) : padding;
            var thumbRect = new Rectangle(x, padding, thumbSize, thumbSize);

            using (var thumbBrush = new SolidBrush(ThumbColor))
            using (var thumbPen = new Pen(Color.FromArgb(180, 180, 180)))
            {
                e.Graphics.FillEllipse(thumbBrush, thumbRect);
                e.Graphics.DrawEllipse(thumbPen, thumbRect);
            }
        }

        private static GraphicsPath CreateRoundRect(Rectangle rect, int radius)
        {
            var path = new GraphicsPath();
            var d = radius;

            path.AddArc(rect.X, rect.Y, d, d, 180, 90);
            path.AddArc(rect.Right - d, rect.Y, d, d, 270, 90);
            path.AddArc(rect.Right - d, rect.Bottom - d, d, d, 0, 90);
            path.AddArc(rect.X, rect.Bottom - d, d, d, 90, 90);
            path.CloseFigure();

            return path;
        }
    }
}


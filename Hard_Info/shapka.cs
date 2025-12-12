using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing;
using System.Drawing.Drawing2D;

namespace Hard_Info
{
    public partial class shapka : Component
    {
        bool clouseHowe = false;
        Rectangle rectangle_clouse = new Rectangle();

        bool MousePresed = false;
        Point clickPosishen;
        Point mouseStartPosishen;

        bool minimizeHovered = false;
        Rectangle rectangle_Min = new Rectangle();

        public Form Form { get; set; }
        private fStyle formStyle = fStyle.UserStyle;
        private int HeaderHeigt = 25;

        Color color = Color.DimGray;
        Pen pen = new Pen(Color.White) { Width = 1.55F };

        public StringFormat sf = new StringFormat();
        public Font Font = new Font("Arial", 8.75f, FontStyle.Regular);

        private Size icon_size = new Size(14, 14);
        public fStyle Formstyle
        {
            get => formStyle;
            set
            {
                formStyle = value;
                sign();
            }
        }

        public enum fStyle // Набор стилей
        {
            None,

            UserStyle,

            SimpleDark,
            TelegramStyle
        }

        public shapka()
        {
            InitializeComponent();
        }

        public shapka(IContainer container)
        {
            container.Add(this);

            InitializeComponent();
        }

        private void sign()
        {
            if (Form != null)
            {
                Form.Load += Form_Load;
            }
        }

        private void Apply()
        {
            sf.Alignment = StringAlignment.Near;
            sf.LineAlignment = StringAlignment.Center;
            Form.FormBorderStyle = FormBorderStyle.None;
            Size minimumSize = new Size(100, 50);
            if (Form.MinimumSize.Width < minimumSize.Width || Form.MinimumSize.Height < minimumSize.Height)
                Form.MinimumSize = minimumSize;
            Ofset_controls();
            SetDoubleBuffered(Form);
            Form.Paint += Form_Paint;
            Form.MouseDown += Form_MouseDown;
            Form.MouseUp += Form_MouseUp;
            Form.MouseMove += Form_MouseMove;
            Form.MouseLeave += Form_MouseLeave;
        }

        private void Form_MouseLeave(object sender, EventArgs e)
        {
            clouseHowe = false;
            minimizeHovered = false;
            Form.Invalidate();

        }

        private void Form_MouseMove(object sender, MouseEventArgs e)
        {
            if (MousePresed)
            {
                Size frmOfset = new Size(Point.Subtract(Cursor.Position, new Size(clickPosishen)));
                Form.Location = Point.Add(mouseStartPosishen, frmOfset);
            }
            else
            {
                if (rectangle_clouse.Contains(e.Location))
                {
                    if (clouseHowe == false)
                    {
                        clouseHowe = true;
                        Form.Invalidate();
                    }
                }
                else
                {
                    if (clouseHowe == true)
                    {
                        clouseHowe = false;
                        Form.Invalidate();
                    }
                }
                if (rectangle_Min.Contains(e.Location)){
                    if (minimizeHovered == false)
                    {
                        minimizeHovered = true;
                        Form.Invalidate();
                    }
                }
                else
                    if (minimizeHovered == true)
                    {
                        minimizeHovered = false;
                        Form.Invalidate();
                    }
            }
        }

        private void Form_MouseUp(object sender, MouseEventArgs e)
        {
            MousePresed = false;
            if (e.Button == MouseButtons.Left)
            {
                if (rectangle_clouse.Contains(e.Location))
                    Form.Close();
                if (rectangle_Min.Contains(e.Location))
                    Form.WindowState = FormWindowState.Minimized;
            }

        }

        private void Form_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Location.Y <= HeaderHeigt)
            {
                MousePresed = true;
                clickPosishen = Cursor.Position;
                mouseStartPosishen = Form.Location;
            }
        }

        private void Ofset_controls()
        {
            Form.Height = Form.Height + HeaderHeigt;
            foreach (Control control in Form.Controls)
            {
                control.Location = new Point(control.Location.X, control.Location.Y + HeaderHeigt);
                control.Refresh();
            }
        }

        private void Form_Load(object sender, EventArgs e)
        {
            Apply();
        }
        private void Form_Paint(object sender, PaintEventArgs e)
        {
            DrawStyle(e.Graphics);
        }

        private void DrawStyle(Graphics graphics)
        {
            graphics.SmoothingMode = SmoothingMode.HighQuality;
            Rectangle rectangle_heder = new Rectangle(0, 0, Form.Width - 1, HeaderHeigt);
            Rectangle rectangle_boder = new Rectangle(0, 0, Form.Width - 1, Form.Height - 1);
            Rectangle rectangle_text = new Rectangle(rectangle_heder.X + 20, rectangle_heder.Y,
                rectangle_heder.Width, rectangle_heder.Height);
            Rectangle rectangle_icon = new Rectangle(rectangle_heder.Height / 2 - icon_size.Width / 2,
                rectangle_heder.Height / 2 - icon_size.Height / 2, icon_size.Width, icon_size.Height);
            rectangle_clouse = new Rectangle(rectangle_heder.Width - rectangle_heder.Height,
                rectangle_heder.Y, rectangle_heder.Height, rectangle_heder.Height);
            Rectangle rectangle_X = new Rectangle(rectangle_clouse.X + rectangle_clouse.Width / 2 - 5,
                rectangle_clouse.Height / 2 - 5, 10, 10);
            rectangle_Min = new Rectangle(rectangle_heder.Width - rectangle_heder.Height * 2,
                rectangle_heder.Y, rectangle_heder.Height, rectangle_heder.Height);
            Rectangle rectangle_ = new Rectangle(rectangle_Min.X + rectangle_Min.Width / 2 - 5,
                rectangle_clouse.Height / 2 - 5, 10, 10);


            //Шапка
            graphics.DrawRectangle(new Pen(color), rectangle_heder);
            graphics.FillRectangle(new SolidBrush(color), rectangle_heder);
            //Кнопка X
            graphics.DrawRectangle(new Pen(clouseHowe ? Color.Red : color), rectangle_clouse);
            graphics.FillRectangle(new SolidBrush(clouseHowe ? Color.Red : color), rectangle_clouse);
            DrowCrossheir(graphics, rectangle_X, pen);
            //Кнопка _
            graphics.DrawRectangle(new Pen(minimizeHovered ? Color.Blue : color), rectangle_Min);
            graphics.FillRectangle(new SolidBrush(minimizeHovered ? Color.Blue : color), rectangle_Min);
            DrowCrossheir_(graphics, rectangle_, pen);
            //Иконка
            graphics.DrawImage(Form.Icon.ToBitmap(), rectangle_icon);
            //Обводка
            graphics.DrawRectangle(new Pen(Color.Black), rectangle_boder);
            //Текст заголовка
            graphics.DrawString(Form.Text, Font, new SolidBrush(Color.White), rectangle_text, sf);
        }

        private void DrowCrossheir(Graphics graphics, Rectangle rectangle, Pen p)
        {
            graphics.DrawLine(p, rectangle.X, rectangle.Y, rectangle.X + rectangle.Width, rectangle.Y + rectangle.Height);
            graphics.DrawLine(p, rectangle.X + rectangle.Width, rectangle.Y, rectangle.X, rectangle.Y + rectangle.Height);
        }

        private void DrowCrossheir_(Graphics graphics, Rectangle rectangle, Pen p)
        {
            graphics.DrawLine(p, rectangle.X, rectangle.Y + rectangle.Height, rectangle.X + rectangle.Width, rectangle.Y + rectangle.Height);
            //graphics.DrawLine(p, rectangle.X + rectangle.Width, rectangle.Y, rectangle.X, rectangle.Y + rectangle.Height);
        }

        public static void SetDoubleBuffered(Control c)
        {
            if (SystemInformation.TerminalServerSession)
                return;

            System.Reflection.PropertyInfo pDoubleBuffered =
                  typeof(Control).GetProperty(
                        "DoubleBuffered",
                        System.Reflection.BindingFlags.NonPublic |
                        System.Reflection.BindingFlags.Instance);

            pDoubleBuffered.SetValue(c, true, null);
        }
    }
}

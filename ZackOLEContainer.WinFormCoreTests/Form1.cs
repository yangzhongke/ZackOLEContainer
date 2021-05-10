using System;
using System.Windows.Forms;
using ZackOLEContainer.WinFormCore;

namespace ZackOLEContainer.WinFormCoreTests
{
    public partial class Form1 : Form
    {
        private OLEContainer oleContainer;
        public Form1()
        {
            InitializeComponent();
            this.oleContainer = new OLEContainer();
            this.oleContainer.Dock = DockStyle.Fill;
            this.Controls.Add(this.oleContainer);
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // host.Open(@"E:\主同步盘\我的坚果云\个人资料\201810杨中科体检.docx");
            this.oleContainer.OpenFile(@"E:\主同步盘\我的坚果云\UoZ\SE101-玩着学编程\Part3课件\5-搞“对象”.pptx");
        }
    }
}

using System.Windows.Forms;
using System.Windows.Forms.Integration;

namespace TestOutlookAddin
{
    public partial class TaskPaneContainerControl : UserControl
    {
        public TaskPaneContainerControl()
        {
            InitializeComponent();
            Controls.Add(new ElementHost { Child = new UserControlWpf(), Dock = DockStyle.Fill });
        }
    }
}

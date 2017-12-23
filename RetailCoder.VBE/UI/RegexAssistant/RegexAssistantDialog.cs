using System;
using System.Windows.Forms;
using Rubberduck.UI.Command;

namespace Rubberduck.UI.RegexAssistant
{
    public partial class RegexAssistantDialog : Form
    {
        public RegexAssistantDialog()
        {
            InitializeComponent();
            ViewModel = new RegexAssistantViewModel();
        }

        public RegexAssistantDialog(RegexWithOption regex)
        {
            InitializeComponent();
            ViewModel = new RegexAssistantViewModel(regex);
        }

        private RegexAssistantViewModel _viewModel;

        private RegexAssistantViewModel ViewModel { get { return _viewModel; }
        set
            {
                _viewModel = value;
                
                RegexAssistant.DataContext = _viewModel;
            }
        }

        private void CloseButton_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}

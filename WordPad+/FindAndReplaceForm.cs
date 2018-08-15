using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using WordPad;

namespace WordPad_
{
    public partial class FindAndReplaceForm : Form
    {
        private MainForm _mainForm;
        private GroupBox _directionGB;
        private RadioButton _upRB;
        private RadioButton _downRB;
        private Button _cancelButton;
        private Label _toLB;
        private TextBox _toTB;
        private Button _replaceButton;
        private Button _replaceAllButton;
        public FindAndReplaceForm(MainForm mainForm, bool findForm)
        {
            _mainForm = mainForm;
            InitializeComponent();
            if (findForm)
            {
                FindForm();
            }
            else
            {
                ReplaceForm();
            }
            findTB.Text = _mainForm.DataTextBox.SelectedText;
        }

        private void FindForm()
        {
            Name = "Find";
            findWhatLabel.Text = "Find what:";
            _directionGB = new GroupBox();
            _directionGB.Name = "directionGB";
            _directionGB.Text = "Direction";
            _directionGB.Size = new Size(107, 47);
            _directionGB.Location = new Point(154, 38);
            Controls.Add(_directionGB);
            _upRB = new RadioButton();
            _upRB.Parent = _directionGB;
            _upRB.Name = "upRB";
            _upRB.Text = "Up";
            _upRB.Location = new Point(6, 19);
            _upRB.Size = new Size(39, 17);
            _downRB = new RadioButton();
            _downRB.Parent = _directionGB;
            _downRB.Name = "downRB";
            _downRB.Text = "Down";
            _downRB.Checked = true;
            _downRB.Location = new Point(51, 19);
            _downRB.Size = new Size(53, 17);
            _cancelButton = new Button();
            _cancelButton.Name = "cancelButton";
            _cancelButton.Text = "Cancel";
            _cancelButton.Location = new Point(267, 39);
            _cancelButton.Click += CancelButtonOnClick;
            Controls.Add(_cancelButton);
        }

        private void ReplaceForm()
        {
            Name = "Replace";
            this.Size = new Size(370, 172);
            _toLB = new Label();
            _toLB.Name = "toLB";
            _toLB.Text = "To:";
            _toLB.Location = new Point(22, 47);
            _toLB.Size = new Size(23, 13);
            this.Controls.Add(_toLB);
            _toTB = new TextBox();
            _toTB.Name = "toTB";
            _toTB.Size = new Size(187, 20);
            _toTB.Location = new Point(74, 44);
            Controls.Add(_toTB);
            _replaceButton = new Button();
            _replaceButton.Name = "replaceButton";
            _replaceButton.Text = "Replace";
            _replaceButton.Location = new Point(267, 41);
            _replaceButton.Click += ReplaceButtonOnClick;
            Controls.Add(_replaceButton);
            _replaceAllButton = new Button();
            _replaceAllButton.Name = "replaceAllButton";
            _replaceAllButton.Text = "Replace all";
            _replaceAllButton.Location = new Point(267, 71);
            _replaceAllButton.Click += ReplaceAllButtonOnClick;
            Controls.Add(_replaceAllButton);
            _cancelButton = new Button();
            _cancelButton.Name = "cancelButton";
            _cancelButton.Text = "Cancel";
            _cancelButton.Location = new Point(267, 100);
            _cancelButton.Click += CancelButtonOnClick;
            Controls.Add(_cancelButton);
            caseSensitiveCB.Location = new Point(13, 104);
            findWhatLabel.Text = "From:";
        }

        private void ReplaceAllButtonOnClick(object sender, EventArgs eventArgs)
        {
            _mainForm.DataTextBox.SelectionStart = 0;
            _mainForm.DataTextBox.SelectionLength = 0;
            while (true)
            {
                int selectionStart = _mainForm.DataTextBox.SelectionStart;
                FindDown();
                if (selectionStart == _mainForm.DataTextBox.SelectionStart)
                {
                    MessageBox.Show("Replace ended!", "Result", MessageBoxButtons.OK);
                    break;
                }
                _mainForm.DataTextBox.SelectedText = _toTB.Text;
            }
        }

        private void ReplaceButtonOnClick(object sender, EventArgs eventArgs)
        {
            if (_mainForm.DataTextBox.SelectedText == String.Empty)
            {
                findNextButton_Click(sender, eventArgs);
            }
            if (_mainForm.DataTextBox.SelectedText != String.Empty)
            {
                _mainForm.DataTextBox.SelectedText = _toTB.Text;
            }
        }

        private void CancelButtonOnClick(object sender, EventArgs eventArgs)
        {
            Close();
        }

        private void findNextButton_Click(object sender, EventArgs e)
        {
            FindInTheText();
        }

        private void FindInTheText()
        {
            if (_downRB == null)
            {
                FindDown();
                _mainForm.Activate();
            }
            else
            {
                if (_downRB.Checked)
                {
                    FindDown();
                }
                else
                {
                    FindUp();
                }
                _mainForm.Activate();
            }
        }

        private void FindDown()
        {
            String text =
                _mainForm.Data.Substring(_mainForm.DataTextBox.SelectionStart +
                                         _mainForm.DataTextBox.SelectionLength);
            if (text.Length > 0)
            {
                int firstPos;

                if (caseSensitiveCB.Checked == false)
                {
                    firstPos = text.IndexOf(findTB.Text, StringComparison.OrdinalIgnoreCase);
                }
                else
                {
                    firstPos = text.IndexOf(findTB.Text, StringComparison.Ordinal);
                }

                if (firstPos != -1)
                {
                    _mainForm.DataTextBox.Select(firstPos + (_mainForm.Data.Length - text.Length), findTB.Text.Length);
                }
                else
                {
                    MessageBox.Show(@"Find nothing!", @"Result", MessageBoxButtons.OK);
                }
            }
            else
            {
                MessageBox.Show(@"Find nothing!", @"Result", MessageBoxButtons.OK);
            }
        }

        private void FindUp()
        {
            String text = _mainForm.Data.Substring(0, _mainForm.DataTextBox.SelectionStart);
            if (text.Length > 0)
            {
                int firstPos;

                if (caseSensitiveCB.Checked == false)
                {
                    firstPos = text.LastIndexOf(findTB.Text, StringComparison.OrdinalIgnoreCase);
                }
                else
                {
                    firstPos = text.LastIndexOf(findTB.Text, StringComparison.Ordinal);
                }

                if (firstPos != -1)
                {
                    _mainForm.DataTextBox.Select(firstPos, findTB.Text.Length);
                }
                else
                {
                    MessageBox.Show(@"Find nothing!", @"Result", MessageBoxButtons.OK);
                }
            }
            else
            {
                MessageBox.Show(@"Find nothing!", @"Result", MessageBoxButtons.OK);
            }
        }
    }
}

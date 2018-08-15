using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Text;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using WordPad_;

namespace WordPad
{
    public partial class MainForm : Form
    {
        private FindAndReplaceForm _findAndReplaceForm = null;
        private List<Document> _documents = new List<Document>();
        public MainForm()
        {
            InitializeComponent();
            LoadFonts();
            LoadSizes();
            LoadStyles();
            MainPreparing();
            //((RichTextBox)tabControl.SelectedTab.Controls[0]).DragEnter += RichTextBoxOnDragEnter;
            //((RichTextBox)tabControl.SelectedTab.Controls[0]).DragDrop += RichTextBoxOnDragDrop;
        }

        public string Data
        {
            get { return ((RichTextBox)tabControl.SelectedTab.Controls[0]).Text; }
        }

        public RichTextBox DataTextBox
        {
            get { return ((RichTextBox)tabControl.SelectedTab.Controls[0]); }
        }

        #region PreparedMethods
        private void LoadFonts()
        {
            var fontsCollection = new InstalledFontCollection();
            var ff = fontsCollection.Families;
            foreach (var item in ff)
            {
                fontsCollectionCB.Items.Add(item.Name);
            }
            fontsCollectionCB.SelectedIndex = 61;
        }
        private void LoadSizes()
        {
            fontsSizeCB.Items.Add(8);
            fontsSizeCB.Items.Add(9);
            fontsSizeCB.Items.Add(10);
            fontsSizeCB.Items.Add(11);
            fontsSizeCB.Items.Add(12);
            fontsSizeCB.Items.Add(14);
            fontsSizeCB.Items.Add(16);
            fontsSizeCB.Items.Add(18);
            fontsSizeCB.Items.Add(20);
            fontsSizeCB.Items.Add(22);
            fontsSizeCB.Items.Add(24);
            fontsSizeCB.Items.Add(26);
            fontsSizeCB.Items.Add(28);
            fontsSizeCB.Items.Add(36);
            fontsSizeCB.Items.Add(48);
            fontsSizeCB.Items.Add(72);
            fontsSizeCB.SelectedIndex = 3;
        }
        private void LoadStyles()
        {
            fontStyleCB.Items.Add(FontStyle.Regular.ToString());
            fontStyleCB.Items.Add(FontStyle.Bold.ToString());
            fontStyleCB.Items.Add(FontStyle.Italic.ToString());
            fontStyleCB.Items.Add(FontStyle.Strikeout.ToString());
            fontStyleCB.Items.Add(FontStyle.Underline.ToString());
            fontStyleCB.SelectedIndex = 0;
        }
        private void MainPreparing()
        {
            //Create tag for first tab
            var newDoc = new Document();
            tabControl.SelectedTab.Tag = newDoc;
            _documents.Add(newDoc);
        }
        #endregion

        #region MethodsForEventHandling
        private void newDocumentButton_Click(object sender, EventArgs e)
        {
            TabPage page = new TabPage("newDoc");
            var richTextBox = new RichTextBox();
            richTextBox.Location = new Point(3, 3);
            richTextBox.Dock = DockStyle.Fill;
            richTextBox.TextChanged += richTextBox_TextChanged;
            richTextBox.SelectionChanged += richTextBox_SelectionChanged;
            page.Controls.Add(richTextBox);
            tabControl.TabPages.Add(page);
            tabControl.SelectedTab = page;
            var newDoc = new Document();
            tabControl.SelectedTab.Tag = newDoc;
            _documents.Add(newDoc);
        }
        private void openDocumentButton_Click(object sender, EventArgs e)
        {
            if (DialogResult.OK == openFileDialog.ShowDialog())
            {
                foreach (var fileName in openFileDialog.FileNames)
                {
                    OpenFile(fileName);
                }
            }
        }
        private void printDocumentButton_Click(object sender, EventArgs e)
        {
            printDialog.ShowDialog();
            //((RichTextBox)tabControl.SelectedTab.Controls[0]).
        }
        private void fontsCollectionCB_SelectedIndexChanged(object sender, EventArgs e)
        {
            SetSettingsForText();
        }
        private void fontsSizeCB_TextChanged(object sender, EventArgs e)
        {
            try
            {
                Int32.Parse(fontsSizeCB.Text);
            }
            catch (Exception exception)
            {
                return;
            }
            SetSettingsForText();
        }
        private void fontStyleCB_SelectedIndexChanged(object sender, EventArgs e)
        {
            SetSettingsForText();
        }
        private void tabControl_Selected(object sender, TabControlEventArgs e)
        {
            SetSettingsForText();
        }
        private void richTextBox_TextChanged(object sender, EventArgs e)
        {
            ((Document)tabControl.SelectedTab.Tag).IsEdit = true;
        }
        private void closeSelectedTabToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (((Document)tabControl.SelectedTab.Tag).IsEdit || ((Document)tabControl.SelectedTab.Tag).IsEdit && ((Document)tabControl.SelectedTab.Tag).PathDocument == null)
            {
                var res = MessageBox.Show("Do you want to save changes?", "Closing Tab", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (res == DialogResult.Yes)
                {
                    SaveDoc();
                }
            }
            ((RichTextBox)tabControl.SelectedTab.Controls[0]).Dispose();
            tabControl.SelectedTab.Dispose();
            tabControl.SelectedIndex = tabControl.TabPages.Count - 1;
            if (tabControl.TabPages.Count == 0)
            {
                this.Close();
            }
        }
        private void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveDoc();
        }
        private void saveAsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveAsDoc();
        }
        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("WordPad, C#, WindowsForms, 2017, Nikita Zotov", "ShortInformation", MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }
        private void fontStyleStripButton_Click(object sender, EventArgs e)
        {
            SetSettingsForText();
        }
        private void leftStripButton_Click(object sender, EventArgs e)
        {
            ((RichTextBox)tabControl.SelectedTab.Controls[0]).SelectionAlignment = HorizontalAlignment.Left;
        }
        private void centerStripButton_Click(object sender, EventArgs e)
        {
            ((RichTextBox)tabControl.SelectedTab.Controls[0]).SelectionAlignment = HorizontalAlignment.Center;
        }
        private void rightStripButton_Click(object sender, EventArgs e)
        {
            ((RichTextBox)tabControl.SelectedTab.Controls[0]).SelectionAlignment = HorizontalAlignment.Right;
        }
        private void colorStripButton_Click(object sender, EventArgs e)
        {
            if (colorDialog.ShowDialog() == DialogResult.OK)
            {
                ((RichTextBox)tabControl.SelectedTab.Controls[0]).SelectionColor = colorDialog.Color;
            }
        }
        private void listStripButton_Click(object sender, EventArgs e)
        {
            ((RichTextBox)tabControl.SelectedTab.Controls[0]).SelectionBullet = listStripButton.Checked;
        }
        private void findButton_Click(object sender, EventArgs e)
        {
            if (_findAndReplaceForm == null)
            {
                _findAndReplaceForm = new FindAndReplaceForm(this, true);
                _findAndReplaceForm.FormClosed += (s, a) => _findAndReplaceForm = null;
                _findAndReplaceForm.Show();
            }
            else
            {
                _findAndReplaceForm.Activate();
            }
        }
        private void replaceButton_Click(object sender, EventArgs e)
        {
            if (_findAndReplaceForm == null)
            {
                _findAndReplaceForm = new FindAndReplaceForm(this, false);
                _findAndReplaceForm.FormClosed += (s, a) => _findAndReplaceForm = null;
                _findAndReplaceForm.Show();
            }
            else
            {
                _findAndReplaceForm.Activate();
            }
        }
        private void cutButton_Click(object sender, EventArgs e)
        {
            ((RichTextBox)tabControl.SelectedTab.Controls[0]).Cut();
        }
        private void copyButton_Click(object sender, EventArgs e)
        {
            ((RichTextBox)tabControl.SelectedTab.Controls[0]).Copy();
        }
        private void pasteButton_Click(object sender, EventArgs e)
        {
            ((RichTextBox)tabControl.SelectedTab.Controls[0]).Paste();
        }
        private void undoButton_Click(object sender, EventArgs e)
        {
            ((RichTextBox)tabControl.SelectedTab.Controls[0]).Undo();
        }
        private void redoButton_Click(object sender, EventArgs e)
        {
            ((RichTextBox)tabControl.SelectedTab.Controls[0]).Redo();
        }
        private void deleteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ((RichTextBox)tabControl.SelectedTab.Controls[0]).SelectedText = "";
        }
        private void selectAllToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ((RichTextBox)tabControl.SelectedTab.Controls[0]).SelectAll();
        }
        private void timeAndDateToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ((RichTextBox)tabControl.SelectedTab.Controls[0]).SelectedText = DateTime.Now.ToLongDateString();
        }
        private void wordWrapToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ((RichTextBox)tabControl.SelectedTab.Controls[0]).WordWrap = wordWrapToolStripMenuItem.Checked;
        }
        private void fontToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (DialogResult.OK == fontDialog.ShowDialog())
            {
                boldStripButton.Checked = fontDialog.Font.Bold;
                cursiveStripButton.Checked = fontDialog.Font.Italic;
                underlineStripButton.Checked = fontDialog.Font.Underline;
                if (fontDialog.Font.Strikeout)
                {
                    fontStyleCB.SelectedIndex = 3;
                }
                fontsCollectionCB.Text = fontDialog.Font.Name;
                fontsSizeCB.Text = fontDialog.Font.Size.ToString();
                SetSettingsForText();
            }
        }
        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            foreach (TabPage tabPage in tabControl.TabPages)
            {
                tabControl.SelectedTab = tabPage;
                closeSelectedTabToolStripMenuItem_Click(sender, null);
            }
        }
        private void richTextBox_SelectionChanged(object sender, EventArgs e)
        {
            //if (((RichTextBox)tabControl.SelectedTab.Controls[0]).SelectionFont == null)
            //{
            //    boldStripButton.Checked = false;
            //    cursiveStripButton.Checked = false;
            //    underlineStripButton.Checked = false;
            //    fontStyleCB.Text = "";
            //    fontsCollectionCB.Text = "";
            //    fontsSizeCB.Text = "";
            //    return;
            //}
            boldStripButton.Checked = ((RichTextBox)tabControl.SelectedTab.Controls[0]).SelectionFont.Bold;
            cursiveStripButton.Checked = ((RichTextBox)tabControl.SelectedTab.Controls[0]).SelectionFont.Italic;
            underlineStripButton.Checked = ((RichTextBox)tabControl.SelectedTab.Controls[0]).SelectionFont.Underline;
            if (((RichTextBox)tabControl.SelectedTab.Controls[0]).SelectionFont.Strikeout)
            {
                fontStyleCB.SelectedIndex = 3;
            }
            fontsCollectionCB.Text = ((RichTextBox)tabControl.SelectedTab.Controls[0]).SelectionFont.Name;
            fontsSizeCB.Text = ((RichTextBox)tabControl.SelectedTab.Controls[0]).SelectionFont.Size.ToString();
        }

        //private void RichTextBoxOnDragDrop(object sender, DragEventArgs dragEventArgs)
        //{
        //    OpenFile(dragEventArgs.Data.GetData(dragEventArgs.Data.GetFormats()[0]).ToString());
        //}
        //private void RichTextBoxOnDragEnter(object sender, DragEventArgs dragEventArgs)
        //{
        //    // при попадании на адресат формируем соответствующую иконку 
        //    // для курсора
        //    if (dragEventArgs.Data.GetDataPresent(DataFormats.StringFormat) || dragEventArgs.Data.GetDataPresent(DataFormats.Rtf) || dragEventArgs.Data.GetDataPresent(DataFormats.Text))
        //        dragEventArgs.Effect = DragDropEffects.Copy;
        //    else
        //        dragEventArgs.Effect = DragDropEffects.None;
        //}
        #endregion

        #region ClosedInternalMethods
        private void SetSettingsForText()
        {
            if (fontsCollectionCB.Items.Count > 0 && fontsSizeCB.Items.Count > 0 && fontStyleCB.Items.Count > 0 && tabControl.TabCount > 0)
            {
                ((RichTextBox)tabControl.SelectedTab.Controls[0]).SelectionFont = new Font((string)fontsCollectionCB.SelectedItem, Single.Parse(fontsSizeCB.Text), GetFontStyle());
            }
        }
        private void SaveDoc()
        {
            if (((Document)tabControl.SelectedTab.Tag).PathDocument == null)
            {
                SaveAsDoc();
            }
            else
            {
                ((RichTextBox)tabControl.SelectedTab.Controls[0]).SaveFile(
                    ((Document)tabControl.SelectedTab.Tag).PathDocument, RichTextBoxStreamType.RichText);
            }
        }
        private void SaveAsDoc()
        {
            saveFileDialog.FileName = tabControl.SelectedTab.Text;
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                ((Document)tabControl.SelectedTab.Tag).PathDocument = saveFileDialog.FileName;
                ((RichTextBox)tabControl.SelectedTab.Controls[0]).SaveFile(
                    ((Document)tabControl.SelectedTab.Tag).PathDocument, RichTextBoxStreamType.RichText);
            }
        }
        private FontStyle GetFontStyle()
        {
            FontStyle fontStyle = (FontStyle)Enum.Parse(typeof(FontStyle), fontStyleCB.SelectedItem.ToString());
            if (boldStripButton.Checked)
            {
                fontStyle = fontStyle | FontStyle.Bold;
            }
            if (cursiveStripButton.Checked)
            {
                fontStyle = fontStyle | FontStyle.Italic;
            }
            if (underlineStripButton.Checked)
            {
                fontStyle = fontStyle | FontStyle.Underline;
            }
            return fontStyle;
        }
        private void OpenFile(String fileName)
        {
            var isOk = true;
            foreach (TabPage tab in tabControl.TabPages)
            {
                if (fileName == ((Document)tab.Tag).PathDocument)
                {
                    MessageBox.Show("This file is already open!", "Ooops", MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    tabControl.SelectedTab = tab;
                    isOk = false;
                }
            }
            if (isOk)
            {
                var newDoc = new Document();
                newDoc.PathDocument = fileName;
                TabPage page = new TabPage(newDoc.GetNameDoc());
                var richTextBox = new RichTextBox();
                richTextBox.Location = new Point(3, 3);
                richTextBox.Dock = DockStyle.Fill;
                richTextBox.TextChanged += richTextBox_TextChanged;
                richTextBox.EnableAutoDragDrop = true;
                richTextBox.SelectionChanged += richTextBox_SelectionChanged;
                //richTextBox.DragEnter += RichTextBoxOnDragEnter;
                //richTextBox.DragDrop += RichTextBoxOnDragDrop;
                StreamReader sr = new StreamReader(fileName, Encoding.Default);
                string str = sr.ReadToEnd();
                if (Path.GetExtension(newDoc.PathDocument) == ".txt")
                {
                    richTextBox.Text = str;
                }
                else
                {
                    richTextBox.Rtf = str;
                }
                sr.Close();
                page.Controls.Add(richTextBox);
                tabControl.TabPages.Add(page);
                tabControl.SelectedTab = page;
                tabControl.SelectedTab.Tag = newDoc;
                _documents.Add(newDoc);
            }
        }
        #endregion
    }
}

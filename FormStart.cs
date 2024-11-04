using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.VisualBasic;

namespace MiniWord_NgoNgocTrungAnh
{
    public partial class FormStart : Form
    {
        private Dictionary<string, List<Control>> menuControls;
        private int currentZoom = 100; // Kích thước zoom ban đầu là 100%

        public FormStart()
        {
            InitializeComponent();
            InitializeMenuControls(); // Add this line to your existing constructor
            FixImageSize();
            this.richTextBox1.Focus();
            // Nạp danh sách font vào comboBoxFontFamily
            foreach (FontFamily font in FontFamily.Families)
            {
                comboBoxFontFamily.Items.Add(font.Name);
            }

            // Đặt font mặc định ban đầu, nếu muốn
            comboBoxFontFamily.SelectedIndex = comboBoxFontFamily.Items.IndexOf(richTextBox1.Font.FontFamily.Name);
         
            this.numericUpDown1.ValueChanged += new System.EventHandler(this.numericUpDown1_ValueChanged);

        }

        #region MenuTrip button Pannel Control
        private void InitializeMenuControls()
        {
            // Initialize the dictionary
            menuControls = new Dictionary<string, List<Control>>();

            // Add controls for Home menu
            menuControls["Home"] = new List<Control>
            {
                btnPaste,
                btnCopy,
                btnCut,
                btnRedo,
                btnUndo
            };

            // Add controls for Insert menu (add your Insert controls here)
            menuControls["Insert"] = new List<Control>();
           
            // Add controls for Layout menu (add your Layout controls here)
            menuControls["Layout"] = new List<Control>();
            menuControls["Layout"] = new List<Control>
            {
                btnZoomIn,
                btnZoomOut,
                btnWordBackGroundColor,
                btnWordColor,

            };
            // Hide all controls initially except Home
            ShowMenuControls("Home");

            // Add click handlers for menu items
            homeToolStripMenuItem.Click += (s, e) => ShowMenuControls("Home");
            insertToolStripMenuItem.Click += (s, e) => ShowMenuControls("Insert");
            layoutToolStripMenuItem.Click += (s, e) => ShowMenuControls("Layout");
        }
        private void ShowMenuControls(string menuName)
        {
            // Hide all controls first
            foreach (var controls in menuControls.Values)
            {
                foreach (var control in controls)
                {
                    control.Visible = false;
                }
            }

            // Show controls for selected menu
            if (menuControls.ContainsKey(menuName))
            {
                foreach (var control in menuControls[menuName])
                {
                    control.Visible = true;
                }
            }
        }

        
        #endregion

        #region Image on btn Fix

        public Image ResizeImage(Image img, int width, int height)
        {
            Bitmap bmp = new Bitmap(width, height);
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                g.DrawImage(img, 0, 0, width, height);
            }
            return bmp;
        }
        public void FixImageSize()
        {
            //home 
            btnPaste.Image = ResizeImage(global::MiniWord_NgoNgocTrungAnh.Properties.Resources.paste,
                                btnPaste.Width - 10,
                                btnPaste.Height - 20);
            btnCopy.Image = ResizeImage(global::MiniWord_NgoNgocTrungAnh.Properties.Resources.copy,
                                btnCopy.Width - 10,
                                btnCopy.Height - 20);
            btnRedo.Image = ResizeImage(global::MiniWord_NgoNgocTrungAnh.Properties.Resources.redo,
                                btnRedo.Width - 10,
                                btnRedo.Height - 20);
            btnUndo.Image = ResizeImage(global::MiniWord_NgoNgocTrungAnh.Properties.Resources.undo,
                                btnUndo.Width - 10,
                                btnUndo.Height - 20);
            btnCut.Image = ResizeImage(global::MiniWord_NgoNgocTrungAnh.Properties.Resources.cut,
                                btnCut.Width - 10,
                                btnCut.Height - 20);
            //layout
            btnZoomIn.Image = ResizeImage(global::MiniWord_NgoNgocTrungAnh.Properties.Resources.zoomin,
                               btnCut.Width - 10,
                               btnCut.Height - 20);
            btnZoomOut.Image = ResizeImage(global::MiniWord_NgoNgocTrungAnh.Properties.Resources.zoomout,
                               btnCut.Width - 10,
                               btnCut.Height - 20);
            btnWordColor.Image = ResizeImage(global::MiniWord_NgoNgocTrungAnh.Properties.Resources.wordcolor,
                               btnCut.Width - 10,
                               btnCut.Height - 20);
            btnWordBackGroundColor.Image = ResizeImage(global::MiniWord_NgoNgocTrungAnh.Properties.Resources.wordbgcolor,
                               btnCut.Width - 10,
                               btnCut.Height - 20);
        }
        #endregion

        #region Button on Home 
        private void btnPaste_Click(object sender, EventArgs e)
        {
            // Check if there is text in the clipboard
            if (Clipboard.ContainsText())
            {
                // Get the text from the clipboard
                string clipboardText = Clipboard.GetText();

                // Insert the text at the current cursor position in richTextBox1
                int selectionStart = richTextBox1.SelectionStart;
                richTextBox1.Text = richTextBox1.Text.Insert(selectionStart, clipboardText);

                // Move the cursor after the pasted text
                richTextBox1.SelectionStart = selectionStart + clipboardText.Length;
            }
            else
            {
                MessageBox.Show("Clipboard does not contain text.", "Paste Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
        private void btnCopy_Click(object sender, EventArgs e)
        {
            // Copy the selected text to the clipboard
            if (richTextBox1.SelectionLength > 0)
            {
                Clipboard.SetText(richTextBox1.SelectedText);
            }
            else
            {
                MessageBox.Show("No text selected to copy.", "Copy Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
     

        private void btnCut_Click(object sender, EventArgs e)
        {
            // Cut the selected text: copy to clipboard and remove from RichTextBox
            if (richTextBox1.SelectionLength > 0)
            {
                Clipboard.SetText(richTextBox1.SelectedText);
                richTextBox1.SelectedText = ""; // Remove the selected text
            }
            else
            {
                MessageBox.Show("No text selected to cut.", "Cut Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void btnUndo_Click(object sender, EventArgs e)
        {
            // Undo the last action
            if (richTextBox1.CanUndo)
            {
                richTextBox1.Undo(); // Undo the last action
            }
        }

        private void btnRedo_Click(object sender, EventArgs e)
        {
            // Redo the last undone action
            if (richTextBox1.CanRedo)
            {
                richTextBox1.Redo(); // Redo the last undone action
            }
        }
        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {
            // Lấy giá trị cỡ chữ mới từ numericUpDown1
            float newSize = (float)numericUpDown1.Value;

            // Kiểm tra nếu có đoạn văn bản được chọn
            if (richTextBox1.SelectionLength > 0 && richTextBox1.SelectionFont != null)
            {
                // Lấy font hiện tại của đoạn văn bản được chọn
                Font currentFont = richTextBox1.SelectionFont;
                Font newFont = new Font(currentFont.FontFamily, newSize, currentFont.Style);

                // Áp dụng font mới cho đoạn văn bản được chọn
                richTextBox1.SelectionFont = newFont;
            }
            else
            {
                // Nếu không có đoạn văn bản nào được chọn, áp dụng cỡ chữ mới cho văn bản sẽ được nhập sau đó
                richTextBox1.Font = new Font(richTextBox1.Font.FontFamily, newSize, richTextBox1.Font.Style);
            }
        }


        private void richTextBox1_SelectionChanged(object sender, EventArgs e)
        {
            if (richTextBox1.SelectionFont != null)
            {
                // Cập nhật kích thước chữ vào numericUpDown1
                numericUpDown1.Value = (decimal)richTextBox1.SelectionFont.Size;

                // Cập nhật font hiện tại trong comboBoxFontFamily
                //comboBoxFontFamily.SelectedItem = richTextBox1.SelectionFont.FontFamily.Name;
            }
        }

        private void checkBoxBold_CheckedChanged(object sender, EventArgs e)
        {
            UpdateFontStyle();
        }


        private void checkBoxItalic_CheckedChanged(object sender, EventArgs e)
        {
            UpdateFontStyle();
        }


        private void checkBoxUnderline_CheckedChanged(object sender, EventArgs e)
        {
            UpdateFontStyle();
        }
        private void UpdateFontStyle()
        {
            FontStyle style = FontStyle.Regular;

            if (checkBoxBold.Checked)
                style |= FontStyle.Bold;
            if (checkBoxItalic.Checked)
                style |= FontStyle.Italic;
            if (checkBoxUnderline.Checked)
                style |= FontStyle.Underline;

            if (richTextBox1.SelectionLength > 0)
            {
                richTextBox1.SelectionFont = new Font(richTextBox1.SelectionFont.FontFamily, richTextBox1.SelectionFont.Size, style);
            }
            else
            {
                richTextBox1.Font = new Font(richTextBox1.Font.FontFamily, richTextBox1.Font.Size, style);
            }
        }


        private void comboBoxFontFamily_SelectedIndexChanged(object sender, EventArgs e)
        {
            string selectedFont = comboBoxFontFamily.SelectedItem.ToString();

            if (richTextBox1.SelectionLength > 0)
            {
                richTextBox1.SelectionFont = new Font(selectedFont, richTextBox1.SelectionFont.Size, richTextBox1.SelectionFont.Style);
            }
            else
            {
                richTextBox1.Font = new Font(selectedFont, richTextBox1.Font.Size, richTextBox1.Font.Style);
            }
        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
            string searchText = searchTextBox.Text;
            if (string.IsNullOrWhiteSpace(searchText))
                return;

            // Tìm vị trí đầu tiên của văn bản cần tìm
            int startIndex = 0;
            while ((startIndex = richTextBox1.Find(searchText, startIndex, RichTextBoxFinds.None)) != -1)
            {
                // Đánh dấu văn bản tìm thấy
                richTextBox1.SelectionStart = startIndex;
                richTextBox1.SelectionLength = searchText.Length;
                richTextBox1.SelectionBackColor = Color.Yellow; // Màu nền cho từ tìm thấy
                startIndex += searchText.Length; // Tiến tới vị trí tiếp theo để tìm kiếm
            }
        }


        private void btnReplace_Click(object sender, EventArgs e)
        {
            // Hiển thị hộp thoại nhập từ cần thay thế
            string replaceText = Interaction.InputBox("Nhập từ cần thay thế:", "Thay thế văn bản", "");

            if (string.IsNullOrWhiteSpace(replaceText))
                return; // Nếu không có từ nào để thay thế, thoát khỏi phương thức

            // Giả sử bạn có một TextBox để nhập từ cần tìm
            string searchText = searchTextBox.Text;

            if (string.IsNullOrWhiteSpace(searchText))
                return; // Nếu không có từ tìm kiếm, thoát khỏi phương thức

            // Tìm kiếm và thay thế
            int startIndex = 0;
            while ((startIndex = richTextBox1.Find(searchText, startIndex, RichTextBoxFinds.None)) != -1)
            {
                // Thay thế văn bản
                richTextBox1.SelectionStart = startIndex;
                richTextBox1.SelectionLength = searchText.Length;
                richTextBox1.SelectedText = replaceText; // Thay thế văn bản tìm thấy
                startIndex += replaceText.Length; // Tiến tới vị trí tiếp theo để tìm kiếm
            }
        }

        #endregion

        #region Button on Layout
        private void btnWordColor_Click(object sender, EventArgs e)
        {
            // Mở hộp thoại chọn màu
            using (ColorDialog colorDialog = new ColorDialog())
            {
                if (colorDialog.ShowDialog() == DialogResult.OK)
                {
                    // Đặt màu chữ cho văn bản đã chọn trong RichTextBox
                    richTextBox1.SelectionColor = colorDialog.Color;
                }
            }
        }

        private void btnZoomOut_Click(object sender, EventArgs e)
        {
            // Giảm kích thước zoom
            if (currentZoom > 10) // Đảm bảo kích thước không nhỏ hơn 10%
            {
                currentZoom -= 10;
                UpdateZoom();
            }
        }

        private void btnZoomIn_Click(object sender, EventArgs e)
        {
            // Tăng kích thước zoom
            currentZoom += 10;
            UpdateZoom();
        }

        // Cập nhật kích thước cho RichTextBox
        private void UpdateZoom()
        {
            richTextBox1.ZoomFactor = currentZoom / 100f; // Chuyển đổi sang tỷ lệ
        }


        private void btnWordBackGroundColor_Click(object sender, EventArgs e)
        {
            // Mở hộp thoại chọn màu nền
            using (ColorDialog colorDialog = new ColorDialog())
            {
                if (colorDialog.ShowDialog() == DialogResult.OK)
                {
                    // Đặt màu nền cho văn bản đã chọn trong RichTextBox
                    richTextBox1.SelectionBackColor = colorDialog.Color;
                }
            }
        }

        #endregion

        #region Rich Textbox


        #endregion
        #region
        #endregion
        #region File menutrip
        private void newToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // Xóa nội dung của RichTextBox để tạo một tài liệu mới
            richTextBox1.Clear();
            this.Text = "New Document"; // Cập nhật tiêu đề form
        }

        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*";
                openFileDialog.Title = "Open a Document";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    // Đọc nội dung từ file và hiển thị trong RichTextBox
                    richTextBox1.Text = System.IO.File.ReadAllText(openFileDialog.FileName);
                    this.Text = openFileDialog.FileName; // Cập nhật tiêu đề form với tên file
                }
            }
        }

        private void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // Lưu tài liệu hiện tại
            int tempZoom = currentZoom; // Lưu lại kích thước zoom hiện tại
            currentZoom = 100; // Đặt lại kích thước zoom về 100%
            UpdateZoom(); // Cập nhật lại zoom

            using (SaveFileDialog saveFileDialog = new SaveFileDialog())
            {
                saveFileDialog.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*";
                saveFileDialog.Title = "Save a Document";

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    // Ghi nội dung của RichTextBox vào file
                    System.IO.File.WriteAllText(saveFileDialog.FileName, richTextBox1.Text);
                    this.Text = saveFileDialog.FileName; // Cập nhật tiêu đề form với tên file
                }
            }

            currentZoom = tempZoom; // Khôi phục kích thước zoom ban đầu
            UpdateZoom(); // Cập nhật lại zoom
        }

        private void saveASToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // Lưu tài liệu hiện tại với tên khác
            int tempZoom = currentZoom; // Lưu lại kích thước zoom hiện tại
            currentZoom = 100; // Đặt lại kích thước zoom về 100%
            UpdateZoom(); // Cập nhật lại zoom

            using (SaveFileDialog saveFileDialog = new SaveFileDialog())
            {
                saveFileDialog.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*";
                saveFileDialog.Title = "Save As";

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    // Ghi nội dung của RichTextBox vào file
                    System.IO.File.WriteAllText(saveFileDialog.FileName, richTextBox1.Text);
                    this.Text = saveFileDialog.FileName; // Cập nhật tiêu đề form với tên file
                }
            }

            currentZoom = tempZoom; // Khôi phục kích thước zoom ban đầu
            UpdateZoom(); // Cập nhật lại zoom
        }


        private void closeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // Đóng tài liệu hiện tại, hỏi người dùng nếu có nội dung chưa lưu
            if (!string.IsNullOrWhiteSpace(richTextBox1.Text))
            {
                var result = MessageBox.Show("Do you want to save changes?", "Confirm", MessageBoxButtons.YesNoCancel);
                if (result == DialogResult.Yes)
                {
                    saveToolStripMenuItem_Click(sender, e); // Gọi hàm lưu tài liệu
                }
                else if (result == DialogResult.Cancel)
                {
                    return; // Hủy bỏ hành động đóng
                }
            }
            richTextBox1.Clear(); // Xóa nội dung RichTextBox
            this.Text = "Untitled"; // Đặt lại tiêu đề form
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // Đóng ứng dụng
            this.Close(); // Gọi phương thức đóng form
        }



        #endregion

        

    }

}

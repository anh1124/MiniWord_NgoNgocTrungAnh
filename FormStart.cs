using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.VisualBasic;
using MiniWord_NgoNgocTrungAnh.Properties;

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
            richTextBox1.Focus();
            // Nạp danh sách font vào comboBoxFontFamily
            foreach (var font in FontFamily.Families) comboBoxFontFamily.Items.Add(font.Name);

            // Đặt font mặc định ban đầu, nếu muốn
            comboBoxFontFamily.SelectedIndex = comboBoxFontFamily.Items.IndexOf(richTextBox1.Font.FontFamily.Name);

            numericUpDown1.ValueChanged += numericUpDown1_ValueChanged;

            richTextBox1.TextChanged += RichTextBox1_TextChanged;
            richTextBox1.AcceptsTab = true;
            richTextBox1.KeyDown += RichTextBox1_KeyDown;
            btnAddTable.Click += btnAddTable_Click;
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
            menuControls["Insert"] = new List<Control>
            {
                btnAddTable
            };


            // Add controls for Layout menu (add your Layout controls here)
            menuControls["Layout"] = new List<Control>();
            menuControls["Layout"] = new List<Control>
            {
                btnZoomIn,
                btnZoomOut,
                btnWordBackGroundColor,
                btnWordColor
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
            foreach (var control in controls)
                control.Visible = false;

            // Show controls for selected menu
            if (menuControls.ContainsKey(menuName))
                foreach (var control in menuControls[menuName])
                    control.Visible = true;
        }

        #endregion

        #region Image on btn Fix

        public Image ResizeImage(Image img, int width, int height)
        {
            var bmp = new Bitmap(width, height);
            using (var g = Graphics.FromImage(bmp))
            {
                g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                g.DrawImage(img, 0, 0, width, height);
            }

            return bmp;
        }

        public void FixImageSize()
        {
            //home 
            btnPaste.Image = ResizeImage(Resources.paste,
                btnPaste.Width - 10,
                btnPaste.Height - 20);
            btnCopy.Image = ResizeImage(Resources.copy,
                btnCopy.Width - 10,
                btnCopy.Height - 20);
            btnRedo.Image = ResizeImage(Resources.redo,
                btnRedo.Width - 10,
                btnRedo.Height - 20);
            btnUndo.Image = ResizeImage(Resources.undo,
                btnUndo.Width - 10,
                btnUndo.Height - 20);
            btnCut.Image = ResizeImage(Resources.cut,
                btnCut.Width - 10,
                btnCut.Height - 20);
            //layout
            btnZoomIn.Image = ResizeImage(Resources.zoomin,
                btnCut.Width - 10,
                btnCut.Height - 20);
            btnZoomOut.Image = ResizeImage(Resources.zoomout,
                btnCut.Width - 10,
                btnCut.Height - 20);
            btnWordColor.Image = ResizeImage(Resources.wordcolor,
                btnCut.Width - 10,
                btnCut.Height - 20);
            btnWordBackGroundColor.Image = ResizeImage(Resources.wordbgcolor,
                btnCut.Width - 10,
                btnCut.Height - 20);
            //insert
            /*btnAddTable.Image =  ResizeImage(global::MiniWord_NgoNgocTrungAnh.Properties.Resources.table,
                               btnCut.Width - 10,
                               btnCut.Height - 20);*/
            btnAddTable.Image = ResizeImage(Resources.table,
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
                var clipboardText = Clipboard.GetText();

                // Insert the text at the current cursor position in richTextBox1
                var selectionStart = richTextBox1.SelectionStart;
                richTextBox1.Text = richTextBox1.Text.Insert(selectionStart, clipboardText);

                // Move the cursor after the pasted text
                richTextBox1.SelectionStart = selectionStart + clipboardText.Length;
            }
            else
            {
                MessageBox.Show("Clipboard does not contain text.", "Paste Error", MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }
        }

        private void btnCopy_Click(object sender, EventArgs e)
        {
            // Copy the selected text to the clipboard
            if (richTextBox1.SelectionLength > 0)
                Clipboard.SetText(richTextBox1.SelectedText);
            else
                MessageBox.Show("No text selected to copy.", "Copy Error", MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
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
            if (richTextBox1.CanUndo) richTextBox1.Undo(); // Undo the last action
        }

        private void btnRedo_Click(object sender, EventArgs e)
        {
            // Redo the last undone action
            if (richTextBox1.CanRedo) richTextBox1.Redo(); // Redo the last undone action
        }

        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {
            // Lấy giá trị cỡ chữ mới từ numericUpDown1
            var newSize = (float)numericUpDown1.Value;

            // Kiểm tra nếu có đoạn văn bản được chọn
            if (richTextBox1.SelectionLength > 0 && richTextBox1.SelectionFont != null)
            {
                // Lấy font hiện tại của đoạn văn bản được chọn
                var currentFont = richTextBox1.SelectionFont;
                var newFont = new Font(currentFont.FontFamily, newSize, currentFont.Style);

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
                // Cập nhật kích thước chữ vào numericUpDown1
                numericUpDown1.Value = (decimal)richTextBox1.SelectionFont.Size;
            // Cập nhật font hiện tại trong comboBoxFontFamily
            //comboBoxFontFamily.SelectedItem = richTextBox1.SelectionFont.FontFamily.Name;
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
            var style = FontStyle.Regular;

            if (checkBoxBold.Checked)
                style |= FontStyle.Bold;
            if (checkBoxItalic.Checked)
                style |= FontStyle.Italic;
            if (checkBoxUnderline.Checked)
                style |= FontStyle.Underline;

            if (richTextBox1.SelectionLength > 0)
                richTextBox1.SelectionFont = new Font(richTextBox1.SelectionFont.FontFamily,
                    richTextBox1.SelectionFont.Size, style);
            else
                richTextBox1.Font = new Font(richTextBox1.Font.FontFamily, richTextBox1.Font.Size, style);
        }


        private void comboBoxFontFamily_SelectedIndexChanged(object sender, EventArgs e)
        {
            var selectedFont = comboBoxFontFamily.SelectedItem.ToString();

            if (richTextBox1.SelectionLength > 0)
                richTextBox1.SelectionFont = new Font(selectedFont, richTextBox1.SelectionFont.Size,
                    richTextBox1.SelectionFont.Style);
            else
                richTextBox1.Font = new Font(selectedFont, richTextBox1.Font.Size, richTextBox1.Font.Style);
        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
            var searchText = searchTextBox.Text;
            if (string.IsNullOrWhiteSpace(searchText))
                return;

            // Tìm vị trí đầu tiên của văn bản cần tìm
            var startIndex = 0;
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
            var replaceText = Interaction.InputBox("Nhập từ cần thay thế:", "Thay thế văn bản");

            if (string.IsNullOrWhiteSpace(replaceText))
                return; // Nếu không có từ nào để thay thế, thoát khỏi phương thức

            // Giả sử bạn có một TextBox để nhập từ cần tìm
            var searchText = searchTextBox.Text;

            if (string.IsNullOrWhiteSpace(searchText))
                return; // Nếu không có từ tìm kiếm, thoát khỏi phương thức

            // Tìm kiếm và thay thế
            var startIndex = 0;
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
            using (var colorDialog = new ColorDialog())
            {
                if (colorDialog.ShowDialog() == DialogResult.OK)
                    // Đặt màu chữ cho văn bản đã chọn trong RichTextBox
                    richTextBox1.SelectionColor = colorDialog.Color;
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
            using (var colorDialog = new ColorDialog())
            {
                if (colorDialog.ShowDialog() == DialogResult.OK)
                    // Đặt màu nền cho văn bản đã chọn trong RichTextBox
                    richTextBox1.SelectionBackColor = colorDialog.Color;
            }
        }

        #endregion

        #region File menutrip

        private void newToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // Kiểm tra xem nội dung của RichTextBox có thay đổi chưa
            if (!string.IsNullOrEmpty(richTextBox1.Text))
            {
                // Hỏi người dùng nếu họ có muốn lưu tài liệu trước khi tạo tài liệu mới
                var result = MessageBox.Show("Do you want to save the current document?", 
                    "Save Document", 
                    MessageBoxButtons.YesNoCancel, 
                    MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    // Nếu người dùng chọn "Yes", gọi hàm lưu (bạn có thể tạo một phương thức lưu)
                    SaveDocument();
                }
                else if (result == DialogResult.No)
                {
                    // Nếu người dùng chọn "No", bỏ qua và xóa nội dung
                    richTextBox1.Clear();
                    Text = "New Document"; // Cập nhật tiêu đề form
                }
                // Nếu chọn "Cancel", không làm gì cả (giữ nguyên tài liệu hiện tại)
            }
            else
            {
                // Nếu không có nội dung trong RichTextBox, chỉ cần xóa
                richTextBox1.Clear();
                Text = "New Document"; // Cập nhật tiêu đề form
            }
        }

        private void SaveDocument()
        {
            // Đoạn mã lưu tài liệu, ví dụ như lưu vào file
            // Sử dụng SaveFileDialog để chọn đường dẫn và lưu tệp
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*";
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                // Lưu nội dung của RichTextBox vào file
                System.IO.File.WriteAllText(saveFileDialog.FileName, richTextBox1.Text);
                MessageBox.Show("Document saved successfully.");
            }
        }


        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using (var openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*";
                openFileDialog.Title = "Open a Document";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    // Đọc nội dung từ file và hiển thị trong RichTextBox
                    richTextBox1.Text = File.ReadAllText(openFileDialog.FileName);
                    Text = openFileDialog.FileName; // Cập nhật tiêu đề form với tên file
                }
            }
        }

        private void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // Lưu tài liệu hiện tại
            var tempZoom = currentZoom; // Lưu lại kích thước zoom hiện tại
            currentZoom = 100; // Đặt lại kích thước zoom về 100%
            UpdateZoom(); // Cập nhật lại zoom

            using (var saveFileDialog = new SaveFileDialog())
            {
                saveFileDialog.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*";
                saveFileDialog.Title = "Save a Document";

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    // Ghi nội dung của RichTextBox vào file
                    File.WriteAllText(saveFileDialog.FileName, richTextBox1.Text);
                    Text = saveFileDialog.FileName; // Cập nhật tiêu đề form với tên file
                }
            }

            currentZoom = tempZoom; // Khôi phục kích thước zoom ban đầu
            UpdateZoom(); // Cập nhật lại zoom
        }

        private void saveASToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // Lưu tài liệu hiện tại với tên khác
            var tempZoom = currentZoom; // Lưu lại kích thước zoom hiện tại
            currentZoom = 100; // Đặt lại kích thước zoom về 100%
            UpdateZoom(); // Cập nhật lại zoom

            using (var saveFileDialog = new SaveFileDialog())
            {
                saveFileDialog.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*";
                saveFileDialog.Title = "Save As";

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    // Ghi nội dung của RichTextBox vào file
                    File.WriteAllText(saveFileDialog.FileName, richTextBox1.Text);
                    Text = saveFileDialog.FileName; // Cập nhật tiêu đề form với tên file
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
                    saveToolStripMenuItem_Click(sender, e); // Gọi hàm lưu tài liệu
                else if (result == DialogResult.Cancel) return; // Hủy bỏ hành động đóng
            }

            richTextBox1.Clear(); // Xóa nội dung RichTextBox
            Text = "Untitled"; // Đặt lại tiêu đề form
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // Đóng ứng dụng
            Close(); // Gọi phương thức đóng form
        }

        #endregion

        #region Other feather

        private void RichTextBox1_TextChanged(object sender, EventArgs e)
        {
            // Kiểm tra xem có emoji mới được thêm vào không
            if (richTextBox1.Text.Length > 0)
            {
                richTextBox1.SelectionStart = richTextBox1.Text.Length; // Di chuyển con trỏ về cuối
                richTextBox1.ScrollToCaret(); // Cuộn đến vị trí con trỏ
            }
        }

        private void iconBtn_Click(object sender, EventArgs e)
        {
            using (var iconPicker = new IconPickerForm())
            {
                // Hiển thị form chọn icon gần vị trí nút
                iconPicker.StartPosition = FormStartPosition.Manual;
                var location = iconBtn.PointToScreen(Point.Empty);
                iconPicker.Location = new Point(location.X, location.Y + iconBtn.Height);

                if (iconPicker.ShowDialog() == DialogResult.OK)
                {
                    // Lấy vị trí con trỏ hiện tại
                    var cursorPosition = richTextBox1.SelectionStart;

                    // Chèn emoji vào vị trí con trỏ
                    richTextBox1.Text = richTextBox1.Text.Insert(cursorPosition, iconPicker.SelectedEmoji);

                    // Di chuyển con trỏ về sau emoji vừa chèn
                    richTextBox1.SelectionStart = cursorPosition + iconPicker.SelectedEmoji.Length;
                    richTextBox1.Focus();
                }
            }
        }


        private void RichTextBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Tab)
            {
                e.Handled = true;

                var currentPosition = richTextBox1.SelectionStart;
                var tabString =
                    new string(' ', (int)numericUpDownTabSpace.Value); // Tạo chuỗi khoảng trắng theo tabSpace

                if (richTextBox1.SelectionLength > 0)
                {
                    var selectedText = richTextBox1.SelectedText;
                    var lines = selectedText.Split(new[] { Environment.NewLine }, StringSplitOptions.None);

                    if (e.Shift)
                    {
                        // Giảm indent (xóa khoảng trắng đầu dòng theo tabSpace)
                        for (var i = 0; i < lines.Length; i++)
                            if (lines[i].StartsWith(tabString))
                                lines[i] = lines[i].Substring(tabString.Length);
                    }
                    else
                    {
                        // Tăng indent (thêm khoảng trắng đầu dòng theo tabSpace)
                        for (var i = 0; i < lines.Length; i++) lines[i] = tabString + lines[i];
                    }

                    var newText = string.Join(Environment.NewLine, lines);
                    richTextBox1.SelectedText = newText;
                }
                else
                {
                    // Thêm khoảng trắng tại vị trí con trỏ
                    richTextBox1.SelectedText = tabString;
                }

                // Đặt lại vị trí con trỏ
                richTextBox1.SelectionStart = currentPosition + tabString.Length;
                richTextBox1.Focus();
            }
        }

        #endregion

        #region Addtable

        private void btnAddTable_Click(object sender, EventArgs e)
        {
            var tableCreator = new TableCreator(richTextBox1);
            tableCreator.ShowTableDialog();
        }

        public class TableProperties
        {
            public int Rows { get; set; }
            public int Columns { get; set; }
            public List<int> ColumnWidths { get; set; }
            public Color BorderColor { get; set; }
            public Color BackgroundColor { get; set; }
            public int CellPadding { get; set; }
            public ContentAlignment DefaultAlignment { get; set; }
        }

        public class TableCell
        {
            public string Text { get; set; }
            public ContentAlignment Alignment { get; set; }
            public Color BackgroundColor { get; set; }
            public Font Font { get; set; }
        }

        public class TableCreator
        {
            private readonly RichTextBox richTextBox;
            private TableProperties properties;
            private TableCell[,] cells;
            private bool isInTable;

            public TableCreator(RichTextBox richTextBox)
            {
                this.richTextBox = richTextBox;
                this.richTextBox.KeyDown += RichTextBox_KeyDown;
            }

            private void RichTextBox_KeyDown(object sender, KeyEventArgs e)
            {
                // Kiểm tra xem con trỏ có đang ở trong bảng không
                isInTable = IsCaretInTable();

                if (isInTable)
                    if (e.KeyCode == Keys.Enter || e.KeyCode == Keys.Tab)
                        // Nếu đang ở cuối bảng và nhấn Enter hoặc Tab
                        if (IsAtTableEnd())
                        {
                            e.Handled = true;
                            // Di chuyển con trỏ ra ngoài bảng
                            richTextBox.SelectionStart = GetTableEndPosition() + 1;
                            richTextBox.SelectedText = Environment.NewLine;
                        }
            }

            private bool IsCaretInTable()
            {
                // Kiểm tra xem vị trí con trỏ hiện tại có nằm trong bảng không
                var rtf = richTextBox.Rtf;
                var caretPosition = richTextBox.SelectionStart;

                // Tìm các marker của bảng trong RTF
                var tableStart = rtf.LastIndexOf(@"\trowd", caretPosition);
                var tableEnd = rtf.IndexOf(@"\row", caretPosition);

                return tableStart != -1 && tableEnd != -1 && tableStart < caretPosition && caretPosition < tableEnd;
            }

            private int GetTableEndPosition()
            {
                var rtf = richTextBox.Rtf;
                var caretPosition = richTextBox.SelectionStart;
                var tableEnd = rtf.IndexOf(@"\row", caretPosition);

                if (tableEnd != -1)
                    // Tìm vị trí thực tế trong văn bản
                    return richTextBox.GetFirstCharIndexFromLine(
                        richTextBox.GetLineFromCharIndex(caretPosition) + 1);
                return caretPosition;
            }

            private bool IsAtTableEnd()
            {
                var rtf = richTextBox.Rtf;
                var caretPosition = richTextBox.SelectionStart;
                var tableEnd = rtf.IndexOf(@"\row", caretPosition);

                // Kiểm tra xem con trỏ có ở gần cuối bảng không
                return tableEnd != -1 && tableEnd - caretPosition < 10;
            }

            public void ShowTableDialog()
            {
                using (var dialog = new Form())
                {
                    dialog.Text = "Create Table";
                    dialog.Size = new Size(450, 350);
                    dialog.StartPosition = FormStartPosition.CenterParent;
                    dialog.FormBorderStyle = FormBorderStyle.FixedDialog;

                    // Tạo controls với kích thước lớn hơn và font rõ ràng hơn
                    var font = new Font("Segoe UI", 10F);

                    var lblRows = new Label
                    {
                        Text = "Rows:",
                        Location = new Point(30, 30),
                        Size = new Size(100, 25),
                        Font = font
                    };

                    var lblCols = new Label
                    {
                        Text = "Columns:",
                        Location = new Point(30, 70),
                        Size = new Size(100, 25),
                        Font = font
                    };

                    var numRows = new NumericUpDown
                    {
                        Location = new Point(140, 30),
                        Size = new Size(80, 25),
                        Minimum = 1,
                        Maximum = 20,
                        Value = 3,
                        Font = font
                    };

                    var numCols = new NumericUpDown
                    {
                        Location = new Point(140, 70),
                        Size = new Size(80, 25),
                        Minimum = 1,
                        Maximum = 10,
                        Value = 3,
                        Font = font
                    };

                    var lblPadding = new Label
                    {
                        Text = "Cell Padding:",
                        Location = new Point(30, 110),
                        Size = new Size(100, 25),
                        Font = font
                    };

                    var numPadding = new NumericUpDown
                    {
                        Location = new Point(140, 110),
                        Size = new Size(80, 25),
                        Minimum = 1,
                        Maximum = 20,
                        Value = 5, // Tăng padding mặc định
                        Font = font
                    };

                    var lblCellWidth = new Label
                    {
                        Text = "Cell Width:",
                        Location = new Point(30, 150),
                        Size = new Size(100, 25),
                        Font = font
                    };

                    var numCellWidth = new NumericUpDown
                    {
                        Location = new Point(140, 150),
                        Size = new Size(80, 25),
                        Minimum = 50,
                        Maximum = 500,
                        Value = 150, // Tăng độ rộng mặc định của ô
                        Font = font
                    };

                    var btnBorderColor = new Button
                    {
                        Text = "Border Color",
                        Location = new Point(30, 190),
                        Size = new Size(120, 35),
                        Font = font
                    };

                    var btnBgColor = new Button
                    {
                        Text = "Background Color",
                        Location = new Point(160, 190),
                        Size = new Size(120, 35),
                        Font = font
                    };

                    var borderColor = Color.Black;
                    var bgColor = Color.White;

                    btnBorderColor.Click += (s, e) =>
                    {
                        using (var colorDialog = new ColorDialog())
                        {
                            if (colorDialog.ShowDialog() == DialogResult.OK)
                            {
                                borderColor = colorDialog.Color;
                                btnBorderColor.BackColor = borderColor;
                            }
                        }
                    };

                    btnBgColor.Click += (s, e) =>
                    {
                        using (var colorDialog = new ColorDialog())
                        {
                            if (colorDialog.ShowDialog() == DialogResult.OK)
                            {
                                bgColor = colorDialog.Color;
                                btnBgColor.BackColor = bgColor;
                            }
                        }
                    };

                    var btnOk = new Button
                    {
                        Text = "OK",
                        DialogResult = DialogResult.OK,
                        Location = new Point(100, 250),
                        Size = new Size(100, 35),
                        Font = font
                    };

                    var btnCancel = new Button
                    {
                        Text = "Cancel",
                        DialogResult = DialogResult.Cancel,
                        Location = new Point(220, 250),
                        Size = new Size(100, 35),
                        Font = font
                    };

                    dialog.Controls.AddRange(new Control[]
                    {
                        lblRows, lblCols, numRows, numCols,
                        lblPadding, numPadding,
                        lblCellWidth, numCellWidth,
                        btnBorderColor, btnBgColor,
                        btnOk, btnCancel
                    });

                    if (dialog.ShowDialog() == DialogResult.OK)
                    {
                        properties = new TableProperties
                        {
                            Rows = (int)numRows.Value,
                            Columns = (int)numCols.Value,
                            BorderColor = borderColor,
                            BackgroundColor = bgColor,
                            CellPadding = (int)numPadding.Value,
                            DefaultAlignment = ContentAlignment.MiddleLeft,
                            ColumnWidths = Enumerable.Repeat((int)numCellWidth.Value, (int)numCols.Value).ToList()
                        };

                        InitializeCells();
                        InsertTable();
                    }
                }
            }

            private void InitializeCells()
            {
                cells = new TableCell[properties.Rows, properties.Columns];
                for (var i = 0; i < properties.Rows; i++)
                for (var j = 0; j < properties.Columns; j++)
                    cells[i, j] = new TableCell
                    {
                        Text = "",
                        Alignment = properties.DefaultAlignment,
                        BackgroundColor = properties.BackgroundColor,
                        Font = richTextBox.Font
                    };
            }

            private void InsertTable()
            {
                var startPosition = richTextBox.SelectionStart;
                var rtf = new StringBuilder();
                rtf.Append(@"{\rtf1\ansi\deff0");

                // Tăng khoảng cách giữa các ô
                rtf.Append(@"\trowd\trgaph" + properties.CellPadding * 30); // Tăng hệ số padding
                rtf.Append(
                    @"\trbrdrl\brdrs\brdrw20\trbrdrt\brdrs\brdrw20\trbrdrr\brdrs\brdrw20\trbrdrb\brdrs\brdrw20"); // Thêm đường viền đậm cho bảng

                // Thiết lập độ rộng và style cho các ô
                var totalWidth = 0;
                for (var j = 0; j < properties.Columns; j++)
                {
                    totalWidth += properties.ColumnWidths[j] * 15; // Tăng hệ số nhân để có ô rộng hơn
                    rtf.Append(
                        @"\clbrdrl\brdrw10\brdrs\clbrdrt\brdrw10\brdrs\clbrdrr\brdrw10\brdrs\clbrdrb\brdrw10\brdrs"); // Thêm đường viền cho từng ô
                    rtf.Append(@"\cellx" + totalWidth);
                }

                // Thêm nội dung các ô
                for (var i = 0; i < properties.Rows; i++)
                {
                    for (var j = 0; j < properties.Columns; j++)
                        rtf.Append(@"\intbl\cell"); // Thêm ô trống với định dạng bảng
                    rtf.Append(@"\row\trowd\trgaph" + properties.CellPadding * 30);

                    // Reset cell boundaries cho hàng mới
                    totalWidth = 0;
                    for (var j = 0; j < properties.Columns; j++)
                    {
                        totalWidth += properties.ColumnWidths[j] * 15;
                        rtf.Append(
                            @"\clbrdrl\brdrw10\brdrs\clbrdrt\brdrw10\brdrs\clbrdrr\brdrw10\brdrs\clbrdrb\brdrw10\brdrs");
                        rtf.Append(@"\cellx" + totalWidth);
                    }
                }

                rtf.Append(@"\pard");
                rtf.Append(@"}");

                // Chèn bảng RTF
                richTextBox.SelectedRtf = rtf.ToString();

                // Thêm dòng trống sau bảng
                richTextBox.SelectionStart = richTextBox.SelectionStart + richTextBox.SelectionLength;
                richTextBox.SelectedText = Environment.NewLine + Environment.NewLine;
            }
        }

        #endregion

        private void button1_Click(object sender, EventArgs e)
        {
            
                using (OpenFileDialog openFileDialog = new OpenFileDialog())
                {
                    openFileDialog.Filter = "Image Files (.bmp;.jpg;*.jpeg;*.png;*.gif)|*.bmp;*.jpg;*.jpeg;*.png;*.gif";

                    if (openFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        string imagePath = openFileDialog.FileName;
                        Image image = Image.FromFile(imagePath);

                        Clipboard.SetImage(image);
                        richTextBox1.Paste();
                    }
                }
            
        }
    }
}
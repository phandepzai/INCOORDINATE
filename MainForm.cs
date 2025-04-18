using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Diagnostics;
using System.Reflection;

namespace ENTER_COORDINATE
{
    public partial class MainForm : Form
    {
        // Mảng lưu trữ các control nhập liệu (đầu tiên là sAPN, còn lại là các trường X/Y)
        private TextBox[] inputControls = new TextBox[13]; // Mảng chứa 13 ô nhập: 1 cho sAPN và 12 cho tọa độ X1-Y6
        private TextBox focusedControl = null; // Lưu trữ ô nhập liệu đang được focus
        private string lastSavedPath = string.Empty; // Đường dẫn lưu file cuối cùng
        private Label statusLabel; // Nhãn hiển thị trạng thái của ứng dụng
        private Label authorLabel; // Nhãn hiển thị tên tác giả
        private Label counterLabel; // Nhãn hiển thị số lượng sAPN đã lưu
        private int savedCount = 0; // Biến đếm số lượng sAPN đã lưu
        private CheckBox nasCheckBox; // Checkbox để bật/tắt chức năng lưu dữ liệu lên NAS
        private string nasPath = @"\\192.168.1.100\FAB"; // Đường dẫn mặc định đến thư mục FAB trên NAS

        // Màu sắc giao diện
        private Color authorTextColor = Color.FromArgb(200, 200, 200); // Màu xám nhạt cho tên tác giả
        private Color editTextColor = Color.Black; // Màu đen cho văn bản nhập
        private Color statusTextColor = Color.FromArgb(0, 100, 0); // Màu xanh lá cho thông báo thành công
        private Color errorTextColor = Color.Red; // Màu đỏ cho thông báo lỗi

        // Hàm khởi tạo form chính
        public MainForm()
        {
            InitializeUI(); // Khởi tạo giao diện người dùng
            var assembly = Assembly.GetExecutingAssembly();
            using (var stream = assembly.GetManifestResourceStream("ENTER_COORDINATE.icon.ico"))
            {
                this.Icon = new Icon(stream); // Đặt icon cho ứng dụng
            }
        }

        // Hàm khởi tạo giao diện người dùng
        private void InitializeUI()
        {
            // Thiết lập thuộc tính cửa sổ
            this.Text = "Tool nhập tọa độ"; // Tiêu đề của cửa sổ
            this.FormBorderStyle = FormBorderStyle.FixedSingle; // Không cho phép thay đổi kích thước
            this.MaximizeBox = false; // Không cho phép phóng to
            this.MinimizeBox = true; // Cho phép thu nhỏ
            this.Size = new Size(500, 460); // Kích thước cửa sổ
            this.BackColor = Color.White; // Màu nền trắng

            // Tạo font chữ cho các thành phần giao diện
            Font boldFont = new Font("Arial", 12, FontStyle.Bold); // Font đậm cho các nhãn
            Font editFont = new Font("Arial", 14, FontStyle.Regular); // Font thường cho ô nhập liệu
            Font keyFont = new Font("Arial", 14, FontStyle.Bold); // Font đậm cho bàn phím số
            Font statusFont = new Font("Arial", 9, FontStyle.Regular); // Font nhỏ cho nhãn trạng thái
            Font smallFont = new Font("Arial", 8, FontStyle.Regular); // Font rất nhỏ cho thông tin tác giả
            Font ButtonFont = new Font("Arial", 10, FontStyle.Bold); // Font cho các nút

            // Thêm nhãn sAPN
            Label labelSAPN = new Label
            {
                Text = "sAPN", // Nội dung nhãn
                Location = new Point(10, 25), // Vị trí
                Size = new Size(60, 25), // Kích thước
                Font = boldFont, // Sử dụng font đậm
                TextAlign = ContentAlignment.MiddleLeft // Căn lề trái
            };
            this.Controls.Add(labelSAPN); // Thêm nhãn vào form

            // Thêm ô nhập sAPN
            inputControls[0] = new TextBox
            {
                Location = new Point(70, 20), // Vị trí
                Size = new Size(300, 30), // Giảm kích thước để nhường chỗ cho checkbox
                Font = editFont, // Sử dụng font nhập liệu
                MaxLength = 300, // Giới hạn độ dài văn bản
                TextAlign = HorizontalAlignment.Left // Căn lề trái
            };
            inputControls[0].Enter += InputControl_Enter; // Gắn sự kiện khi focus vào ô
            this.Controls.Add(inputControls[0]); // Thêm ô nhập vào form

            // Thêm các nhãn và ô nhập tọa độ X/Y
            int startY = 60; // Vị trí Y bắt đầu cho các ô nhập tọa độ
            int coordSpacing = 40; // Khoảng cách giữa các ô

            for (int i = 1; i <= 12; i += 2)
            {
                // Thêm nhãn X
                Label labelX = new Label
                {
                    Text = string.Format("X{0}", (i + 1) / 2),
                    Location = new Point(10, startY + ((i - 1) / 2) * coordSpacing),
                    Size = new Size(30, 25),
                    Font = boldFont,
                    TextAlign = ContentAlignment.MiddleLeft
                };
                this.Controls.Add(labelX);

                // Thêm ô nhập X
                inputControls[i] = new TextBox
                {
                    Location = new Point(40, startY + ((i - 1) / 2) * coordSpacing),
                    Size = new Size(100, 30),
                    Font = editFont,
                    MaxLength = 10,
                    TextAlign = HorizontalAlignment.Left
                };
                inputControls[i].Enter += InputControl_Enter;
                this.Controls.Add(inputControls[i]);

                // Thêm nhãn Y
                Label labelY = new Label
                {
                    Text = string.Format("Y{0}", (i + 1) / 2),
                    Location = new Point(150, startY + ((i - 1) / 2) * coordSpacing),
                    Size = new Size(30, 25),
                    Font = boldFont,
                    TextAlign = ContentAlignment.MiddleLeft
                };
                this.Controls.Add(labelY);

                // Thêm ô nhập Y
                inputControls[i + 1] = new TextBox
                {
                    Location = new Point(180, startY + ((i - 1) / 2) * coordSpacing),
                    Size = new Size(100, 30),
                    Font = editFont,
                    MaxLength = 10,
                    TextAlign = HorizontalAlignment.Left
                };
                inputControls[i + 1].Enter += InputControl_Enter;
                this.Controls.Add(inputControls[i + 1]);
            }

            // Tạo nút RESET
            Button resetButton = new Button
            {
                Text = "RESET", // Nội dung nút
                Location = new Point(175, 320), // Vị trí
                Size = new Size(100, 40), // Kích thước
                Font = ButtonFont // Sử dụng font nút
            };
            resetButton.Click += ResetButton_Click; // Gắn sự kiện khi nhấn nút
            this.Controls.Add(resetButton); // Thêm nút vào form

            // Tạo checkbox để cho phép người dùng chọn có lưu dữ liệu lên NAS hay không
            nasCheckBox = new CheckBox
            {
                Text = "Lưu lên NAS", // Nội dung checkbox
                Location = new Point(290, 325), // Vị trí mới, ngang hàng với nút RESET
                Size = new Size(100, 30), // Kích thước nhỏ hơn
                Font = new Font("Arial", 9, FontStyle.Bold), // Font chữ nhỏ hơn
                Checked = false, // Mặc định không chọn
                ForeColor = Color.FromArgb(0, 100, 0), // Màu xanh lá
                BackColor = Color.White // Nền trắng
            };
            this.Controls.Add(nasCheckBox); // Thêm checkbox vào form

            // Tạo nhãn hiển thị số lượng sAPN đã lưu
            counterLabel = new Label
            {
                Text = string.Format("sAPN đã lưu: {0} ", savedCount),
                Location = new Point(290, 355), // Vị trí mới, bên dưới checkbox
                Size = new Size(100, 30),
                Font = new Font("Arial", 9, FontStyle.Bold),
                TextAlign = ContentAlignment.MiddleLeft,
                ForeColor = Color.FromArgb(0, 100, 0)
            };
            this.Controls.Add(counterLabel);

            // Tạo nút XÁC NHẬN
            Button confirmButton = new Button
            {
                Text = "XÁC NHẬN", // Nội dung nút
                Location = new Point(45, 320), // Vị trí
                Size = new Size(100, 40), // Kích thước
                Font = ButtonFont // Sử dụng font nút
            };
            confirmButton.Click += ConfirmButton_Click; // Gắn sự kiện khi nhấn nút
            this.Controls.Add(confirmButton); // Thêm nút vào form

            // Tạo nhãn trạng thái
            statusLabel = new Label
            {
                Text = "", // Ban đầu không hiển thị gì
                Location = new Point(10, 365), // Vị trí
                Size = new Size(330, 70), // Kích thước
                Font = statusFont, // Sử dụng font trạng thái
                TextAlign = ContentAlignment.MiddleLeft // Căn lề trái
            };
            statusLabel.Click += StatusLabel_Click; // Gắn sự kiện khi nhấn vào nhãn (để mở thư mục)
            this.Controls.Add(statusLabel); // Thêm nhãn vào form

            // Tạo nhãn tác giả
            authorLabel = new Label
            {
                Text = "Tác giả: Nông Văn Phấn", // Nội dung nhãn
                Location = new Point(200, 395), // Vị trí cũ
                Size = new Size(280, 25), // Kích thước
                Font = smallFont, // Sử dụng font nhỏ
                TextAlign = ContentAlignment.MiddleRight, // Căn lề phải
                ForeColor = authorTextColor // Sử dụng màu xám nhạt
            };
            this.Controls.Add(authorLabel); // Thêm nhãn vào form

            // Tạo bàn phím số ảo
            string[] keys = { "7", "8", "9", "4", "5", "6", "1", "2", "3", "0", "« XÓA" }; // Các phím số và phím xóa
            int buttonSize = 55; // Kích thước chuẩn của một nút
            int spacing = 9;     // Khoảng cách giữa các nút

            for (int i = 0; i < keys.Length; i++)
            {
                Button keyButton = new Button
                {
                    Text = keys[i], // Nội dung nút
                    Font = keyFont // Sử dụng font phím
                };

                if (keys[i] == "0")
                {
                    // Đặt nút số 0 ở hàng cuối, chiếm 1 cột
                    keyButton.Location = new Point(290, 70 + 3 * (buttonSize + spacing));
                    keyButton.Size = new Size(buttonSize, buttonSize);
                }
                else if (keys[i] == "« XÓA")
                {
                    // Đặt nút XÓA ở hàng cuối, chiếm 2 cột
                    keyButton.Location = new Point(290 + (buttonSize + spacing), 70 + 3 * (buttonSize + spacing));
                    keyButton.Size = new Size(2 * buttonSize + spacing, buttonSize);
                }
                else
                {
                    // Các nút khác sắp xếp theo lưới 3x3
                    keyButton.Location = new Point(290 + (i % 3) * (buttonSize + spacing),
                                                   70 + (i / 3) * (buttonSize + spacing));
                    keyButton.Size = new Size(buttonSize, buttonSize);
                }

                keyButton.Tag = keys[i]; // Lưu trữ giá trị của phím
                keyButton.Click += KeypadButton_Click; // Gắn sự kiện khi nhấn phím
                this.Controls.Add(keyButton); // Thêm nút vào form
            }

            // Đặt focus ban đầu vào ô nhập sAPN
            inputControls[0].Focus();
        }

        // Xử lý sự kiện khi người dùng focus vào một ô nhập
        private void InputControl_Enter(object sender, EventArgs e)
        {
            focusedControl = (TextBox)sender; // Lưu trữ ô đang được focus để bàn phím số biết nhập vào đâu
        }

        // Xử lý sự kiện khi người dùng nhấn một phím trên bàn phím số ảo
        private void KeypadButton_Click(object sender, EventArgs e)
        {
            if (focusedControl == null) return; // Nếu không có ô nào được focus, không làm gì cả

            Button button = (Button)sender;
            string key = button.Tag.ToString(); // Lấy giá trị của phím (số hoặc "XÓA")

            if (key == "« XÓA")
            {
                // Xử lý nút xóa - xóa ký tự cuối cùng
                if (focusedControl.Text.Length > 0)
                {
                    focusedControl.Text = focusedControl.Text.Substring(0, focusedControl.Text.Length - 1);
                    focusedControl.SelectionStart = focusedControl.Text.Length; // Đặt con trỏ ở cuối
                }
            }
            else
            {
                // Chỉ cho phép số trong các trường X/Y
                if (focusedControl != inputControls[0])
                {
                    // Thêm số vào các ô X/Y
                    focusedControl.Text += key;
                    focusedControl.SelectionStart = focusedControl.Text.Length; // Đặt con trỏ ở cuối
                }
                else
                {
                    // Cho phép mọi văn bản trong trường sAPN
                    focusedControl.Text += key;
                    focusedControl.SelectionStart = focusedControl.Text.Length; // Đặt con trỏ ở cuối
                }
            }
        }

        // Xử lý sự kiện khi reset tất cả các trường nhập liệu
        private void ResetInputs()
        {
            foreach (TextBox control in inputControls)
            {
                control.Text = ""; // Xóa nội dung của tất cả các ô nhập
            }
            inputControls[0].Focus(); // Đặt focus vào ô sAPN
            statusLabel.Text = "Đã xóa dữ liệu các ô. Vui lòng nhập lại."; // Hiển thị thông báo
            statusLabel.ForeColor = statusTextColor; // Đặt màu xanh lá cho thông báo
        }

        // Xử lý sự kiện khi nhấn nút RESET
        private void ResetButton_Click(object sender, EventArgs e)
        {
            ResetInputs(); // Gọi hàm reset các trường nhập liệu
        }

        // Xử lý sự kiện khi nhấn nút XÁC NHẬN
        private void ConfirmButton_Click(object sender, EventArgs e)
        {
            SaveToCSV(); // Gọi hàm lưu dữ liệu vào file CSV
        }

        // Hàm lưu dữ liệu vào file CSV
        private void SaveToCSV()
        {
            // Thu thập giá trị từ các ô nhập liệu
            List<string> values = new List<string>();
            bool hasSAPN = false; // Cờ kiểm tra xem ô sAPN có dữ liệu không
            bool hasAnyXY = false; // Cờ kiểm tra xem có ít nhất một ô X/Y có dữ liệu không

            foreach (TextBox control in inputControls)
            {
                values.Add(control.Text); // Thêm giá trị của từng ô vào danh sách
            }

            // Kiểm tra xem sAPN có dữ liệu không
            if (!string.IsNullOrEmpty(values[0]))
            {
                hasSAPN = true;
            }

            // Kiểm tra xem ít nhất một trường X/Y có dữ liệu không
            for (int i = 1; i < values.Count; i++)
            {
                if (!string.IsNullOrEmpty(values[i]))
                {
                    hasAnyXY = true;

                    // Kiểm tra xem trường X/Y chỉ chứa số
                    foreach (char c in values[i])
                    {
                        if (!char.IsDigit(c))
                        {
                            statusLabel.Text = "Lỗi: Ô X/Y chỉ được chứa số!"; // Thông báo lỗi
                            statusLabel.ForeColor = errorTextColor; // Đặt màu đỏ cho thông báo lỗi
                            return; // Dừng việc lưu
                        }
                    }
                }
            }

            // Kiểm tra điều kiện: sAPN và ít nhất một trường X/Y phải có dữ liệu
            if (!hasSAPN || !hasAnyXY)
            {
                statusLabel.Text = "Lỗi: Cần nhập sAPN và ít nhất một ô X hoặc Y!"; // Thông báo lỗi
                statusLabel.ForeColor = errorTextColor; // Đặt màu đỏ cho thông báo lỗi
                return; // Dừng việc lưu
            }

            // Lấy thời gian hiện tại
            DateTime now = DateTime.Now;

            // Tạo thư mục FAB trên Desktop nếu chưa tồn tại
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop); // Lấy đường dẫn đến Desktop
            string fabPath = Path.Combine(desktopPath, "FAB"); // Tạo đường dẫn đến thư mục FAB

            try
            {
                if (!Directory.Exists(fabPath))
                {
                    Directory.CreateDirectory(fabPath); // Tạo thư mục FAB nếu chưa tồn tại
                }
            }
            catch
            {
                statusLabel.Text = "Lỗi: Không thể tạo thư mục FAB!"; // Thông báo lỗi
                statusLabel.ForeColor = errorTextColor; // Đặt màu đỏ cho thông báo lỗi
                return; // Dừng việc lưu
            }

            // Đặt tên file theo định dạng ngày tháng: NVP_ddMMyyyy.csv
            string filename = Path.Combine(fabPath, string.Format("NVP_{0:ddMMyyyy}.csv", now));

            // Kiểm tra xem file đã có tiêu đề chưa
            bool hasHeader = false;
            if (File.Exists(filename)) // Nếu file đã tồn tại
            {
                try
                {
                    string firstLine = File.ReadLines(filename).FirstOrDefault(); // Đọc dòng đầu tiên
                    hasHeader = firstLine != null && firstLine.Contains("sAPN"); // Kiểm tra xem dòng đầu tiên có chứa "sAPN" không
                }
                catch
                {
                    // Không thể đọc file
                    statusLabel.Text = "Lỗi: Không thể đọc file!\nFile đang mở, hãy tắt đi."; // Thông báo lỗi
                    statusLabel.ForeColor = errorTextColor; // Đặt màu đỏ cho thông báo lỗi
                    return; // Dừng việc lưu
                }
            }

            // Ghi dữ liệu vào file
            try
            {
                // Lưu vào thư mục local trên máy tính
                using (StreamWriter file = new StreamWriter(filename, true, Encoding.UTF8))
                {
                    if (!hasHeader) // Nếu file chưa có tiêu đề, thêm tiêu đề
                    {
                        file.WriteLine("sAPN,X1|Y1,X2|Y2,X3|Y3,X4|Y4,X5|Y5,X6|Y6,EVENTIME");
                    }

                    // Ghi dữ liệu sAPN
                    file.Write(EscapeCSV(values[0]));

                    // Ghi các cặp tọa độ X/Y
                    for (int i = 1; i <= 11; i += 2)
                    {
                        file.Write(",");
                        string pair = values[i];
                        if (!string.IsNullOrEmpty(values[i]) || !string.IsNullOrEmpty(values[i + 1]))
                        {
                            pair += "," + values[i + 1];
                        }
                        file.Write(EscapeCSV(pair));
                    }

                    // Ghi thời gian
                    file.WriteLine(string.Format(",{0:dd/MM/yyyy HH:mm:ss}", now));
                }

                // Kiểm tra nếu người dùng đã chọn lưu lên NAS
                if (nasCheckBox.Checked)
                {
                    try
                    {
                        // Tạo thư mục FAB trên NAS nếu chưa tồn tại
                        if (!Directory.Exists(nasPath))
                        {
                            Directory.CreateDirectory(nasPath);
                        }

                        // Tạo tên file trên NAS theo định dạng ngày tháng
                        string nasFilename = Path.Combine(nasPath, string.Format("NVP_{0:ddMMyyyy}.csv", now));

                        // Ghi dữ liệu vào file trên NAS
                        using (StreamWriter file = new StreamWriter(nasFilename, true, Encoding.UTF8))
                        {
                            // Kiểm tra và thêm tiêu đề nếu file mới
                            string firstLine = null;
                            if (File.Exists(nasFilename))
                            {
                                firstLine = File.ReadLines(nasFilename).FirstOrDefault();
                            }
                            if (!File.Exists(nasFilename) || firstLine == null || !firstLine.Contains("sAPN"))
                            {
                                file.WriteLine("sAPN,X1|Y1,X2|Y2,X3|Y3,X4|Y4,X5|Y5,X6|Y6,EVENTIME");
                            }

                            // Ghi dữ liệu sAPN
                            file.Write(EscapeCSV(values[0]));

                            // Ghi các cặp tọa độ X/Y
                            for (int i = 1; i <= 11; i += 2)
                            {
                                file.Write(",");
                                string pair = values[i];
                                if (!string.IsNullOrEmpty(values[i]) || !string.IsNullOrEmpty(values[i + 1]))
                                {
                                    pair += "," + values[i + 1];
                                }
                                file.Write(EscapeCSV(pair));
                            }

                            // Ghi thời gian
                            file.WriteLine(string.Format(",{0:dd/MM/yyyy HH:mm:ss}", now));
                        }
                    }
                    catch (Exception ex)
                    {
                        // Hiển thị thông báo lỗi nếu không thể lưu lên NAS
                        statusLabel.Text = string.Format("Lỗi khi lưu lên NAS: {0}", ex.Message);
                        statusLabel.ForeColor = errorTextColor;
                        return;
                    }
                }

                // Lưu đường dẫn FAB để có thể mở thư mục sau khi lưu
                lastSavedPath = fabPath;

                // Tăng biến đếm và cập nhật hiển thị số lượng sAPN đã lưu
                savedCount++;
                counterLabel.Text = string.Format("sAPN đã lưu: {0} ", savedCount);

                // Xóa các ô nhập và đặt focus vào ô đầu tiên
                ResetInputs();

                // Hiển thị thông báo thành công, có thêm thông tin về việc lưu lên NAS nếu được chọn
                string successMessage = nasCheckBox.Checked ? 
                    string.Format("Đã lưu dữ liệu thành công (Local + NAS): {0:dd/MM/yyyy HH:mm:ss}\nPath: {1}", now, filename) :
                    string.Format("Đã lưu dữ liệu thành công: {0:dd/MM/yyyy HH:mm:ss}\nPath: {1}", now, filename);
                
                statusLabel.Text = successMessage;
                statusLabel.ForeColor = statusTextColor;
            }
            catch
            {
                // Hiển thị thông báo lỗi nếu không thể lưu file
                statusLabel.Text = "Lỗi: Không thể ghi dữ liệu vào file!\nFile đang mở, hãy tắt đi.";
                statusLabel.ForeColor = errorTextColor;
            }
        }

        // Hàm xử lý chuỗi để đảm bảo đúng định dạng CSV
        private string EscapeCSV(string input)
        {
            if (string.IsNullOrEmpty(input) || !input.Contains(",") && !input.Contains("\"") && !input.Contains("\n"))
            {
                return input; // Nếu không có ký tự đặc biệt, giữ nguyên chuỗi
            }

            return "\"" + input.Replace("\"", "\"\"") + "\""; // Đặt chuỗi trong dấu ngoặc kép và thay thế dấu ngoặc kép bằng hai dấu ngoặc kép
        }

        // Xử lý sự kiện khi nhấn vào nhãn trạng thái
        private void StatusLabel_Click(object sender, EventArgs e)
        {
            // Mở thư mục FAB khi nhấp vào nhãn trạng thái
            if (!string.IsNullOrEmpty(lastSavedPath) && Directory.Exists(lastSavedPath))
            {
                Process.Start(lastSavedPath); // Mở thư mục FAB trong File Explorer
            }
        }
    }
}
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using WindowsFormsApp1;

namespace SE_MVICHECKERLOG
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        string _checker = null;
        string _pathServer = null;

        public MainWindow()
        {
            InitializeComponent();

            string iniFilePath = AppDomain.CurrentDomain.BaseDirectory.ToString() + "Setting\\setting.ini";
            iniHelper iniHandler = new iniHelper(iniFilePath);
            _checker = iniHandler.ReadValue("MACHINE", "checker");
            _pathServer = iniHandler.ReadValue("FILE", "pathserver");
            if(_pathServer=="")
            {
                _pathServer = Environment.CurrentDirectory + "\\MVILog";
            }

            
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            ResetButton(Brushes.WhiteSmoke);
            EnableButtonError(false);

            cboError.ItemsSource = new List<string> {"OK", "Dirt/Stain/Contamination", "White dot", "Scratch", "Lens Drop", "Bubble", "Black dot", "Fresnel Yellowing", "Discolor", "Shiny Line", "Edge Chipping", "Fiber" };

            txtIdSheet.Focus();

            //btnTA1.Background = Brushes.Red;
        }

        private void RadioButton_Checked(object sender, RoutedEventArgs e)
        {
            // Đây là nơi xử lý sự kiện khi một RadioButton được chọn
            RadioButton radioButton = sender as RadioButton;
            if (radioButton != null)
            {
                MessageBox.Show(radioButton.Content + " được chọn.");
            }
        }

        public string[] dataSheet = new string[]
        {
            "", "", "Stop",
            "OK", "OK", "OK", "OK", "OK", "OK", "OK",   // G-1 -> A-1
            "OK", "OK", "OK", "OK", "OK", "OK", "OK",   // G-2 -> A-2
            "OK", "OK", "OK", "OK", "OK", "OK", "OK",   // G-3 -> A-3
            "OK", "OK", "OK", "OK", "OK", "OK", "OK",   // G-4 -> A-4
            "OK", "OK", "OK", "OK", "OK", "OK", "OK",   // G-5 -> A-5
            "OK", "OK", "OK", "OK", "OK", "OK", "OK",   // G-6 -> A-6
            "OK", "OK", "OK", "OK", "OK", "OK", "OK",   // G-7 -> A-7
            "OK", "OK", "OK", "OK", "OK", "OK", "OK",   // G-8 -> A-8
            "OK", "OK", "OK", "OK", "OK", "OK", "OK",   // G-9 -> A-9
            "63", "100%"
         };

        public string[] arrMaps =
        {
            "G1", "F1", "E1", "D1", "C1", "B1", "A1",
            "G2", "F2", "E2", "D2", "C2", "B2", "A2",
            "G3", "F3", "E3", "D3", "C3", "B3", "A3",
            "G4", "F4", "E4", "D4", "C4", "B4", "A4",
            "G5", "F5", "E5", "D5", "C5", "B5", "A5",
            "G6", "F6", "E6", "D6", "C6", "B6", "A6",
            "G7", "F7", "E7", "D7", "C7", "B7", "A7",
            "G8", "F8", "E8", "D8", "C8", "B8", "A8",
            "G9", "F9", "E9", "D9", "C9", "B9", "A9"
        };
        private void TextBox_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Key == Key.Enter)
                {
                    if (txtIdSheet.Text.Trim().Length > 0)
                    {
                        dataSheet = new string[]
                        {
                            "", "", "Stop",
                            "OK", "OK", "OK", "OK", "OK", "OK", "OK",   // G-1 -> A-1
                            "OK", "OK", "OK", "OK", "OK", "OK", "OK",   // G-2 -> A-2
                            "OK", "OK", "OK", "OK", "OK", "OK", "OK",   // G-3 -> A-3
                            "OK", "OK", "OK", "OK", "OK", "OK", "OK",   // G-4 -> A-4
                            "OK", "OK", "OK", "OK", "OK", "OK", "OK",   // G-5 -> A-5
                            "OK", "OK", "OK", "OK", "OK", "OK", "OK",   // G-6 -> A-6
                            "OK", "OK", "OK", "OK", "OK", "OK", "OK",   // G-7 -> A-7
                            "OK", "OK", "OK", "OK", "OK", "OK", "OK",   // G-8 -> A-8
                            "OK", "OK", "OK", "OK", "OK", "OK", "OK",   // G-9 -> A-9
                            "63", "100%"
                        };

                        DateTime dtStart = DateTime.Now;
                        dataSheet[0] = txtIdSheet.Text.Trim();
                        dataSheet[1] = dtStart.ToString("yyyy/MM/dd HH:mm:ss");
                        //dataSheet[67] = "10%";
                        txtSheetCF.Text = txtIdSheet.Text.Trim();

                        btnCF.IsEnabled = true;
                        ResetButton(Brushes.Green);
                        EnableButtonError(true);
                    }
                }
                //KeyDown += MainWindow_KeyDown;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return;
            }
        }

        private void Window_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            //if (txtIdSheet.Text.Trim().Length == 0)
            //{ return; }

            ////Next
            //if (e.Key == Key.Right && CheckLen == true)
            //{
            //    if (selectedButtonIndex < btnError.Count - 1)
            //    {
            //        selectedButtonIndex++;
            //        string namebt = btnError[selectedButtonIndex].Name;
            //        txtPosition.Text = namebt.Substring(namebt.Length - 2, 2);
            //        txtError.Text = lstError[selectedButtonIndex].ToString();
            //        CheckLen = false;

            //        ImgBot.Visibility = Visibility.Visible;
            //        ImgTop.Visibility = Visibility.Hidden;
            //    }
            //}

            ////Đổi nền
            //if (e.Key == Key.Insert)
            //{
            //    ChangeBackGround_Click(sender, e);
            //}

            ////Điền OK
            //if (e.Key == Key.Home)
            //{
            //    btnError[selectedButtonIndex].Background = new SolidColorBrush(Colors.DarkSeaGreen);

            //    CheckLen = true;
            //}

            ////Điền NG
            //if (e.Key == Key.End)
            //{
            //    btnError[selectedButtonIndex].Background = new SolidColorBrush(Colors.DarkRed);

            //    CheckLen = true;
            //}

            //New sheet
            if (e.Key == Key.Insert)
            {
                txtIdSheet.Clear();
                txtIdSheet.Focus();
            }
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            // Xử lý sự kiện khi một trong 10 nút được nhấn
            Button clickedButton = sender as Button;
            if (clickedButton != null)
            {
                //MessageBox.Show($"Bạn đã nhấn nút: {clickedButton.Name}");
                txtPosition.Text = clickedButton.Name.Substring(4);
                cboError.SelectedItem = dataSheet[Array.IndexOf(arrMaps, txtPosition.Text) + 2];
                txtErrorCF.Text = dataSheet[Array.IndexOf(arrMaps, txtPosition.Text) + 2];
                clickedButton.Content = "o";

            }
        }

        private void Button_GotFocus(object sender, RoutedEventArgs e)
        {
            // Xử lý sự kiện khi một trong 5 nút nhận được focus
            Button focusedButton = sender as Button;
            if (focusedButton != null)
            {
                //MessageBox.Show($"Nút {focusedButton.Name} đã nhận được focus.");
                txtPosition.Text = focusedButton.Name.Substring(4);
                cboError.SelectedItem = dataSheet[Array.IndexOf(arrMaps, txtPosition.Text) + 2];
                txtErrorCF.Text = dataSheet[Array.IndexOf(arrMaps, txtPosition.Text) + 2];
            }
        }

        private void Button_LostFocus(object sender, RoutedEventArgs e)
        {
            // Xử lý sự kiện khi một trong 5 nút nhận được focus
            Button focusedButton = sender as Button;
            if (focusedButton != null)
            {
                focusedButton.Content = "";
            }
        }

        private void cboError_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            string valuecbo = cboError.SelectedItem.ToString();

            if (txtPosition.Text.Trim() == "-"
                || txtPosition.Text.Trim() == "")
            {
                //MessageBox.Show("Chọn vị trí Lens!");
                cboError.SelectedIndex = 0;
                return;
            }    

            else
            {
                if (valuecbo == "OK")
                {
                    Button btn = this.FindName("btnT" + txtPosition.Text) as Button;
                    if (btn != null)
                    {
                        dataSheet[Array.IndexOf(arrMaps, txtPosition.Text) + 2] = "OK";

                        ChangeColorBTN(btn, true);
                        return;
                    }
                }
                else
                {
                    Button btn = this.FindName("btnT" + txtPosition.Text) as Button;
                    if (btn != null)
                    {
                        dataSheet[Array.IndexOf(arrMaps, txtPosition.Text) + 2] = valuecbo;

                        ChangeColorBTN(btn, false);
                        return;
                    }

                }    
            } 
        }

        private void ChangeColorBTN(Button btn, bool color)
        {
            if(color)
                btn.Background = Brushes.Green;
            else
                btn.Background = Brushes.Red;
        }

        private void btnCF_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                dataSheet[2] = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                int count = 0;
                for (int i = 0; i < dataSheet.Length; i++)
                {
                    if (dataSheet[i] == "OK")
                    {
                        count++;
                    }
                }
                dataSheet[66] = count.ToString();
                dataSheet[67] = Math.Round((((double)count / 63) * 100), 1) + "%";
                string templatePath = Environment.CurrentDirectory + "\\Template\\Checker.csv";
                string outputPath = "";
                if (rdTruoc.IsChecked == true)
                {
                    outputPath = _pathServer + "FRONT" + "\\" + _checker + "_" + DateTime.Now.ToString("yyyyMMdd") + ".csv";

                    bool exists = System.IO.Directory.Exists(_pathServer + "FRONT");
                    if (!exists)
                        System.IO.Directory.CreateDirectory(_pathServer + "FRONT");
                }
                if (rdSau.IsChecked == true)
                {
                    outputPath = _pathServer + "BACK" + "\\" + _checker + "_" + DateTime.Now.ToString("yyyyMMdd") + ".csv";

                    bool exists = System.IO.Directory.Exists(_pathServer + "BACK");
                    if (!exists)
                        System.IO.Directory.CreateDirectory(_pathServer + "BACK");
                }

                if (!File.Exists(outputPath))
                {
                    File.Copy(templatePath, outputPath);
                }

                WriteCSVLogMVI(outputPath, dataSheet);
                MessageBox.Show("Lưu thành công!");
                ResetButton(Brushes.WhiteSmoke);
                EnableButtonError(false);
                txtIdSheet.Text = "";
                txtPosition.Text = "_";
                txtErrorCF.Text = "-";
                txtSheetCF.Text = "";
                cboError.SelectedIndex = 0;
                txtIdSheet.Focus();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + ex.StackTrace);
                return;
            }
            

            //// Đọc nội dung từ file template
            //var lines = new List<string>(File.ReadAllLines(templatePath));

            //// Thêm dữ liệu mới
            //lines.Add("A1,B1,C1,D1");
            //lines.Add("A2,B2,C2,D2");
            //lines.Add("A3,B3,C3,D3");


        }

        private void WriteCSVLogMVI(string filePath, string[] outputData)
        {
            //string fileTemp = Environment.CurrentDirectory + "\\" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".csv";
            //using (var reader = new StreamReader(filePath))
            //using (var writer = new StreamWriter(fileTemp))
            //{
            //    int currentRow = 0;

            //    while (!reader.EndOfStream)
            //    {
            //        string line = reader.ReadLine();
            //        currentRow++;

            //        string[] columns = line.Split(','); // Tách cột

            //        // Ghi vào cột cuối cùng nếu là dòng 1,2,3
            //        if (targetRows.Contains(currentRow))
            //        {
            //            Array.Resize(ref columns, columns.Length + 1); // Thêm 1 cột
            //            columns[columns.Length - 1] = newValue; // Ghi vào cột cuối
            //        }

            //        writer.WriteLine(string.Join(",", columns)); // Ghi vào file tạm
            //    }
            //}
            // Ghi đè file CSV gốc bằng file mới
            //File.Delete(filePath);
            //File.Move(fileTemp, filePath);


            #region Cập nhật số lương summary
            List<string[]> csvData = new List<string[]>();
            // Đọc toàn bộ file CSV
            // Đọc toàn bộ file CSV

            // Danh sách các ô cần tăng (dòng, cột)
            var targetCells = new List<Tuple<int, int>>
            {
                Tuple.Create(3,Array.IndexOf(lstErrorCSV, outputData[3])+2),  // Dòng 3, cột 2
                Tuple.Create(4,Array.IndexOf(lstErrorCSV, outputData[4])+2),  // Dòng 4, cột 5
                Tuple.Create(5, Array.IndexOf(lstErrorCSV, outputData[5]) + 2), // Dòng 10, cột 8
                Tuple.Create(6, Array.IndexOf(lstErrorCSV, outputData[6]) + 2), // Dòng 67, cột 13
                Tuple.Create(7, Array.IndexOf(lstErrorCSV, outputData[7]) + 2),  // Dòng 3, cột 2
                Tuple.Create(8, Array.IndexOf(lstErrorCSV, outputData[8]) + 2),  // Dòng 4, cột 5
                Tuple.Create(9, Array.IndexOf(lstErrorCSV, outputData[9]) + 2), // Dòng 10, cột 8
                Tuple.Create(10, Array.IndexOf(lstErrorCSV, outputData[10]) + 2), // Dòng 67, cột 13
                Tuple.Create(11, Array.IndexOf(lstErrorCSV, outputData[11]) + 2),  // Dòng 3, cột 2
                Tuple.Create(12, Array.IndexOf(lstErrorCSV, outputData[12]) + 2),  // Dòng 4, cột 5
                Tuple.Create(13, Array.IndexOf(lstErrorCSV, outputData[13]) + 2), // Dòng 10, cột 8
                Tuple.Create(14, Array.IndexOf(lstErrorCSV, outputData[14]) + 2), // Dòng 67, cột 13
                Tuple.Create(15, Array.IndexOf(lstErrorCSV, outputData[15]) + 2),  // Dòng 3, cột 2
                Tuple.Create(16, Array.IndexOf(lstErrorCSV, outputData[16]) + 2),  // Dòng 4, cột 5
                Tuple.Create(17, Array.IndexOf(lstErrorCSV, outputData[17]) + 2), // Dòng 10, cột 8
                Tuple.Create(18, Array.IndexOf(lstErrorCSV, outputData[18]) + 2), // Dòng 67, cột 13
                Tuple.Create(19, Array.IndexOf(lstErrorCSV, outputData[19]) + 2),  // Dòng 3, cột 2
                Tuple.Create(20, Array.IndexOf(lstErrorCSV, outputData[20]) + 2),  // Dòng 4, cột 5
                Tuple.Create(21, Array.IndexOf(lstErrorCSV, outputData[21]) + 2), // Dòng 10, cột 8
                Tuple.Create(22, Array.IndexOf(lstErrorCSV, outputData[22]) + 2), // Dòng 67, cột 13
                Tuple.Create(23, Array.IndexOf(lstErrorCSV, outputData[23]) + 2),  // Dòng 3, cột 2
                Tuple.Create(24, Array.IndexOf(lstErrorCSV, outputData[24]) + 2),  // Dòng 4, cột 5
                Tuple.Create(25, Array.IndexOf(lstErrorCSV, outputData[25]) + 2), // Dòng 10, cột 8
                Tuple.Create(26, Array.IndexOf(lstErrorCSV, outputData[26]) + 2), // Dòng 67, cột 13
                Tuple.Create(27, Array.IndexOf(lstErrorCSV, outputData[27]) + 2),  // Dòng 3, cột 2
                Tuple.Create(28, Array.IndexOf(lstErrorCSV, outputData[28]) + 2),  // Dòng 4, cột 5
                Tuple.Create(29, Array.IndexOf(lstErrorCSV, outputData[29]) + 2), // Dòng 10, cột 8
                Tuple.Create(30, Array.IndexOf(lstErrorCSV, outputData[30]) + 2), // Dòng 67, cột 13
                Tuple.Create(31, Array.IndexOf(lstErrorCSV, outputData[31]) + 2),  // Dòng 3, cột 2
                Tuple.Create(32, Array.IndexOf(lstErrorCSV, outputData[32]) + 2),  // Dòng 4, cột 5
                Tuple.Create(33, Array.IndexOf(lstErrorCSV, outputData[33]) + 2), // Dòng 10, cột 8
                Tuple.Create(34, Array.IndexOf(lstErrorCSV, outputData[34]) + 2), // Dòng 67, cột 13
                Tuple.Create(35, Array.IndexOf(lstErrorCSV, outputData[35]) + 2),  // Dòng 3, cột 2
                Tuple.Create(36, Array.IndexOf(lstErrorCSV, outputData[36]) + 2),  // Dòng 4, cột 5
                Tuple.Create(37, Array.IndexOf(lstErrorCSV, outputData[37]) + 2), // Dòng 10, cột 8
                Tuple.Create(38,Array.IndexOf(lstErrorCSV, outputData[38]) + 2), // Dòng 67, cột 13
                Tuple.Create(39,Array.IndexOf(lstErrorCSV, outputData[39]) + 2),  // Dòng 3, cột 2
                Tuple.Create(40,Array.IndexOf(lstErrorCSV, outputData[40]) + 2),  // Dòng 4, cột 5
                Tuple.Create(41,Array.IndexOf(lstErrorCSV, outputData[41]) + 2), // Dòng 10, cột 8
                Tuple.Create(42,Array.IndexOf(lstErrorCSV, outputData[42]) + 2), // Dòng 67, cột 13
                Tuple.Create(43,Array.IndexOf(lstErrorCSV, outputData[43]) + 2),  // Dòng 3, cột 2
                Tuple.Create(44,Array.IndexOf(lstErrorCSV, outputData[44]) + 2),  // Dòng 4, cột 5
                Tuple.Create(45,Array.IndexOf(lstErrorCSV, outputData[45]) + 2), // Dòng 10, cột 8
                Tuple.Create(46,Array.IndexOf(lstErrorCSV, outputData[46]) + 2), // Dòng 67, cột 13
                Tuple.Create(47,Array.IndexOf(lstErrorCSV, outputData[47]) + 2),  // Dòng 3, cột 2
                Tuple.Create(48,Array.IndexOf(lstErrorCSV, outputData[48]) + 2),  // Dòng 4, cột 5
                Tuple.Create(49,Array.IndexOf(lstErrorCSV, outputData[49]) + 2), // Dòng 10, cột 8
                Tuple.Create(50,Array.IndexOf(lstErrorCSV, outputData[50]) + 2), // Dòng 67, cột 13
                Tuple.Create(51,Array.IndexOf(lstErrorCSV, outputData[51]) + 2), // Dòng 67, cột 13
                Tuple.Create(52,Array.IndexOf(lstErrorCSV, outputData[52]) + 2),  // Dòng 3, cột 2
                Tuple.Create(53,Array.IndexOf(lstErrorCSV, outputData[53]) + 2),  // Dòng 4, cột 5
                Tuple.Create(54,Array.IndexOf(lstErrorCSV, outputData[54]) + 2), // Dòng 10, cột 8
                Tuple.Create(55,Array.IndexOf(lstErrorCSV, outputData[55]) + 2), // Dòng 67, cột 13
                Tuple.Create(56, Array.IndexOf(lstErrorCSV, outputData[56]) + 2),  // Dòng 3, cột 2
                Tuple.Create(57, Array.IndexOf(lstErrorCSV, outputData[57]) + 2),  // Dòng 4, cột 5
                Tuple.Create(58, Array.IndexOf(lstErrorCSV, outputData[58]) + 2), // Dòng 10, cột 8
                Tuple.Create(59, Array.IndexOf(lstErrorCSV, outputData[59]) + 2), // Dòng 67, cột 13
                Tuple.Create(60, Array.IndexOf(lstErrorCSV, outputData[60]) + 2),  // Dòng 3, cột 2
                Tuple.Create(61, Array.IndexOf(lstErrorCSV, outputData[61]) + 2),  // Dòng 4, cột 5
                Tuple.Create(62, Array.IndexOf(lstErrorCSV, outputData[62]) + 2), // Dòng 10, cột 8
                Tuple.Create(63, Array.IndexOf(lstErrorCSV, outputData[63]) + 2), // Dòng 67, cột 13
                Tuple.Create(64, Array.IndexOf(lstErrorCSV, outputData[64]) + 2),  // Dòng 3, cột 2
                Tuple.Create(65, Array.IndexOf(lstErrorCSV, outputData[65]) + 2),  // Dòng 4, cột 5
                //Tuple.Create(66, Array.IndexOf(lstErrorCSV, outputData[66]) + 2), // Dòng 10, cột 8
                //Tuple.Create(67, Array.IndexOf(lstErrorCSV, outputData[67]) + 2) // Dòng 67, cột 13
            };

            //// Đọc toàn bộ file CSV
            //var lines = File.ReadAllLines(filePath).ToList();
            // Đọc dữ liệu từ file CSV
            List<string[]> lines1 = File.Exists(filePath)
                ? File.ReadAllLines(filePath).Select(line => line.Split(',')).ToList()
                : new List<string[]>();

            foreach (var cell in targetCells)
            {
                int rowIndex = cell.Item1; // Chuyển về chỉ mục bắt đầu từ 0
                int colIndex = cell.Item2;

                // Kiểm tra nếu ô có giá trị là số thì tăng thêm 1
                if (int.TryParse(lines1[rowIndex][colIndex], out int number))
                {
                    lines1[rowIndex][colIndex] = (number + 1).ToString();
                }
                //else
                //{
                //    lines1[rowIndex][colIndex] = "1";
                //}
            }

            // Ghi lại dữ liệu vào CSV
            File.WriteAllLines(filePath, lines1.Select(line => string.Join(",", line)));
            #endregion

            // Đọc toàn bộ nội dung file CSV
            List<string> lines = File.Exists(filePath)
                ? File.ReadAllLines(filePath).ToList()
                : new List<string>();

            // Lấy giá trị dòng 1 cột 1 và tăng thêm 1
            if (lines.Count > 0)
            {
                var firstLine = lines[0].Split(',');
                if (firstLine.Length > 0 && int.TryParse(firstLine[0], out int value))
                {
                    firstLine[0] = (value + 1).ToString();
                    lines[0] = string.Join(",", firstLine);
                }
            }

            // Ghi dữ liệu vào cột cuối cùng
            for (int i = 0; i < outputData.Length; i++)
            {
                if (i >= lines.Count)
                {
                    lines.Add(outputData[i]);
                }
                else
                {
                    var lineData = lines[i].Split(',');
                    lines[i] = string.Join(",", lineData) + "," + outputData[i];
                }
            }

            // Ghi lại dữ liệu vào file CSV
            File.WriteAllLines(filePath, lines);
        }

        private void ResetButton(Brush color)
        {
            var buttons = GetAllButtons();
            foreach (var btn in buttons)
            {
                btn.Background = color;
            }
        }

        private List<Button> GetAllButtons()
        {
            return FindVisualChildren<Button>(Application.Current.MainWindow)
                   .Where(btn => btn.Name.Contains("btnT"))
                   .ToList();
        }

        private List<Button> GetAllButtonsError()
        {
            return FindVisualChildren<Button>(Application.Current.MainWindow)
                   .Where(btn => btn.Name.Contains("btnError"))
                   .ToList();
        }

        private void EnableButtonError(bool check)
        {
            var buttons = GetAllButtonsError();
            foreach (var btn in buttons)
            {
                btn.IsEnabled = check;
            }

            var buttons1 = GetAllButtons();
            foreach (var btn in buttons1)
            {
                btn.IsEnabled = check;
            }
        }

        public static IEnumerable<T> FindVisualChildren<T>(DependencyObject depObj) where T : DependencyObject
        {
            if (depObj != null)
            {
                for (int i = 0; i < VisualTreeHelper.GetChildrenCount(depObj); i++)
                {
                    DependencyObject child = VisualTreeHelper.GetChild(depObj, i);
                    if (child != null && child is T)
                    {
                        yield return (T)child;
                    }

                    foreach (T childOfChild in FindVisualChildren<T>(child))
                    {
                        yield return childOfChild;
                    }
                }
            }
        }

        List<string> lstError = new List<string> {"Dirt/Stain/Contamination", "White dot", "Scratch", "Lens Drop", "Bubble", "Black dot", "Fresnel Yellowing", "Discolor", "Shiny Line", "Edge Chipping", "Fiber", "OK" };
        string[] lstErrorCSV = new string[] { "OK", "Dirt/Stain/Contamination", "White dot", "Scratch", "Lens Drop", "Bubble", "Black dot", "Fresnel Yellowing", "Discolor", "Shiny Line", "Edge Chipping", "Fiber" };

        private void ButtonError_Click(object sender, RoutedEventArgs e)
        {
            //Array.IndexOf(lstErrorCSV,)


            if (txtIdSheet.Text.Trim() == ""
                || txtPosition.Text == "-")
                return;

            // Xử lý sự kiện khi một trong 10 nút được nhấn
            Button clickedButton = sender as Button;
            if (clickedButton != null)
            {
                //MessageBox.Show($"Bạn đã nhấn nút: {clickedButton.Name}");


                //cboError.SelectedIndex = 0;
                int indexError = Convert.ToInt32(clickedButton.Name.Substring(8));
                //string errorText = clickedButton.Name.Substring(3);
                string errorText = lstError[indexError];

                txtErrorCF.Text = dataSheet[Array.IndexOf(arrMaps, txtPosition.Text) + 2];


                if (errorText == "OK")
                {
                    Button btn = this.FindName("btnT" + txtPosition.Text) as Button;
                    if (btn != null)
                    {
                        dataSheet[Array.IndexOf(arrMaps, txtPosition.Text) + 2] = "OK";
                        txtErrorCF.Text = errorText;
                        ChangeColorBTN(btn, true);
                        return;
                    }
                }
                else
                {
                    Button btn = this.FindName("btnT" + txtPosition.Text) as Button;
                    if (btn != null)
                    {
                        dataSheet[Array.IndexOf(arrMaps, txtPosition.Text) + 2] = errorText;
                        txtErrorCF.Text = errorText;
                        ChangeColorBTN(btn, false);
                        return;
                    }

                }
            }
        }
    }
}

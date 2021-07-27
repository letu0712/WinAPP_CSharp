using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
//Khai báo thư viện Firebase
using FireSharp;
using FireSharp.Config;
using FireSharp.Interfaces;
using FireSharp.Response;
//Khai báo thư viện Excel
using Spire.Xls;

namespace IOT_PROJECT
{
    public partial class Form1 : Form
    {
        //Khai báo kết nối đến Firebase
        IFirebaseConfig config = new FirebaseConfig
        {
            AuthSecret = "0BhT5D8K0spFBs44Vnc75FsqgI9QQ9Mi3R2iWbrd",
            BasePath = "https://iot-agriculture-9fe2e-default-rtdb.firebaseio.com/"
        };
        IFirebaseClient client;
        public Form1()
        {
            InitializeComponent();
        }
        //Đăng ký
        private async void button1_Click(object sender, EventArgs e)
        {
            client = new FirebaseClient(config);
            TaiKhoan taikhoan = new TaiKhoan();
            if (textBox2.Text == textBox3.Text)
            {
                taikhoan.tendn = textBox1.Text;
                taikhoan.matkhau = textBox2.Text;
                taikhoan.sdt = textBox4.Text;
                taikhoan.id = textBox5.Text;
                SetResponse resp1 = await client.SetAsync("Thông tin tài khoản/" + taikhoan.tendn, taikhoan);
                label11.Text = "Đăng ký thành công";
                label11.ForeColor = Color.Blue;
            }
            else
            {
                label11.Text = "Đăng ký thất bại";
                label11.ForeColor = Color.Red;
            }
        }
        //ID khu vực quản lý
        public string IDqli;

        //Đăng nhập
        private async void button2_Click(object sender, EventArgs e)
        {
            client = new FirebaseClient(config);
            try
            {
                FirebaseResponse resp = await client.GetAsync("Thông tin tài khoản/" + textBox7.Text);
                TaiKhoan result = resp.ResultAs<TaiKhoan>();
                if (textBox6.Text == result.matkhau)
                {
                    IDqli = result.id;
                    MessageBox.Show("Đăng nhập thành công");
                    label14.Text = "Khu vực " + IDqli;

                    ThongTin();
                }
                else
                {
                    label12.Text = "Đăng nhập thất bại";
                    label12.ForeColor = Color.Red;
                }
            }
            catch 
            {
                label12.ForeColor = Color.Red;
                label12.Text = "Tên đăng nhập không tồn tại";
            }
        }
        //Tạo dữ liệu DataTable
        System.Data.DataTable dataTable = new System.Data.DataTable();
        void Initial_FlowerDataTable_Excel_Chart()
        {
            //tạo các tên trường trong datatable
            dataTable.Columns.Add("ID");
            dataTable.Columns.Add("Khu vực");
            dataTable.Columns.Add("Ánh sáng");
            dataTable.Columns.Add("Thời gian");
            dataGridView1.DataSource = dataTable;

            //Tạo workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Giám sát nông nghiệp";
            //Truyền các header từ datagridview sang worksheet
            for (int i = 0; i < dataGridView1.Columns.Count; i++)
            {
                sheet.Range[Convert.ToChar('A' + i) + "1"].Text = dataGridView1.Columns[i].HeaderText;

            }
            workbook.SaveToFile(@"D:\Thông tin giám sát Flower.xlsx");

        }

        void Initial_VegetableDataTable_Excel_Chart()
        {
            //tạo các tên trường trong datatable
            dataTable.Columns.Add("ID");
            dataTable.Columns.Add("Khu vực");
            dataTable.Columns.Add("Nhiệt độ");
            dataTable.Columns.Add("Độ ẩm");
            dataTable.Columns.Add("Thời gian");
            dataGridView1.DataSource = dataTable;

            //Tạo workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Giám sát nông nghiệp";
            //Truyền các header từ datagridview sang worksheet
            for (int i = 0; i < dataGridView1.Columns.Count; i++)
            {
                sheet.Range[Convert.ToChar('A' + i) + "1"].Text = dataGridView1.Columns[i].HeaderText;

            }
            workbook.SaveToFile(@"D:\Thông tin giám sát Vegetable.xlsx");

        }

        void Initial_AllDataTable_Excel_Chart()
        {
            //tạo các tên trường trong datatable
            dataTable.Columns.Add("ID");
            dataTable.Columns.Add("Khu vực");
            dataTable.Columns.Add("Ánh sáng");
            dataTable.Columns.Add("Nhiệt độ");
            dataTable.Columns.Add("Độ ẩm");
            dataTable.Columns.Add("Thời gian");
            dataGridView1.DataSource = dataTable;

            //Tạo workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Giám sát nông nghiệp";
            //Truyền các header từ datagridview sang worksheet
            for (int i = 0; i < dataGridView1.Columns.Count; i++)
            {
                sheet.Range[Convert.ToChar('A' + i) + "1"].Text = dataGridView1.Columns[i].HeaderText;

            }
            workbook.SaveToFile(@"D:\Thông tin giám sát tất cả khu vực.xlsx");

        }

        int rangeRow;

        //Điền dữ liệu Flower vào Excel
        async void LoadDTGridViewFlower()
        {
            //Reset lại datagridview
            dataTable.Rows.Clear();
            int r = 0;
            while (true)
            {
                if (r == 1)      //Vì có 1 hàng dữ liệu nên break
                {
                    break;
                }
                r++;
                //Lấy dữ liệu Light
                FirebaseResponse resp = await client.GetAsync("Area/Flowers");
                Flower flower = resp.ResultAs<Flower>();
                DataRow row = dataTable.NewRow();
                row["ID"] = IDqli;
                row["Khu vực"] = "Flower";
                row["Ánh sáng"] = flower.light;
                row["Thời gian"] = DateTime.Now;
                dataTable.Rows.Add(row);

                //Load dữ liệu vào file Excel
                LoadExcelFlower();

                //Load dữ liệu vào biểu đồ
                LoadChartFlower();
            }
        }

        async void LoadDTGridViewVegetable()
        {
            //Reset lại datagridview
            dataTable.Rows.Clear();
            int r = 0;
            while (true)
            {
                if (r == 1)      //Vì có 1 hàng dữ liệu nên break
                {
                    break;
                }
                r++;
                //Lấy dữ liệu Light
                FirebaseResponse resp = await client.GetAsync("Area/Vegetables");
                Vegetable vegetable = resp.ResultAs<Vegetable>();
                DataRow row = dataTable.NewRow();
                row["ID"] = IDqli;
                row["Khu vực"] = "Vegetable";
                row["Nhiệt độ"] = vegetable.temperature;
                row["Độ ẩm"] = vegetable.humidity;
                row["Thời gian"] = DateTime.Now;
                dataTable.Rows.Add(row);

                //Load dữ liệu vào file Excel
                LoadExcelVegetable();

                //Load dữ liệu vào biểu đồ
                LoadChartVegetable();
            }
        }

        async void LoadDTGridViewAll()
        {
            dataTable.Rows.Clear();
            int r = 0;
            while (true)
            {
                if (r == 2)      //Vì có 1 hàng dữ liệu nên break
                {
                    break;
                }               
                //Lấy dữ liệu Light
                FirebaseResponse resp1 = await client.GetAsync("Area/Flowers");
                Flower flower = resp1.ResultAs<Flower>();
                DataRow row1 = dataTable.NewRow();
                row1["ID"] = 1;
                row1["Khu vực"] = "Flower";
                row1["Ánh sáng"] = flower.light;
                row1["Thời gian"] = DateTime.Now;
                dataTable.Rows.Add(row1);

                FirebaseResponse resp2 = await client.GetAsync("Area/Vegetables");
                Vegetable vegetable = resp2.ResultAs<Vegetable>();
                DataRow row2 = dataTable.NewRow();
                row2["ID"] = 2;
                row2["Khu vực"] = "Vegetable";
                row2["Nhiệt độ"] = vegetable.temperature;
                row2["Độ ẩm"] = vegetable.humidity;
                row2["Thời gian"] = DateTime.Now;
                dataTable.Rows.Add(row2);
                r += 2;
            }
            LoadExcelAll();

            LoadChartAll();
        } 
        void LoadExcelFlower()         
        {
        //    LoadDTGridViewFlower();
            //Truyền dữ liệu vào Excel
            Workbook wb = new Workbook();
            wb.LoadFromFile(@"D:\Thông tin giám sát Flower.xlsx");
            Worksheet sheet = wb.Worksheets[0];

            //Kiểm tra dòng trống thì mới điền dữ liệu vào
            for (int i = 1; i < 10000; i++)
            {
                if (sheet.Range["A" + i.ToString()].Text == null)
                {
                    rangeRow = i;
                    break;
                }
            }

            //Điền từng hàng từ datagridview sang bảng Excel
            for (int i = 0; i < dataGridView1.RowCount - 1; i++)    //2 hàng, trừ 1 vì trừ hàng header
            {
                for (int j = 0; j < dataGridView1.ColumnCount; j++)    //4 cột
                {
                    sheet.Range[Convert.ToChar('A' + j) + (rangeRow + i).ToString()].Text = dataGridView1.Rows[i].Cells[j].Value.ToString();
                }
            }
            wb.SaveToFile(@"D:\Thông tin giám sát Flower.xlsx");
        }

        void LoadExcelVegetable()
        {
            Workbook wb = new Workbook();
            wb.LoadFromFile(@"D:\Thông tin giám sát Vegetable.xlsx");
            Worksheet sheet = wb.Worksheets[0];

            //Kiểm tra dòng trống thì mới điền dữ liệu vào
            for (int i = 1; i < 10000; i++)
            {
                if (sheet.Range["A" + i.ToString()].Text == null)
                {
                    rangeRow = i;
                    break;
                }
            }

            //Điền từng hàng từ datagridview sang bảng Excel
            for (int i = 0; i < dataGridView1.RowCount - 1; i++)    //2 hàng, trừ 1 vì trừ hàng header
            {
                for (int j = 0; j < dataGridView1.ColumnCount; j++)    //4 cột
                {
                    sheet.Range[Convert.ToChar('A' + j) + (rangeRow + i).ToString()].Text = dataGridView1.Rows[i].Cells[j].Value.ToString();
                }
            }
            wb.SaveToFile(@"D:\Thông tin giám sát Vegetable.xlsx");
        }

        void LoadExcelAll()
        {
            Workbook wb = new Workbook();
            wb.LoadFromFile(@"D:\Thông tin giám sát tất cả khu vực.xlsx");
            Worksheet sheet = wb.Worksheets[0];

            //Kiểm tra dòng trống thì mới điền dữ liệu vào
            for (int i = 1; i < 10000; i++)
            {
                if (sheet.Range["A" + i.ToString()].Text == null)
                {
                    rangeRow = i;
                    break;
                }
            }
            //Điền từng hàng từ datagridview sang bảng Excel
            for (int i = 0; i < dataGridView1.RowCount - 1; i++)    //RowCount = 3 hàng, trừ 1 vì trừ hàng header
            {
                for (int j = 0; j < dataGridView1.ColumnCount; j++)    //4 cột
                {
                    sheet.Range[Convert.ToChar('A' + j) + (rangeRow + i).ToString()].Text = dataGridView1.Rows[i].Cells[j].Value.ToString();
                }
            }
            wb.SaveToFile(@"D:\Thông tin giám sát tất cả khu vực.xlsx");

        }
        int columnChart = 0;
        void LoadChartFlower()          //Load biểu đồ khu vực Flower
        {
            if (columnChart % 10 == 0)
            {
                chart1.Series["Ánh sáng (Lux)"].Points.Clear();
            }
            chart1.Series["Ánh sáng (Lux)"].Points.AddXY(dataGridView1.Rows[0].Cells[3].Value.ToString(), Convert.ToInt32(dataGridView1.Rows[0].Cells[2].Value));
            columnChart++;
        }
        void LoadChartVegetable()           //Load biểu đồ khu vực Vegetable
        {
            if (columnChart % 10 == 0)
            {
                chart1.Series["Nhiệt độ (độ C)"].Points.Clear();
                chart1.Series["Độ ẩm (%)"].Points.Clear();
            }
            chart1.Series["Nhiệt độ (độ C)"].Points.AddXY(dataGridView1.Rows[0].Cells[4].Value.ToString(), Convert.ToInt32(dataGridView1.Rows[0].Cells[2].Value));
            chart1.Series["Độ ẩm (%)"].Points.AddXY(dataGridView1.Rows[0].Cells[4].Value.ToString(), Convert.ToInt32(dataGridView1.Rows[0].Cells[3].Value));
            columnChart++;
        }      
        void LoadChartAll()         //Load biểu đồ của tất cả các khu vực 
        {
            if (columnChart % 10 == 0)
            {
                chart1.Series["Ánh sáng (Lux)"].Points.Clear();
            }
            chart1.Series["Ánh sáng (Lux)"].Points.AddXY(dataGridView1.Rows[0].Cells[5].Value.ToString(), Convert.ToInt32(dataGridView1.Rows[0].Cells[2].Value));

            if (columnChart % 10 == 0)
            {
                chart2.Series["Nhiệt độ (độ C)"].Points.Clear();
                chart2.Series["Độ ẩm (%)"].Points.Clear();
            }
            chart2.Series["Nhiệt độ (độ C)"].Points.AddXY(dataGridView1.Rows[1].Cells[5].Value.ToString(), Convert.ToInt32(dataGridView1.Rows[1].Cells[3].Value));
            chart2.Series["Độ ẩm (%)"].Points.AddXY(dataGridView1.Rows[1].Cells[5].Value.ToString(), Convert.ToInt32(dataGridView1.Rows[1].Cells[4].Value));
            columnChart++;
        }

        void ThongTin() 
        {
            //Nếu ID = 1 là khu vực Flower
            if(IDqli == "1")
            {
                Initial_FlowerDataTable_Excel_Chart();
                //Khởi tạo bộ timer = 5 giây
                timer1.Enabled = true;
                timer1.Interval = 5000;
            }
            //Nếu ID = 2 là khu vực Vegetable
            if (IDqli == "2")
            {
                Initial_VegetableDataTable_Excel_Chart();
                timer2.Enabled = true;
                timer2.Interval = 5000;
            }
            //Nếu ID = 0: Quản lý cả khu vực 1 và 2
            if(IDqli == "0")
            {
                Initial_AllDataTable_Excel_Chart();
                timer3.Enabled = true;
                timer3.Interval = 5000;
            }    
        }
        private void timer1_Tick(object sender, EventArgs e)
        {
            LoadDTGridViewFlower();
        }

        private void timer2_Tick(object sender, EventArgs e)
        {
            LoadDTGridViewVegetable();
        }

        private void timer3_Tick(object sender, EventArgs e)
        {
            LoadDTGridViewAll();
        }

        //TAB điều khiển
        private async void button3_Click(object sender, EventArgs e)
        {
            client = new FirebaseClient(config);

            Control control = new Control();
            control.ctrl = "On";
            SetResponse resp1 = await client.SetAsync("Control", control.ctrl);
        }

        private async void button4_Click(object sender, EventArgs e)
        {
            client = new FirebaseClient(config);

            Control control = new Control();
            control.ctrl = "Off";
            SetResponse resp1 = await client.SetAsync("Control", control.ctrl);
        }
    }
}

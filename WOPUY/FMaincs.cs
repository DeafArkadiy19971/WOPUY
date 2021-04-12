using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using Excel = Microsoft.Office.Interop.Excel;

namespace WOPUY
{
    public partial class FMaincs : Form
    {
        public FMaincs()
        {
            InitializeComponent();
            ConnectionString = "Data Source=DESKTOP-4M90N4D;Initial Catalog=MetopDeaf;Integrated Security=True;Connect Timeout=30;Encrypt=False;TrustServerCertificate=False;ApplicationIntent=ReadWrite;MultiSubnetFailover=False";
            conn(ConnectionString, PCK, dataGridView1);
            conn(ConnectionString, Prepod, dataGridView2);
            conn(ConnectionString, VidMetod, dataGridView3);
            conn(ConnectionString, Metod, dataGridView4);

            conn2(ConnectionString, PCK, comboBox1, "Название пцк", "Код пцк");
            conn2(ConnectionString, Prepod, comboBox2, "ФИО преподавателя", "Код преподавателя");
            conn2(ConnectionString, VidMetod, comboBox3, "Название вида", "Код вид метод разработок");
        }
        public string ConnectionString = "";

        string PCK = "SELECT KodPCK AS [Код пцк], NamePCK AS [Название пцк], PredsedPCK AS [ФИО председателя ПЦК] FROM PCK";
        string Prepod = "SELECT Prepod.KodPrepod AS [Код преподавателя], Prepod.FIOPrepod AS [ФИО преподавателя], Prepod.Dolgnost AS Должность, Prepod.DataPriem AS [Дата приема], Prepod.DataUvol AS [Дата увольнения],  Prepod.DataObuch AS[Дата последнего обучения], Prepod.KodPCK AS[Код пцк] FROM PCK INNER JOIN Prepod ON PCK.KodPCK = Prepod.KodPCK";
        string VidMetod = "SELECT KodVidMetod AS [Код вид метод разработок], NameVid AS [Название вида] FROM VidMetod";
        string Metod = "SELECT Metod.KodMetod AS [Код метод разработки], Metod.NameMetod AS [Название метод], Metod.DateUtv AS [Дата утверждения], Metod.Opis AS [Описание], Metod.KodPrepod AS [Код преподавателя], Metod.KodVidMetod AS[Код вид метод разработок] FROM Metod INNER JOIN Prepod ON Metod.KodPrepod = Prepod.KodPrepod INNER JOIN  VidMetod ON Metod.KodVidMetod = VidMetod.KodVidMetod";

        private void FMaincs_Load(object sender, EventArgs e)
        {

        }
        public void conn(string CS, string cmdT, DataGridView dgv)
        {
            //создание экземпляра адаптера

            SqlDataAdapter Adapter = new SqlDataAdapter(cmdT, CS);
            // сздание обьекта  DataSet (набор данных)
            DataSet ds = new DataSet();
            // Заполнение таблицы набора данных DataSet
            Adapter.Fill(ds, "Table");
            // Связыаем источник данных компонета dataGridView на форме, с таблицей
            dgv.DataSource = ds.Tables["Table"].DefaultView;
        }


        private void button6_Click(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {
            //Новое подключение
            SqlConnection connect = new SqlConnection();
            connect.ConnectionString = ConnectionString;
            //Теперь можно устанавливать соединение, вызыва метод Open обьекта
            connect.Open();
            //Создаем новый экземпляр SQLCommand
            SqlCommand cmd = connect.CreateCommand();
            //Определяем тип SQLCommand=StoredProcedure
            cmd.CommandType = CommandType.StoredProcedure;
            //определяем имя вызывемой процедуры
            cmd.CommandText = "[T1]";
            //Создаем параметр
            //аналогично для все остальных параметров
            cmd.Parameters.Add("@NamePCK", SqlDbType.Char, 170);
            cmd.Parameters["@NamePCK"].Value = textBox1.Text;

            cmd.Parameters.Add("@PredsedPCK", SqlDbType.Char, 150);
            cmd.Parameters["@PredsedPCK"].Value = textBox2.Text;

            //Выполнение хранимой процедуры-добавление записи
            cmd.ExecuteNonQuery();
            //вывод сообщения
            MessageBox.Show("Запись изменеа!");
            //обновление записей в таблице в daataGridview
            conn(ConnectionString, PCK, dataGridView1);
        }

        private void button1_Click(object sender, EventArgs e)
        {

        }

        private void button8_Click(object sender, EventArgs e)
        {
            //Новое подключение
            SqlConnection connect = new SqlConnection();
            connect.ConnectionString = ConnectionString;
            //Теперь можно устанавливать соединение, вызыва метод Open обьекта
            connect.Open();
            //Создаем новый экземпляр SQLCommand
            SqlCommand cmd = connect.CreateCommand();
            //Определяем тип SQLCommand=StoredProcedure
            cmd.CommandType = CommandType.StoredProcedure;
            //определяем имя вызывемой процедуры
            cmd.CommandText = "[T2]";
            //Создаем параметр
            //аналогично для все остальных параметров
            cmd.Parameters.Add("@FIOPrepod", SqlDbType.Char, 170);
            cmd.Parameters["@FIOPrepod"].Value = textBox3.Text;

            cmd.Parameters.Add("@Dolgnost", SqlDbType.Char, 170);
            cmd.Parameters["@Dolgnost"].Value = textBox4.Text;

            cmd.Parameters.Add("@DataPriem", SqlDbType.Date);
            cmd.Parameters["@DataPriem"].Value = dateTimePicker1.Value;

            cmd.Parameters.Add("@DataUvol", SqlDbType.Date);
            cmd.Parameters["@DataUvol"].Value = dateTimePicker2.Value;

            cmd.Parameters.Add("@DataObuch", SqlDbType.Date);
            cmd.Parameters["@DataObuch"].Value = dateTimePicker3.Value;

            cmd.Parameters.Add("@KodPCK", SqlDbType.Int);
            cmd.Parameters["@KodPCK"].Value = comboBox1.SelectedValue;

            //Выполнение хранимой процедуры-добавление записи
            cmd.ExecuteNonQuery();
            //вывод сообщения
            MessageBox.Show("Запись изменеа!");
            //обновление записей в таблице в daataGridview
            conn(ConnectionString, Prepod, dataGridView2);
        }

        public void conn2(string CS, string cmdT, ComboBox CB, string field1, string field2)
        {
            //создание экземпляра адаптера
            SqlDataAdapter adapter = new SqlDataAdapter(cmdT, CS);
            //создание обекта DataSet (набор данных)
            DataSet ds = new DataSet();
            adapter.Fill(ds, "Table");
            //привязка comboBox к таблице БД
            CB.DataSource = ds.Tables["Table"];

            CB.DisplayMember = field1; //установка отбражаемого в списке поля
            CB.ValueMember = field2; //установка клчевого поля
        }

        private void button12_Click(object sender, EventArgs e)
        {

            //Новое подключение
            SqlConnection connect = new SqlConnection();
            connect.ConnectionString = ConnectionString;
            //Теперь можно устанавливать соединение, вызыва метод Open обьекта
            connect.Open();
            //Создаем новый экземпляр SQLCommand
            SqlCommand cmd = connect.CreateCommand();
            //Определяем тип SQLCommand=StoredProcedure
            cmd.CommandType = CommandType.StoredProcedure;
            //определяем имя вызывемой процедуры
            cmd.CommandText = "[T3]";
            //Создаем параметр
            //аналогично для все остальных параметров
            cmd.Parameters.Add("@NameVid", SqlDbType.Char, 150);
            cmd.Parameters["@NameVid"].Value = textBox5.Text;
            //Выполнение хранимой процедуры-добавление записи
            cmd.ExecuteNonQuery();
            //вывод сообщения
            MessageBox.Show("Запись изменеа!");
            //обновление записей в таблице в daataGridview
            conn(ConnectionString, VidMetod, dataGridView3);
        }

        private void button16_Click(object sender, EventArgs e)
        {
            //Новое подключение
            SqlConnection connect = new SqlConnection();
            connect.ConnectionString = ConnectionString;
            //Теперь можно устанавливать соединение, вызыва метод Open обьекта
            connect.Open();
            //Создаем новый экземпляр SQLCommand
            SqlCommand cmd = connect.CreateCommand();
            //Определяем тип SQLCommand=StoredProcedure
            cmd.CommandType = CommandType.StoredProcedure;
            //определяем имя вызывемой процедуры
            cmd.CommandText = "[T4]";
            //Создаем параметр
            //аналогично для все остальных параметров
            cmd.Parameters.Add("@NameMetod", SqlDbType.Char, 150);
            cmd.Parameters["@NameMetod"].Value = textBox6.Text;

            cmd.Parameters.Add("@DateUtv", SqlDbType.Date);
            cmd.Parameters["@DateUtv"].Value = dateTimePicker4.Value;

            cmd.Parameters.Add("@Opis", SqlDbType.Char, 170);
            cmd.Parameters["@Opis"].Value = textBox7.Text;

            cmd.Parameters.Add("@KodPrepod", SqlDbType.Int);
            cmd.Parameters["@KodPrepod"].Value = comboBox2.SelectedValue;

            cmd.Parameters.Add("@KodVidMetod", SqlDbType.Int);
            cmd.Parameters["@KodVidMetod"].Value = comboBox3.SelectedValue;

            //Выполнение хранимой процедуры-добавление записи
            cmd.ExecuteNonQuery();
            //вывод сообщения
            MessageBox.Show("Запись изменеа!");
            //обновление записей в таблице в daataGridview
            conn(ConnectionString, Metod, dataGridView4);
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            //объявление переменных
            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();

            Microsoft.Office.Interop.Excel.Workbook ExcelWorkBook;

            Microsoft.Office.Interop.Excel.Worksheet ExcelWorkSheet;

            Microsoft.Office.Interop.Excel.Range ExcelCells;
            //создание новой раббочеий книги
            ExcelWorkBook = ExcelApp.Workbooks.Add(System.Reflection.Missing.Value);
            //создание лист 
            ExcelWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelWorkBook.Worksheets.get_Item(1);
            //сформируем даты из dateTimePicker3 и dateTimPicker4
            string d8 = dateTimePicker5.Value.Day.ToString() + "." + dateTimePicker5.Value.Month.ToString() + "." +
                dateTimePicker5.Value.Year.ToString();
            string d9 = dateTimePicker6.Value.Day.ToString() + "." + dateTimePicker6.Value.Month.ToString() + "." +
                dateTimePicker6.Value.Year.ToString();
            //Вывод заголовка отчета
            ExcelApp.Cells[1, 1] = "Отчет о книговыдаче за период с " + d8 + " по " + d9;
            //вывод Заголвков полей таблицы
            for (int i = 0; i < dataGridView5.ColumnCount; i++)
            {
                ExcelApp.Cells[2, i + 1] = dataGridView5.Columns[i].HeaderCell.Value;
            }
            //вывод содержимого dataGridview

            for (int i = 0; i < dataGridView5.Rows.Count; i++)
            {
                for (int j = 0; j < dataGridView5.ColumnCount; j++)
                {
                    ExcelApp.Cells[i + 3, j + 1] = dataGridView5.Rows[i].Cells[j].Value;
                }
            }
            int istr = dataGridView5.Rows.Count + 1;
            //форматирование ячеек Excel
            //автоподбор ширины столбцов
            ExcelCells = ExcelApp.Range[ExcelWorkSheet.Columns[1], ExcelWorkSheet.Columns[8]];
            ExcelCells.EntireColumn.AutoFit();
            //горизонтальное выравнивание
            ExcelCells = ExcelApp.Range[ExcelWorkSheet.Columns[2], ExcelWorkSheet.Columns[8]];
            ExcelCells.HorizontalAlignment = Excel.Constants.xlLast;
            //обрамление линиями
            ExcelCells = ExcelApp.Range[ExcelWorkSheet.Cells[2, 1], ExcelWorkSheet.Cells[istr, 8]];
            ExcelCells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            ExcelCells.Borders.Weight = Excel.XlBorderWeight.xlThin;
            //Открывем Excel
            ExcelApp.Visible = true;
            ExcelApp.UserControl = true;
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {

        }
    }
}

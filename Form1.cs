using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using MySql.Data.MySqlClient;
using System.Data.OleDb;


namespace Bridge
{
    public partial class Form1 : Form
    {
        private string server = "localhost";
        private string userid = "root";
        private string password = "w/rwLke0er=k";
        private string database = "bridge";
        private string CurrentDatabase = "";
        public Form1()
        {
            InitializeComponent();
        }

        private void 退出ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("确定退出吗？", "title", MessageBoxButtons.OKCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == System.Windows.Forms.DialogResult.OK)
                Application.Exit();
        }

        private void groupBox3_MouseHover(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "勾选显示相应列，取消勾选隐藏相应列";
        }

        private void groupBox1_MouseHover(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "只有登录后才可以操作";
        }

        //Form1初始化
        private void Form1_Load(object sender, EventArgs e)
        {
            //comboBox1初始化
            ArrayList mylist = new ArrayList();
            var rdr = GetData("show tables");
            int i = 1;
            while (rdr.Read())
            {
                
                mylist.Add(new DictionaryEntry(i.ToString(), rdr.GetString(0).ToString()));
                i++;
            }
            comboBox1.DataSource = mylist;           
            comboBox1.DisplayMember = "Value";
            comboBox1.ValueMember = "Key";         
        }

        public MySqlDataReader GetData(string str)
        {
            string cs = connectionInformation();
            MySqlConnection conn = new MySqlConnection(cs);
            conn.Open();
            MySqlCommand cmd = new MySqlCommand(str, conn);
            MySqlDataReader rdr = cmd.ExecuteReader();                
            return rdr;
        }
        //生成mysql数据库连接字符串
        public string connectionInformation()
        {
            string cs = "SERVER=" + server + ";" + "DATABASE=" +
        database + ";" + "UID=" + userid + ";" + "PASSWORD=" + password + ";";
            return cs;
        }

        //当comboBox内容选中后进行dataGridView数据呈现
        private void comboBox1_SelectedValueChanged(object sender, EventArgs e)
        {
            
            if (comboBox1.Text.ToString() != "System.Collections.DictionaryEntry") 
            {
                string cs = connectionInformation();
                MySqlConnection conn = new MySqlConnection(cs);
                //变量赋值便于之后使用
                CurrentDatabase = comboBox1.Text.ToString();
                string sqlStr = "select * from " + CurrentDatabase + ";"; //
                var adapter = new MySqlDataAdapter(sqlStr, conn);
                var set = new DataSet(); //数据集、本地微型数据库可以存储多张表。
                adapter.Fill(set, "测试");
                dataGridView1.DataSource = set;
                dataGridView1.DataMember = "测试";
                initCheckBox();
                CurrentDatabase = comboBox1.Text.ToString();
            }         
        }
        //保存当前数据为excel文件
        private void 保存为ExcelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Thread generateMyWord = new Thread(new ThreadStart(OutputAsExcelFile));
            //设置为后台线程
            generateMyWord.IsBackground = true;
            //开启线程
            generateMyWord.Start();         
        }

        private void OutputAsExcelFile()
        {
            //将datagridView中的数据导出到一张表中
            DataTable tempTable = this.exporeDataToTable(this.dataGridView1);
            //导出信息到Excel表
            Microsoft.Office.Interop.Excel.ApplicationClass myExcel;
            Microsoft.Office.Interop.Excel.Workbooks myWorkBooks;
            Microsoft.Office.Interop.Excel.Workbook myWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet myWorkSheet;
            char myColumns;
            Microsoft.Office.Interop.Excel.Range myRange;
            object[,] myData = new object[500, 35];
            int i, j;//j代表行,i代表列
            myExcel = new Microsoft.Office.Interop.Excel.ApplicationClass();
            //显示EXCEL
            myExcel.Visible = true;
            if (myExcel == null)
            {
                MessageBox.Show("本地Excel程序无法启动!请检查您的Microsoft Office正确安装并能正常使用", "提示");
                return;
            }
            myWorkBooks = myExcel.Workbooks;
            myWorkBook = myWorkBooks.Add(System.Reflection.Missing.Value);
            myWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)myWorkBook.Worksheets[1];
            myColumns = (char)(tempTable.Columns.Count + 64);//设置列
            myRange = myWorkSheet.get_Range("A4", myColumns.ToString() + "5");//设置列宽
            int count = 0;
            //设置列名
            foreach (DataColumn myNewColumn in tempTable.Columns)
            {
                myData[0, count] = myNewColumn.ColumnName;
                count = count + 1;
            }
            //输出datagridview中的数据记录并放在一个二维数组中
            j = 1;
            foreach (DataRow myRow in tempTable.Rows)//循环行
            {
                for (i = 0; i < tempTable.Columns.Count; i++)//循环列
                {
                    myData[j, i] = myRow[i].ToString();
                }
                j++;
            }
            //将二维数组中的数据写到Excel中
            myRange = myRange.get_Resize(tempTable.Rows.Count + 1, tempTable.Columns.Count);//创建列和行
            myRange.Value2 = myData;
            myRange.EntireColumn.AutoFit();
            myRange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
            //删除前方空白三行
            Microsoft.Office.Interop.Excel.Range deleteRng = (Microsoft.Office.Interop.Excel.Range)myWorkSheet.Rows[1, System.Type.Missing];
            deleteRng.Delete(Microsoft.Office.Interop.Excel.XlDeleteShiftDirection.xlShiftUp);
            Microsoft.Office.Interop.Excel.Range deleteRng1 = (Microsoft.Office.Interop.Excel.Range)myWorkSheet.Rows[1, System.Type.Missing];
            deleteRng1.Delete(Microsoft.Office.Interop.Excel.XlDeleteShiftDirection.xlShiftUp);
            Microsoft.Office.Interop.Excel.Range deleteRng2 = (Microsoft.Office.Interop.Excel.Range)myWorkSheet.Rows[1, System.Type.Missing];
            deleteRng2.Delete(Microsoft.Office.Interop.Excel.XlDeleteShiftDirection.xlShiftUp);

        }
        private DataTable exporeDataToTable(DataGridView dataGridView)
        {
            //将datagridview中的数据导入到表中
            DataTable tempTable = new DataTable("tempTable");
            //定义一个模板表，专门用来获取列名
            DataTable modelTable = new DataTable("ModelTable");
            //创建列
            for (int column = 0; column < dataGridView.Columns.Count; column++)
            {
                //可见的列才显示出来
                if (dataGridView.Columns[column].Visible == true)
                {
                    DataColumn tempColumn = new DataColumn(dataGridView.Columns[column].HeaderText, typeof(string));
                    tempTable.Columns.Add(tempColumn);
                    DataColumn modelColumn = new DataColumn(dataGridView.Columns[column].Name, typeof(string));
                    modelTable.Columns.Add(modelColumn);
                }
            }
            //添加datagridview中行的数据到表
            for (int row = 0; row < dataGridView.Rows.Count; row++)
            {
                if (dataGridView.Rows[row].Visible == false)
                {
                    continue;
                }
                DataRow tempRow = tempTable.NewRow();
                for (int i = 0; i < tempTable.Columns.Count; i++)
                {
                    tempRow[i] = dataGridView.Rows[row].Cells[modelTable.Columns[i].ColumnName].Value;
                }
                tempTable.Rows.Add(tempRow);
            }
            
            return tempTable;
        }
        //导入excel
        private void 导入ExcelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string filePath = string.Empty;
            string fileExt = string.Empty;
            string filename = string.Empty; 
            OpenFileDialog file = new OpenFileDialog(); //打开文件选择框
            if (file.ShowDialog() == System.Windows.Forms.DialogResult.OK) 
            {
                filePath = file.FileName; //文件目录
                fileExt = Path.GetExtension(filePath); //文件后缀
                filename = System.IO.Path.GetFileNameWithoutExtension(file.FileName);//文件名
                
                ArrayList mylist = new ArrayList();
                var rdr = GetData("show tables");
                int i = 1;
                int flag = 1;
                while (rdr.Read())
                {
                    mylist.Add(new DictionaryEntry(i.ToString(), rdr.GetString(0).ToString()));
                    if (rdr.GetString(0).ToString()==filename)
                    {
                        flag = 0;
                    }
                    
                    i++;
                }
                //如果数据库内没有这个文件名的表就创建表名为文件名的表
                if (flag==1)
                {
                    string cs = connectionInformation();
                    MySqlConnection conn = new MySqlConnection(cs);
                    conn.Open();
                    string sqlStr = "create table "+filename + "(`排序` int(10),`构件` char(255),`构件类型` char(255),`单位` char(255),`工程量` char(255),`单位成本` char(255),`使用年限` char(255),`常见损伤类型` char(255),`常见维护方式` char(255),`常用检测方法` char(255));";
                    var cmd = new MySqlCommand(sqlStr, conn);
                    cmd.ExecuteNonQuery();
                    conn.Close();
                    mylist.Add(new DictionaryEntry(i.ToString(), filename));

                }
                comboBox1.DataSource = mylist;
                comboBox1.DisplayMember = "Value";
                comboBox1.ValueMember = "Key";
                if (fileExt.CompareTo(".xls") == 0 || fileExt.CompareTo(".xlsx") == 0)
                {
                    try
                    {
                        DataTable dtExcel = new DataTable();
                        dtExcel = ReadExcel(filePath, fileExt); //读取excel
                        
                        CurrentDatabase = filename;
                        BulkInsert(connectionInformation(), dtExcel, CurrentDatabase);//导入excel到数据库
                        string cs = connectionInformation();
                        MySqlConnection conn = new MySqlConnection(cs);
                        //从数据库读取展示
                        string sqlStr = "select * from " + filename + ";"; //
                        var adapter = new MySqlDataAdapter(sqlStr, conn);
                        var set = new DataSet(); 
                        adapter.Fill(set, "测试");

                        dataGridView1.DataSource = set;
                        dataGridView1.DataMember = "测试";
                        initCheckBox();
                        CurrentDatabase = comboBox1.Text.ToString();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message.ToString());
                    }
                }
                else
                {
                    MessageBox.Show("Please choose .xls or .xlsx file only.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error); //custom messageBox to show error  
                }
            }
        }
    
        public DataTable ReadExcel(string fileName, string fileExt)
        {
            string conn = string.Empty;
            DataTable dtexcel = new DataTable();
            if (fileExt.CompareTo(".xls") == 0)
                conn = @"provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fileName + ";Extended Properties='Excel 8.0;HRD=Yes;IMEX=1';"; //for below excel 2007  
            else
                conn = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties='Excel 12.0;HDR=YES';"; //for above excel 2007  
            using (OleDbConnection con = new OleDbConnection(conn))
            {
                try
                {
                    OleDbDataAdapter oleAdpt = new OleDbDataAdapter("select * from [Sheet1$]", con); //here we read data from sheet1  
                    oleAdpt.Fill(dtexcel); //fill excel data into dataTable  
                }
                catch { }
            }
            
            return dtexcel;
        }
        public int BulkInsert(string connectionString, DataTable table, string TableName)
        {
            //if (string.IsNullOrEmpty(table.TableName)) throw new Exception(TableName);
            if (table.Rows.Count == 0) return 0;
            int insertCount = 0;
            string tmpPath = Path.GetTempFileName();
            string csv = DataTableToCsv(table);
            File.WriteAllText(tmpPath, csv);
            // MySqlTransaction tran = null;  

            using (MySqlConnection conn = new MySqlConnection(connectionString))
            {

                try
                {
                    conn.Open();
                    //tran = conn.BeginTransaction();  
                    MySqlBulkLoader bulk = new MySqlBulkLoader(conn)
                    {
                        FieldTerminator = ",",
                        FieldQuotationCharacter = '"',
                        EscapeCharacter = '"',
                        LineTerminator = "\r\n",
                        FileName = tmpPath,
                        NumberOfLinesToSkip = 0,
                        TableName = TableName,

                    };
                    //bulk.Columns.AddRange(table.Columns.Cast<DataColumn>().Select(colum => colum.ColumnName).ToArray());
                    insertCount = bulk.Load();
                    //tran.Commit();
                }
                catch (MySqlException ex)
                {
                    // if (tran != null) tran.Rollback();  
                    throw ex;
                }
            }
            File.Delete(tmpPath);
            return insertCount;
        }

        private static string DataTableToCsv(DataTable table)
        {
            //以半角逗号（即,）作分隔符，列为空也要表达其存在。  
            //列内容如存在半角逗号（即,）则用半角引号（即""）将该字段值包含起来。  
            //列内容如存在半角引号（即"）则应替换成半角双引号（""）转义，并用半角引号（即""）将该字段值包含起来。  
            StringBuilder sb = new StringBuilder();
            DataColumn colum;
            foreach (DataRow row in table.Rows)
            {
                for (int i = 0; i < table.Columns.Count; i++)
                {
                    colum = table.Columns[i];
                    if (i != 0) sb.Append(",");
                    if (colum.DataType == typeof(string) && row[colum].ToString().Contains(","))
                    {
                        sb.Append("\"" + row[colum].ToString().Replace("\"", "\"\"") + "\"");
                    }
                    else sb.Append(row[colum].ToString());
                }
                sb.AppendLine();
            }
            return sb.ToString();
        }

        

        //checkbox 初始化
        public void initCheckBox()
        {
            
            checkBox1.Checked = true;
            checkBox2.Checked = true;
            checkBox3.Checked = true;
            checkBox4.Checked = true;
            checkBox5.Checked = true;
            checkBox6.Checked = true;
        }
        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true)
            {
                dataGridView1.Columns[4].Visible = true;
            }
            else
            {
                dataGridView1.Columns[4].Visible = false;
            }
        }
        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked == true)
            {
                dataGridView1.Columns[5].Visible = true;
            }
            else
            {
                dataGridView1.Columns[5].Visible = false;
            }
        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox3.Checked == true)
            {
                dataGridView1.Columns[6].Visible = true;
            }
            else
            {
                dataGridView1.Columns[6].Visible = false;
            }
        }

        private void checkBox4_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox4.Checked == true)
            {
                dataGridView1.Columns[7].Visible = true;
            }
            else
            {
                dataGridView1.Columns[7].Visible = false;
            }
        }

        private void checkBox5_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox5.Checked == true)
            {
                dataGridView1.Columns[8].Visible = true;
            }
            else
            {
                dataGridView1.Columns[8].Visible = false;
            }
        }

        private void checkBox6_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox6.Checked == true)
            {
                dataGridView1.Columns[9].Visible = true;
            }
            else
            {
                dataGridView1.Columns[9].Visible = false;
            }
        }
        //读取Form3的文本并insert到数据库
        public void Add(List<string> input)
        {
            string cs = connectionInformation();
            MySqlConnection conn = new MySqlConnection(cs);

            string sqlStr = "insert into " + CurrentDatabase + " values(" + input[0] + ",'" + input[1] + "','" + input[2]+ "','" + input[3]+ "','" + input[4]+ "','" + input[5]+ "','" + input[6]+ "','" + input[7]+ "','" + input[8]+ "','" + input[9]+"');";
            var adapter = new MySqlDataAdapter(sqlStr, conn);
            var set1 = new DataSet();
            adapter.Fill(set1, "测试1");
            string sqlStr1 = "select * from " + CurrentDatabase + ";"; //
            var adapter1 = new MySqlDataAdapter(sqlStr1, conn);
            var set = new DataSet(); //数据集、本地微型数据库可以存储多张表。
            adapter1.Fill(set, "测试");

            dataGridView1.DataSource = set;
            dataGridView1.DataMember = "测试";
            initCheckBox();

        }

        private void button1_Click(object sender, EventArgs e)
        {
            var frm = new Form3();
            frm.Show(this);
        }
        

        private void button2_Click(object sender, EventArgs e)
        {
            string cs = connectionInformation();
            MySqlConnection conn = new MySqlConnection(cs);
            conn.Open();
            var row = dataGridView1.SelectedRows[0];
            var index = row.Index;
            var id = dataGridView1.Rows[index].Cells[0].Value.ToString();
            var 构件 = dataGridView1.Rows[index].Cells[1].Value.ToString();
            var 构件类型 = dataGridView1.Rows[index].Cells[2].Value.ToString();
            var 单位 = dataGridView1.Rows[index].Cells[3].Value.ToString();
            var 工程量 = dataGridView1.Rows[index].Cells[4].Value.ToString();
            var 单位成本 = dataGridView1.Rows[index].Cells[5].Value.ToString();
            var 使用年限 = dataGridView1.Rows[index].Cells[6].Value.ToString();
            var 常见损伤类型 = dataGridView1.Rows[index].Cells[7].Value.ToString();
            var 常见维护方式 = dataGridView1.Rows[index].Cells[8].Value.ToString();
            var 常用检测方法 = dataGridView1.Rows[index].Cells[9].Value.ToString();
            string sqlStr = "delete from " + CurrentDatabase + " where 排序=" + id + " and 构件='" + 构件 + "' and 构件类型='" + 构件类型 + "' and 单位='" + 单位 + "' and 工程量='" + 工程量 + "' and 单位成本='" + 单位成本 + "' and 使用年限='" + 使用年限 + "' and 常见损伤类型='" + 常见损伤类型 + "' and 常见维护方式='" + 常见维护方式 + "' and 常用检测方法='" + 常用检测方法+"';";
            var cmd = new MySqlCommand(sqlStr, conn);

            cmd.ExecuteNonQuery();
       
            string sqlStr1 = "select * from " + CurrentDatabase + ";"; //
            var adapter1 = new MySqlDataAdapter(sqlStr1, conn);
            var set = new DataSet(); 
            adapter1.Fill(set, "测试");

            dataGridView1.DataSource = set;
            dataGridView1.DataMember = "测试";
            initCheckBox();
        }
        //搜索框textBox1输入回车之后读取数据展示到comboBox2
        //支持模糊搜索
        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode==Keys.Enter)
            {
                ArrayList mylist = new ArrayList();
                var rdr = search(textBox1.Text.ToString());
                int i = 1;
                while (rdr.Read())
                {

                    mylist.Add(new DictionaryEntry(i.ToString(), rdr.GetString(0).ToString()));
                    i++;
                }
                comboBox2.DataSource = mylist;
                comboBox2.DisplayMember = "Value";
                comboBox2.ValueMember = "Key";
            }
        }
        public MySqlDataReader search(string input) 
        {
            string cs = connectionInformation();
            MySqlConnection conn = new MySqlConnection(cs);
            conn.Open();
            string str= "show tables like '%" + input + "%';";
            MySqlCommand cmd = new MySqlCommand(str, conn);
            MySqlDataReader rdr = cmd.ExecuteReader();
            return rdr;
        }

        private void comboBox2_SelectedValueChanged(object sender, EventArgs e)
        {
            if (comboBox2.Text.ToString() != "System.Collections.DictionaryEntry") 
            {
                string cs = connectionInformation();
                MySqlConnection conn = new MySqlConnection(cs);

                CurrentDatabase = comboBox2.Text.ToString();
                string sqlStr = "select * from " + CurrentDatabase + ";"; //
                var adapter = new MySqlDataAdapter(sqlStr, conn);
                var set = new DataSet(); //数据集、本地微型数据库可以存储多张表。
                adapter.Fill(set, "测试12");

                dataGridView1.DataSource = set;
                dataGridView1.DataMember = "测试12";
                initCheckBox();
                CurrentDatabase = comboBox2.Text.ToString();
            }
                
        }
    }
}

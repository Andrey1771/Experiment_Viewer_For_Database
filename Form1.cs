using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;

namespace Lab_7
{
    public partial class Form1 : Form
    {
        OleDbConnection cn = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\Visual Projects\Db_Labs\Lab_7\Lab_6.mdb;");
        OleDbCommand cmd = new OleDbCommand();
        string nameTable;
        string accessIsShitPrimarykey;

        public Form1()
        {
            InitializeComponent();
        }
        private void button1_Click(object sender, EventArgs e)
        {
            cn.ConnectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\Visual Projects\Db_Labs\Lab_7\Lab_6.mdb";//Db_Labs
            cmd.Connection = cn;
            textBox1.BackColor = Color.FromArgb(255, 255, 255);
            try
            {
                cn.Open(); // установка соединения
            }
            catch
            {
                textBox1.Text = "Ошибка подключения!";
                textBox1.BackColor = Color.FromArgb(255, 100, 100);
            }
            finally
            {
                try
                {
                    cmd.CommandText = "SELECT * FROM [" + textBox1.Text + "];";
                    OleDbDataReader rd = cmd.ExecuteReader();//	.schema
                    List<string[]> data = new List<string[]>();
                    // если запрос вернул результат
                    if (rd.HasRows)
                    {
                        clearDataGridView(dataGridView1);
                        nameTable = textBox1.Text;

                        for (int i = 0; i < rd.FieldCount; ++i)
                            dataGridView1.Columns.Add(rd.GetName(i), rd.GetName(i));

                        while (rd.Read())
                        {

                            // ... добавляем в список содержимое столбца «Фамилия»
                            data.Add(new string[rd.FieldCount]);

                            for (int i = 0; i < rd.FieldCount; ++i)
                                data[data.Count - 1][i] = rd[i].ToString();

                        }
                        foreach (string[] s in data)
                            dataGridView1.Rows.Add(s);
                    }
                }
                catch
                {
                    textBox1.Text = "Ошибка подключения!";
                    textBox1.BackColor = Color.FromArgb(255, 100, 100);
                }
                finally
                {
                    cn.Close(); // закрытие соединения с БД
                }
            }
        }

        private void updateDataRows()
        {
            cn.ConnectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\Visual Projects\Db_Labs\Lab_7\Lab_6.mdb";
            cmd.Connection = cn;
            textBox1.BackColor = Color.FromArgb(255, 255, 255);
            cn.Open(); // установка соединения
            try
            {

                cmd.CommandText = "SELECT * FROM [" + nameTable + "];";
                OleDbDataReader rd = cmd.ExecuteReader();//	.schema
                List<string[]> data = new List<string[]>();
                // если запрос вернул результат
                if (rd.HasRows)
                {
                    clearDataGridView(dataGridView1);

                    for (int i = 0; i < rd.FieldCount; ++i)
                        dataGridView1.Columns.Add(rd.GetName(i), rd.GetName(i));

                    while (rd.Read())
                    {

                        // ... добавляем в список содержимое столбца «Фамилия»
                        data.Add(new string[rd.FieldCount]);

                        for (int i = 0; i < rd.FieldCount; ++i)
                            data[data.Count - 1][i] = rd[i].ToString();

                    }
                    foreach (string[] s in data)
                        dataGridView1.Rows.Add(s);
                }
            }
            catch
            {
                textBox1.Text = "Ошибка подключения!";
                textBox1.BackColor = Color.FromArgb(255, 100, 100);
            }
            finally
            {
                cn.Close(); // закрытие соединения с БД
            }
        }
        private void clearDataGridView(DataGridView view)
        {
            view.Rows.Clear();
            view.Columns.Clear();
        }

        private void dataGridView1_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            accessIsShitPrimarykey = dataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString();
        }
        void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            cn.ConnectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\Visual Projects\Db_Labs\Lab_7\Lab_6.mdb";
            cmd.Connection = cn;
            textBox2.BackColor = Color.FromArgb(255, 255, 255);
            cn.Open();
            try
            {
                string omg = "UPDATE [" + nameTable + "] SET [" + dataGridView1.Columns[e.ColumnIndex].Name + "] = \"" + dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value + "\" WHERE [" + dataGridView1.Columns[0].Name + "] = " + accessIsShitPrimarykey + ";";
                cmd.CommandText = omg;
                cmd.ExecuteNonQuery();
            }
            catch
            {
                textBox2.Text = "Ошибка изменения!";
                textBox2.BackColor = Color.FromArgb(255, 100, 100);
            }
            finally
            {
                cn.Close(); // закрытие соединения с БД
                textBox1.Text = nameTable;
                //updateDataRows();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            cn.ConnectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\Visual Projects\Db_Labs\Lab_7\Lab_6.mdb";
            cmd.Connection = cn;
            textBox1.BackColor = Color.FromArgb(255, 255, 255);
            cn.Open();
            try
            {
                string omg = "DELETE * FROM [" + nameTable + "] WHERE [" + dataGridView1.Columns[0].Name + "] = " + textBox2.Text + ";";
                cmd.CommandText = omg;
                cmd.ExecuteNonQuery();
            }
            catch
            {
                textBox2.Text = "Ошибка удаления!";
                textBox2.BackColor = Color.FromArgb(255, 100, 100);
            }
            finally
            {
                cn.Close(); // закрытие соединения с БД
                textBox1.Text = nameTable;
                updateDataRows();
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            cn.ConnectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\Visual Projects\Db_Labs\Lab_7\Lab_6.mdb";
            cmd.Connection = cn;
            textBox1.BackColor = Color.FromArgb(255, 255, 255);
            cn.Open();
            try
            {

                string omg = "INSERT INTO [" + nameTable + "] ([" + dataGridView1.Columns[0].Name + "])  VALUES (" + dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[0].Value.ToString() + " + 1);";
                cmd.CommandText = omg;
                cmd.ExecuteNonQuery();
            }
            catch
            {
                textBox2.Text = "Ошибка Добавления!";
                textBox2.BackColor = Color.FromArgb(255, 100, 100);
            }
            finally
            {
                cn.Close(); // закрытие соединения с БД
                textBox1.Text = nameTable;
                updateDataRows();


            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            cn.ConnectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;" +
@"Data Source=E:\Visual Projects\Db_Labs\Lab_7\Lab_6.mdb;" +
@"Jet OLEDB:Create System Database=true;" + // разрешение на доступ
@"Jet OLEDB:System database=C:\Users\Andrey\AppData\Roaming\Microsoft\Access\System.mdw";

            cmd.Connection = cn;

            textBox1.BackColor = Color.FromArgb(255, 255, 255);
            cn.Open(); // установка соединения
            try
            {

                textBox2.Text = cn.GetSchema().TableName;
                cmd.CommandText = "SELECT * FROM [" + textBox1.Text + "];";
                OleDbDataReader rd = cmd.ExecuteReader();//	.schema
                List<string[]> data = new List<string[]>();
                // если запрос вернул результат
                if (rd.HasRows)
                {
                    clearDataGridView(dataGridView1);
                    nameTable = textBox1.Text;

                    for (int i = 0; i < rd.FieldCount; ++i)
                        dataGridView1.Columns.Add(rd.GetName(i), rd.GetName(i));

                    while (rd.Read())
                    {

                        // ... добавляем в список содержимое столбца «Фамилия»
                        data.Add(new string[rd.FieldCount]);

                        for (int i = 0; i < rd.FieldCount; ++i)
                            data[data.Count - 1][i] = rd[i].ToString();

                    }
                    foreach (string[] s in data)
                        dataGridView1.Rows.Add(s);
                }
            }
            catch
            {
                textBox1.Text = "Ошибка подключения!";
                textBox1.BackColor = Color.FromArgb(255, 100, 100);
            }
            finally
            {
                cn.Close(); // закрытие соединения с БД
            }
        }

        private void button5_Click(object sender, EventArgs e)// запрос без параметров
        {
            cn.ConnectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\Visual Projects\Db_Labs\Lab_7\Lab_6.mdb";
            cmd.Connection = cn;
            textBox1.BackColor = Color.FromArgb(255, 255, 255);
            cn.Open(); // установка соединения
            try
            {

                switch (comboBox1.SelectedIndex)
                {
                    case (0):
                        {
                            cmd.CommandText = " SELECT * FROM Магазин WHERE [Номер магазина] < 2;";
                            break;
                        }
                    case (1):
                        {
                            cmd.CommandText = "SELECT * FROM Сотрудник WHERE [Пол семейное положение] = (SELECT FIRST([Пол семейное положение]) FROM Сотрудник);";
                            break;

                        }
                    case (2):
                        {
                            cmd.CommandText = "SELECT * FROM Товар WHERE [Дата поставки] BETWEEN #14/03/2020# AND #18/03/2020#;";
                            break;
                        }
                    default:
                        {
                            cmd.CommandText = "SELECT * FROM [" + nameTable + "];";
                            break;
                        }


                }
                OleDbDataReader rd = cmd.ExecuteReader();//	.schema
                List<string[]> data = new List<string[]>();
                // если запрос вернул результат
                if (rd.HasRows)
                {
                    clearDataGridView(dataGridView1);

                    for (int i = 0; i < rd.FieldCount; ++i)
                        dataGridView1.Columns.Add(rd.GetName(i), rd.GetName(i));

                    while (rd.Read())
                    {

                        // ... добавляем в список содержимое столбца «Фамилия»
                        data.Add(new string[rd.FieldCount]);

                        for (int i = 0; i < rd.FieldCount; ++i)
                            data[data.Count - 1][i] = rd[i].ToString();

                    }
                    foreach (string[] s in data)
                        dataGridView1.Rows.Add(s);
                }
            }
            catch
            {
                textBox1.Text = "Ошибка подключения!";
                textBox1.BackColor = Color.FromArgb(255, 100, 100);
            }
            finally
            {
                cn.Close(); // закрытие соединения с БД
            }

        }

        private void button6_Click(object sender, EventArgs e)// запрос с параметрами
        {
            cn.ConnectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\Visual Projects\Db_Labs\Lab_7\Lab_6.mdb";
            cmd.Connection = cn;
            textBox1.BackColor = Color.FromArgb(255, 255, 255);
            cn.Open(); // установка соединения
            try
            {

                switch (comboBox2.SelectedIndex)
                {
                    case (0):
                        {
                            cmd.CommandText = "SELECT * FROM Сотрудник WHERE [Фамилия сотрудника] LIKE \"%" + textBox3.Text + "%\";";
                            break;
                        }
                    case (1):
                        {
                            cmd.CommandText = "SELECT * FROM Магазин WHERE [Номер магазина] > " + textBox3.Text.Split(';')[0] + " AND [Название магазина] NOT LIKE \"%" + textBox3.Text.Split(';')[1] + "%\";";
                            break;

                        }
                    case (2):
                        {
                            cmd.CommandText = "SELECT * FROM Товар WHERE [Дата поставки] LIKE #" + textBox3.Text + "#;";//#14/03/2020#
                            break;
                        }
                    default:
                        {
                            cmd.CommandText = "SELECT * FROM [" + nameTable + "];";
                            break;
                        }


                }
                OleDbDataReader rd = cmd.ExecuteReader();//	.schema
                List<string[]> data = new List<string[]>();
                // если запрос вернул результат
                if (rd.HasRows)
                {
                    clearDataGridView(dataGridView1);

                    for (int i = 0; i < rd.FieldCount; ++i)
                        dataGridView1.Columns.Add(rd.GetName(i), rd.GetName(i));

                    while (rd.Read())
                    {

                        // ... добавляем в список содержимое столбца «Фамилия»
                        data.Add(new string[rd.FieldCount]);

                        for (int i = 0; i < rd.FieldCount; ++i)
                            data[data.Count - 1][i] = rd[i].ToString();

                    }
                    foreach (string[] s in data)
                        dataGridView1.Rows.Add(s);
                }
            }
            catch
            {
                textBox1.Text = "Ошибка подключения!";
                textBox1.BackColor = Color.FromArgb(255, 100, 100);
            }
            finally
            {
                cn.Close(); // закрытие соединения с БД
            }
        }
        private void button7_Click(object sender, EventArgs e)// Удаление
        {
            cn.ConnectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\Visual Projects\Db_Labs\Lab_7\Lab_6.mdb";
            cmd.Connection = cn;
            textBox1.BackColor = Color.FromArgb(255, 255, 255);
            cn.Open(); // установка соединения
            try
            {
                string omg = "DELETE * FROM Магазин WHERE [Название магазина] = \"" + textBox4.Text + "\";";
                cmd.CommandText = omg;

                OleDbDataReader rd = cmd.ExecuteReader();//	.schema
                List<string[]> data = new List<string[]>();
                // если запрос вернул результат
                if (rd.HasRows)
                {
                    clearDataGridView(dataGridView1);

                    for (int i = 0; i < rd.FieldCount; ++i)
                        dataGridView1.Columns.Add(rd.GetName(i), rd.GetName(i));

                    while (rd.Read())
                    {

                        // ... добавляем в список содержимое столбца «Фамилия»
                        data.Add(new string[rd.FieldCount]);

                        for (int i = 0; i < rd.FieldCount; ++i)
                            data[data.Count - 1][i] = rd[i].ToString();

                    }
                    foreach (string[] s in data)
                        dataGridView1.Rows.Add(s);
                }
            }
            catch
            {
                textBox1.Text = "Ошибка подключения!";
                textBox1.BackColor = Color.FromArgb(255, 100, 100);
            }
            finally
            {
                cn.Close(); // закрытие соединения с БД
            }
        }

        private void button8_Click(object sender, EventArgs e)// добавление
        {
            cn.ConnectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\Visual Projects\Db_Labs\Lab_7\Lab_6.mdb";
            cmd.Connection = cn;
            textBox1.BackColor = Color.FromArgb(255, 255, 255);
            cn.Open();
            try
            {

                string omg = "INSERT INTO [Сотрудник] ([Табельный номер сотрудника], [Номер магазина], [Номер отдела], [Фамилия сотрудника])  VALUES (1, 1, 1, \"" + textBox5.Text + "\");";
                cmd.CommandText = omg;
                cmd.ExecuteNonQuery();
            }
            catch
            {
                textBox2.Text = "Ошибка Добавления!";
                textBox2.BackColor = Color.FromArgb(255, 100, 100);
            }
            finally
            {
                cn.Close(); // закрытие соединения с БД
                textBox1.Text = nameTable;
                updateDataRows();


            }
        }

        private void button9_Click(object sender, EventArgs e)// обновление
        {
            cn.ConnectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\Visual Projects\Db_Labs\Lab_7\Lab_6.mdb";
            cmd.Connection = cn;
            textBox2.BackColor = Color.FromArgb(255, 255, 255);
            cn.Open();
            try
            {
                string omg = "UPDATE Сотрудник SET [Фамилия сотрудника] = \"" + textBox7.Text + "\" WHERE [Табельный номер сотрудника] = " + textBox6.Text + ";";
                cmd.CommandText = omg;
                cmd.ExecuteNonQuery();
            }
            catch
            {
                textBox2.Text = "Ошибка изменения!";
                textBox2.BackColor = Color.FromArgb(255, 100, 100);
            }
            finally
            {
                cn.Close(); // закрытие соединения с БД
                textBox1.Text = nameTable;
                //updateDataRows();
            }
        }
    }
}


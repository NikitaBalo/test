using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.CompilerServices;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace table_add
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private void Enable(bool i)
        {
            if (i)
            {
                maskedTextBox1.Enabled = true;
                maskedTextBox1.Text = null;
            }
            else
            {
                maskedTextBox1.Enabled = false;
                maskedTextBox1.Text = null;
            }
        }
        public string py;
        public string po;
        public string pyn;
        private void AddData()
        {
            Data.Add(1, textBox1.Text);
            Data.Add(2, comboBox2.Text);
            Data.Add(3, comboBox1.Text);
            Data.Add(4, py);
            Data.Add(5, po);
            Data.Add(6, pyn);
            Data.Add(7, comboBox3.Text);
            Data.Add(8, textBox2.Text);
            Data.Add(9, textBox3.Text);
            Data.Add(10, dateTimePicker1.Text);
            Data.Add(11, "");
            Data.Add(12, comboBox4.Text.Split('\t')[0]);
            Data.Add(13, comboBox4.Text.Split('\t')[1]);
            Data.Add(14, textBox7.Text);
            Data.Add(15, textBox11.Text);
            Data.Add(16, maskedTextBox2.Text);
            Data.Add(17, maskedTextBox3.Text);
            Data.Add(18, "");
            Data.Add(19, textBox10.Text);
            Data.Add(20, textBox8.Text);
            Data.Add(21, textBox9.Text);
            Data.Add(22, dateTimePicker2.Text);
            Data.Add(23, comboBox6.Text);
            Data.Add(24, maskedTextBox1.Text);
            Data.Add(25, comboBox7.Text.Split('\t')[1]);
        }
        private string WithOutSpace(string str)
        {
            string str0 = "";
            foreach (char c in str)
            {
                if (c != ' ')
                    str0 += c;
            }
            return str0;
        }
        public bool ContainOnlyString(string str)
        {
            if (Regex.IsMatch(WithOutSpace(str), @"\d+"))
            {
                return false;
            }
            else
            {
                return true;
            }
        }
        private string WithOutInt(string str)
        {
            string str0 = "";
            foreach (char c in WithOutSpace(str))
            {
                if (c < '0' || c > '9')
                    str0 += c;
            }

            return str0;
        }
        public bool ContainOnlyInt(string str)
        {
            if (Regex.IsMatch(str, @"\D+"))
            {
                return false;
            }
            else
            {
                return true;
            }
        }
        private string WithOutString(string str)
        {
            string str0 = "";
            foreach (char c in str)
            {
                if (c < '0' || c > '9') { }
                else
                {
                    str0 += c;
                }

            }
            return str0;
        }
        private int k = 0;
        private bool CheckForFill()
        {
            if (String.IsNullOrEmpty(comboBox1.Text))
            {
                comboBox1.BackColor = Color.Red;
                ++k;
            }
            else
            {
                comboBox1.BackColor = Color.White;
            }
            if (String.IsNullOrEmpty(comboBox2.Text))
            {
                comboBox2.BackColor = Color.Red;
                ++k;
            }
            else
            {
                comboBox2.BackColor = Color.White;
            }
            if (String.IsNullOrEmpty(comboBox3.Text))
            {
                comboBox3.BackColor = Color.Red;
                ++k;
            }
            else
            {
                comboBox3.BackColor = Color.White;
            }
            if (String.IsNullOrEmpty(comboBox4.Text))
            {
                comboBox4.BackColor = Color.Red;
                ++k;
            }
            else
            {
                comboBox4.BackColor = Color.White;
            }
            if (String.IsNullOrEmpty(comboBox6.Text))
            {
                comboBox6.BackColor = Color.Red;
                ++k;
            }
            else
            {
                comboBox6.BackColor = Color.White;
            }
            if (String.IsNullOrEmpty(comboBox7.Text))
            {
                comboBox7.BackColor = Color.Red;
                ++k;
            }
            else
            {
                comboBox7.BackColor = Color.White;

            }
            if (String.IsNullOrEmpty(textBox1.Text) | !ContainOnlyString(textBox1.Text))
            {
                textBox1.BackColor = Color.Red;
                ++k;
            }
            else
            {
                textBox1.BackColor = Color.White;
            }
            if (!ContainOnlyString(textBox9.Text))
            {
                textBox9.BackColor = Color.Red;
                ++k;
            }
            else
            {
                textBox9.BackColor = Color.White;
            }
            if (String.IsNullOrEmpty(textBox10.Text) | !ContainOnlyString(textBox10.Text))
            {
                textBox10.BackColor = Color.Red;
                ++k;
            }
            else
            {
                textBox10.BackColor = Color.White;
            }
            if (String.IsNullOrEmpty(textBox8.Text) | !ContainOnlyString(textBox8.Text))
            {
                textBox8.BackColor = Color.Red;
                ++k;
            }
            else
            {
                textBox8.BackColor = Color.White;
            }
            if (String.IsNullOrEmpty(textBox2.Text) | !ContainOnlyInt(textBox2.Text))
            {
                textBox2.BackColor = Color.Red;
                ++k;
            }
            else
            {
                textBox2.BackColor = Color.White;

            }
            if (String.IsNullOrEmpty(textBox3.Text) | !ContainOnlyInt(textBox3.Text))
            {
                textBox3.BackColor = Color.Red;
                ++k;
            }
            else
            {
                textBox3.BackColor = Color.White;

            }
            if (String.IsNullOrEmpty(textBox7.Text) | !ContainOnlyString(textBox7.Text))
            {
                textBox7.BackColor = Color.Red;
                ++k;
            }
            else
            {
                textBox7.BackColor = Color.White;
            }
            if (String.IsNullOrEmpty(textBox11.Text) | !ContainOnlyString(textBox11.Text))
            {
                textBox11.BackColor = Color.Red;
                ++k;
            }
            else
            {
                textBox11.BackColor = Color.White;
            }
            if (!maskedTextBox1.MaskCompleted)
            {
                if (comboBox7.Text == "Российская Федерация\t643")
                {
                    maskedTextBox1.BackColor = Color.Red;
                    ++k;
                }
                else
                {
                    maskedTextBox1.BackColor = Color.White;
                }
            }
            else
            {
                maskedTextBox1.BackColor = Color.White;
            }
            if (!maskedTextBox2.MaskCompleted)
            {
                maskedTextBox2.BackColor = Color.Red;
                ++k;
                if (Convert.ToInt32(maskedTextBox2.Text) >= Convert.ToInt32(maskedTextBox3.Text))
                {
                    maskedTextBox2.BackColor = Color.Red;
                    maskedTextBox3.BackColor = Color.Red;
                }
            }
            else
            {
                maskedTextBox2.BackColor = Color.White;
            }
            
            if (!maskedTextBox3.MaskCompleted)
            {
                maskedTextBox3.BackColor = Color.Red;
                ++k;
                if (Convert.ToInt32(maskedTextBox2.Text) >= Convert.ToInt32(maskedTextBox3.Text))
                {
                    maskedTextBox2.BackColor = Color.Red;
                    maskedTextBox3.BackColor = Color.Red;
                }
            }
            else
            {
                maskedTextBox3.BackColor = Color.White;
            }
            if (!ContainOnlyString(textBox1.Text))
            {
                textBox1.BackColor = Color.Red;
                ++k;
            }
            else
            {
                textBox1.BackColor = Color.White;
            }
            if (k == 0)
            {
                return true;
            }
            else
            {
                k = 0;
                return false;
            }
            
        }

        private void comboBox7_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox7.Text == "Российская Федерация\t643")
            {
                Enable(true);
            }
            else
            {
                Enable(false);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (CheckForFill())
            {
                AddData();
                ClearData();
                Form2 f = new Form2();
                f.Show();
                Hide();
            }
            else
            {
                MessageBox.Show("Введите данные или исправьте данные в ячейках!");
            }
        }

        private void comboBox1_TextUpdate(object sender, EventArgs e)
        {
            MessageBox.Show("Пожалуйста выберите значение из списка.");
        }

        private void comboBox2_TextUpdate(object sender, EventArgs e)
        {
            MessageBox.Show("Пожалуйста выберите значение из списка.");
        }

        private void comboBox3_TextUpdate(object sender, EventArgs e)
        {
            MessageBox.Show("Пожалуйста выберите значение из списка.");
        }
        private void comboBox6_TextUpdate(object sender, EventArgs e)
        {
            MessageBox.Show("Пожалуйста выберите значение из списка.");
        }
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

            if (comboBox1.Text == "Оригинал")
            {
                Data.Set(47, true);
                po = "Нет";
                py = "Нет";
                pyn = "Нет";
            }
            if (comboBox1.Text == "Дубликат")
            {
                Data.Set(47, false);
                po = "Нет";
                py = "Да";
                pyn = "Нет";
            }
        }
        private void ClearData()
        {
            comboBox1.Text = null;
            comboBox2.Text = null;
            comboBox3.Text = null;
            comboBox4.Text = null;
            comboBox6.Text = null;
            comboBox7.Text = null;
            dateTimePicker1.Value = DateTime.Now.AddDays(-1);
            dateTimePicker2.Value = DateTime.Now.AddDays(-1);
            textBox1.Text = null;
            textBox2.Text = null;
            textBox3.Text = null;
            maskedTextBox2.Text = null;
            maskedTextBox3.Text = null;
            maskedTextBox1.Text = null;
            textBox8.Text = null;
            textBox9.Text = null;
            textBox10.Text = null;
            maskedTextBox1.Enabled = false;
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            dateTimePicker1.Format = DateTimePickerFormat.Custom;
            dateTimePicker2.Format = DateTimePickerFormat.Custom;
            dateTimePicker1.CustomFormat = "dd.MM.yyyy";
            dateTimePicker2.CustomFormat = "dd.MM.yyyy";
            dateTimePicker1.MaxDate = DateTime.Today;
            dateTimePicker2.MaxDate = DateTime.Today;
            Data.Add(47, " ");

        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                Excel.Application app = new Excel.Application();
                Excel.Workbook book = app.Workbooks.Open(Data.Get1(0).ToString()); ;
                app.Visible = true;
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
        }

        private void comboBox7_TextChanged(object sender, EventArgs e)
        {
            if (comboBox7.Text == "Российская Федерация\t643")
            {
                Enable(true);
            }
            else
            {
                Enable(false);
            }
        }

        private void maskedTextBox1_MouseUp(object sender, MouseEventArgs e)
        {
            if (!maskedTextBox1.MaskCompleted)
            {
                int startPos = this.maskedTextBox1.MaskedTextProvider.FindUnassignedEditPositionFrom(this.maskedTextBox1.MaskedTextProvider.LastAssignedPosition + 1, true);
                this.maskedTextBox1.Select(startPos, 0);
            }
        }

        private void maskedTextBox2_MouseUp(object sender, MouseEventArgs e)
        {
            if (!maskedTextBox2.MaskCompleted)
            {
                int startPos = this.maskedTextBox2.MaskedTextProvider.FindUnassignedEditPositionFrom(this.maskedTextBox2.MaskedTextProvider.LastAssignedPosition + 1, true);
                this.maskedTextBox2.Select(startPos, 0);
            }
        }

        private void maskedTextBox3_MouseUp(object sender, MouseEventArgs e)
        {
            if (!maskedTextBox3.MaskCompleted)
            {
                int startPos = this.maskedTextBox3.MaskedTextProvider.FindUnassignedEditPositionFrom(this.maskedTextBox3.MaskedTextProvider.LastAssignedPosition + 1, true);
                this.maskedTextBox3.Select(startPos, 0);
            }
        }
    }
}

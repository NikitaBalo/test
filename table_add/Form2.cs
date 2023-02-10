using Microsoft.Office.Core;
using otchet_fill;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace table_add
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }

        public char[] c = { '0', '1', '2', '3', '4', '5', '6', '7', '8', '9', };
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
            if (Regex.IsMatch(str, @"\d+", RegexOptions.IgnoreCase))
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
        public string str;
        private void Enable(bool i)
        {
            textBox1.Enabled = i;
            textBox2.Enabled = i;
            maskedTextBox1.Enabled = i;
            maskedTextBox2.Enabled = i;
            textBox5.Enabled = i;
            maskedTextBox3.Enabled = i;
            maskedTextBox4.Enabled = i;
            comboBox5.Enabled = i;
            dateTimePicker1.Enabled = i;
            if (i)
            {
                textBox1.Text = "";
                textBox2.Text = "";
                maskedTextBox1.Text = "";
                maskedTextBox2.Text = "";
                textBox5.Text = "";
                maskedTextBox3.Text = "";
                maskedTextBox4.Text = "";
                comboBox5.Text = "";
                str = dateTimePicker1.Text;
            }
            else
            {
                textBox1.Text = "-";
                textBox2.Text = "-";
                maskedTextBox1.Text = null;
                maskedTextBox2.Text = null;
                textBox5.Text = "-";
                maskedTextBox3.Text = null;
                maskedTextBox4.Text = null;
                comboBox5.Text = "-";
                str = "-";
            }
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
            if (String.IsNullOrEmpty(comboBox5.Text))
            {
                comboBox5.BackColor = Color.Red;
                ++k;
            }
            else
            {
                comboBox5.BackColor = Color.White;
            }
            if (String.IsNullOrEmpty(textBox1.Text) | !ContainOnlyInt(textBox1.Text))
            {
                if (comboBox4.Text == "Да")
                {
                    textBox1.BackColor = Color.Red;
                    ++k;
                }
                else
                {
                    textBox1.BackColor = Color.White;
                }
            }
            else
            {
                textBox1.BackColor = Color.White;
            }
            if (String.IsNullOrEmpty(textBox2.Text) | !ContainOnlyString(textBox2.Text))
            {
                if (comboBox4.Text == "Да")
                {
                    textBox2.BackColor = Color.Red;
                    ++k;
                }
                else
                {
                    textBox2.BackColor = Color.White;
                }
            }
            else
            {
                textBox2.BackColor = Color.White;
            }
            if (!maskedTextBox1.MaskCompleted)
            {
                if (comboBox4.Text == "Да")
                {
                    maskedTextBox1.BackColor = Color.Red;
                    ++k;
                }
                else
                {
                    maskedTextBox1.BackColor = Color.White;
                }
            }
            if (!maskedTextBox2.MaskCompleted)
            {
                if (comboBox4.Text == "Да")
                {
                    maskedTextBox2.BackColor = Color.Red;
                    ++k;
                }
                else
                {
                    maskedTextBox2.BackColor = Color.White;
                }
            }
            if (String.IsNullOrEmpty(textBox5.Text) | !ContainOnlyString(textBox5.Text))
            {
                if (comboBox4.Text == "Да")
                {
                    textBox5.BackColor = Color.Red;
                    ++k;
                }
                else
                {
                    textBox5.BackColor = Color.White;
                }
            }
            else
            {
                textBox5.BackColor = Color.White;
            }
            if (!maskedTextBox3.MaskCompleted)
            {
                if (comboBox4.Text == "Да")
                {
                    maskedTextBox3.BackColor = Color.Red;
                    ++k;
                }
                else
                {
                    maskedTextBox3.BackColor = Color.White;
                }
            }
            if (!maskedTextBox4.MaskCompleted)
            {
                if (comboBox4.Text == "Да")
                {
                    maskedTextBox4.BackColor = Color.Red;
                    ++k;
                }
                else
                {
                    maskedTextBox4.BackColor = Color.White;
                }
            }
            if (String.IsNullOrEmpty(textBox8.Text) | !ContainOnlyString(textBox8.Text))
            {
                if (Data.Get(3).ToString() == "Оригинал")
                {
                    textBox8.BackColor = Color.Red;
                    ++k;
                }
                else
                {
                    textBox8.BackColor = Color.White;
                }
            }
            else
            {
                textBox8.BackColor = Color.White;
            }
            if (String.IsNullOrEmpty(textBox9.Text) | !ContainOnlyInt(textBox9.Text))
            {
                if (Data.Get(3).ToString() == "Оригинал")
                {
                    textBox9.BackColor = Color.Red;
                    ++k;
                }
                else
                {
                    textBox9.BackColor = Color.White;
                }
            }
            else
            {
                textBox9.BackColor = Color.White;
            }
            if (String.IsNullOrEmpty(textBox10.Text) | !ContainOnlyInt(textBox10.Text))
            {
                if (Data.Get(3).ToString() == "Оригинал")
                {
                    textBox10.BackColor = Color.Red;
                    ++k;
                }
                else
                {
                    textBox10.BackColor = Color.White;
                }
            }
            else
            {
                textBox10.BackColor = Color.White;
            }
            if (String.IsNullOrEmpty(textBox12.Text) | !ContainOnlyString(textBox12.Text))
            {
                if (Data.Get(3).ToString() == "Оригинал")
                {
                    textBox12.BackColor = Color.Red;
                    ++k;
                }
                else
                {
                    textBox12.BackColor = Color.White;
                }
            }
            else
            {
                textBox12.BackColor = Color.White;
            }
            if (String.IsNullOrEmpty(textBox13.Text) | !ContainOnlyString(textBox13.Text))
            {
                if (Data.Get(3).ToString() == "Оригинал")
                {
                    textBox13.BackColor = Color.Red;
                    ++k;
                }
                else
                {
                    textBox13.BackColor = Color.White;
                }
            }
            else
            {
                textBox13.BackColor = Color.White;
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
        private void AddData()
        {
            Data.Add(26, comboBox1.Text);
            Data.Add(27, comboBox2.Text);
            Data.Add(28, comboBox3.Text);
            Data.Add(29, comboBox4.Text);
            Data.Add(30, textBox1.Text);
            Data.Add(31, str);
            Data.Add(32, textBox2.Text);
            Data.Add(33, maskedTextBox1.Text);
            Data.Add(34, maskedTextBox2.Text);
            Data.Add(35, textBox5.Text);
            Data.Add(36, maskedTextBox3.Text);
            Data.Add(37, maskedTextBox4.Text);
            Data.Add(38, comboBox5.Text);
            Data.Add(39, textBox8.Text);
            Data.Add(40, textBox9.Text);
            Data.Add(41, textBox10.Text);
            Data.Add(42, textBox11.Text);
            Data.Add(43, dateTimePicker2.Text);
            Data.Add(44, textBox12.Text);
            Data.Add(45, textBox13.Text);
            Data.Add(46, textBox14.Text);
        }

        private void button1_Enter(object sender, EventArgs e)
        { 
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (CheckForFill()) 
            {
                AddData();
                ClearData();
                if (MessageBox.Show("Вы уверены?", "Предупреждение", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1) == DialogResult.Yes)
                {
                    try
                    {
                        using (ExcelHelper helper = new ExcelHelper())
                        {
                            helper.Open(filePath: Data.Get1(0).ToString());
                            int i = helper.FindLastRow();
                            for (int j = 1; j < 47; j++)
                            {
                                if (j != 18 || j != 11 || j != 42)
                                {
                                    helper.Set(row: i, col: j, data: Data.Get(j));
                                }
                                if (j == 18)
                                {
                                    helper.Set(row: i, col: j, data: Convert.ToInt32(Data.Get(17))- Convert.ToInt32(Data.Get(16)));
                                }
                                if (j == 11)
                                {
                                    helper.Set(row: i, col: j, data: i-1);
                                }
                                if (j == 42)
                                {
                                    if (String.IsNullOrEmpty(Data.Get(42).ToString()))
                                        helper.Set(row: i, col: j, data: i-1);
                                    else
                                        helper.Set(row: i, col: j, data: Data.Get(j));
                                }
                            }
                            helper.Save();
                        }
                    }
                    catch (Exception ex) { MessageBox.Show(ex.Message); }
                    finally
                    {
                        Data.Clear();
                    }

                    if (MessageBox.Show("Продолжить работу?", "", MessageBoxButtons.YesNo, MessageBoxIcon.None, MessageBoxDefaultButton.Button1) == DialogResult.Yes)
                    {
                        Form1 f = new Form1();;
                        f.Show();
                        Hide();
                    }
                    else
                    {
                        Close();
                        Application.Exit();
                    }
                }
            }
            else
            {
                MessageBox.Show("Введите данные или исправьте данные в ячейках!");
            }
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            dateTimePicker1.Format = DateTimePickerFormat.Custom;
            dateTimePicker2.Format = DateTimePickerFormat.Custom;
            dateTimePicker1.CustomFormat = "dd.MM.yyyy";
            dateTimePicker2.CustomFormat = "dd.MM.yyyy";
            dateTimePicker1.MaxDate = DateTime.Today;
            dateTimePicker2.MaxDate = DateTime.Today;
            if (Convert.ToBoolean(Data.Get(47)))
            {
                textBox8.Text = Data.Get(1).ToString();textBox8.Enabled = false;
                textBox9.Text = Data.Get(8).ToString(); textBox9.Enabled = false;
                textBox10.Text = Data.Get(9).ToString(); textBox10.Enabled = false;
                textBox11.Enabled = false;
                dateTimePicker2.Text = Data.Get(10).ToString();dateTimePicker2.Enabled = false;
                textBox12.Text = Data.Get(19).ToString(); textBox12.Enabled = false;
                textBox13.Text = Data.Get(20).ToString(); textBox13.Enabled = false;
                textBox14.Text = Data.Get(21).ToString(); textBox14.Enabled = false;
            }
        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox4.Text == "Да")
            {
                Enable(true);
            }
            else
            {
                Enable(false);
            }
        }

        private void comboBox4_TextUpdate(object sender, EventArgs e)
        {
            MessageBox.Show("Пожалуйста выберите значение из списка.");
        }

        private void comboBox3_TextUpdate(object sender, EventArgs e)
        {
            MessageBox.Show("Пожалуйста выберите значение из списка.");
        }

        private void comboBox2_TextUpdate(object sender, EventArgs e)
        {
            MessageBox.Show("Пожалуйста выберите значение из списка.");
        }

        private void comboBox1_TextUpdate(object sender, EventArgs e)
        {
            MessageBox.Show("Пожалуйста выберите значение из списка.");
        }
        private void ClearData()
        {
            comboBox1.Text = null;
            comboBox2.Text = null;
            comboBox3.Text = null;
            comboBox4.Text = null;
            comboBox5.Text = null;
            textBox1.Text = null;
            textBox2.Text = null;
            maskedTextBox1.Text = null;
            maskedTextBox2.Text = null;
            textBox5.Text = null;
            maskedTextBox3.Text = null;
            maskedTextBox4.Text = null;
            textBox8.Text = null;
            textBox9.Text = null;
            textBox10.Text = null;
            textBox11.Text = null;
            textBox12.Text = null;
            textBox13.Text = null;
            textBox14.Text = null;
            dateTimePicker1.Value = DateTime.Now.AddDays(-1);
            dateTimePicker2.Value = DateTime.Now.AddDays(-1);
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
        private void maskedTextBox4_MouseUp(object sender, MouseEventArgs e)
        {
            if (!maskedTextBox4.MaskCompleted)
            {
                int startPos = this.maskedTextBox4.MaskedTextProvider.FindUnassignedEditPositionFrom(this.maskedTextBox4.MaskedTextProvider.LastAssignedPosition + 1, true);
                this.maskedTextBox4.Select(startPos, 0);
            }
        }      
    }
}

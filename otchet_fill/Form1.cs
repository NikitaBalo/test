using System.Runtime.CompilerServices;
using Excel = Microsoft.Office.Interop.Excel;
namespace otchet_fill
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        public bool T = true;
        public bool F = false;
        private void Form1_Load(object sender, EventArgs e)
        {
            label1.Visible = T; label2.Visible = T; label3.Visible = T; label7.Visible = T; label8.Visible = T; label9.Visible = T; label10.Visible = T; dateTimePicker1.Visible = T; comboBox1.Visible = T; comboBox2.Visible = T; comboBox6.Visible = T; textBox1.Visible = T; textBox2.Visible = T; textBox3.Visible = T; button2.Visible = T;
            dateTimePicker1.Format = DateTimePickerFormat.Custom;
            dateTimePicker2.Format = DateTimePickerFormat.Custom;
            dateTimePicker3.Format = DateTimePickerFormat.Custom;
            dateTimePicker4.Format = DateTimePickerFormat.Custom;
            dateTimePicker1.MaxDate = DateTime.Today;
            dateTimePicker2.MaxDate = DateTime.Today;
            dateTimePicker3.MaxDate = DateTime.Today;
            dateTimePicker4.MaxDate = DateTime.Today;
            dateTimePicker1.CustomFormat = "dd.MM.yyyy";
            dateTimePicker2.CustomFormat = "dd.MM.yyyy";
            dateTimePicker3.CustomFormat = "dd.MM.yyyy";
            dateTimePicker4.CustomFormat = "dd.MM.yyyy";
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void label7_Click(object sender, EventArgs e)
        {

        }

        private void label8_Click(object sender, EventArgs e)
        {

        }

        private void label9_Click(object sender, EventArgs e)
        {

        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void comboBox6_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox2.Text == "Дубликат")
            {
                comboBox3.Visible = T; comboBox4.Visible = T; comboBox5.Visible = T; label4.Visible = T; label5.Visible = T; label6.Visible = T;
                comboBox3.Text = "Да"; comboBox4.Text = "Нет"; comboBox5.Text = "Нет";
            }
            else
            {
                if (comboBox2.Text == "Оригинал")
                {
                    comboBox3.Text = "Нет"; comboBox4.Text = "Нет"; comboBox5.Text = "Нет";
                }
                comboBox3.Visible = F; comboBox4.Visible = F; comboBox5.Visible = F; label4.Visible = F; label5.Visible = F; label6.Visible = F;
            }
        }


        private void label11_Click(object sender, EventArgs e)
        {

        }

        private void label17_Click(object sender, EventArgs e)
        {

        }

        private void comboBox14_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox14.Text == "Да")
            {
                textBox12.Visible = T; dateTimePicker3.Visible = T; textBox15.Visible = T; textBox16.Visible = T; textBox17.Visible = T; label28.Visible = T; label32.Visible = T; label33.Visible = T; label34.Visible = T; label35.Visible = T;
                textBox12.Text = ""; textBox15.Text = ""; textBox16.Text = ""; textBox17.Text = "";
            }
            else
            {
                textBox12.Text = "-"; textBox15.Text = "-"; textBox16.Text = "-"; textBox17.Text = "-";
                textBox12.Visible = F; dateTimePicker3.Visible = F; textBox15.Visible = F; textBox16.Visible = F; textBox17.Visible = F; label28.Visible = F; label32.Visible = F; label33.Visible = F; label34.Visible = F; label35.Visible = F;
            }
        }

        private void comboBox9_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (!String.IsNullOrEmpty(textBox1.Text) && !String.IsNullOrEmpty(textBox2.Text) && !String.IsNullOrEmpty(textBox3.Text) && !String.IsNullOrEmpty(comboBox1.Text) && !String.IsNullOrEmpty(comboBox2.Text) && !String.IsNullOrEmpty(comboBox3.Text) && !String.IsNullOrEmpty(comboBox4.Text) && !String.IsNullOrEmpty(comboBox5.Text) && !String.IsNullOrEmpty(comboBox6.Text))
            {
                label1.Visible = F; label2.Visible = F; label3.Visible = F; label7.Visible = F; label8.Visible = F; label9.Visible = F; label10.Visible = F; dateTimePicker1.Visible = F; comboBox1.Visible = F; comboBox2.Visible = F; comboBox6.Visible = F; textBox1.Visible = F; textBox2.Visible = F; textBox3.Visible = F; button2.Visible = F; comboBox3.Visible = F; comboBox4.Visible = F; comboBox5.Visible = F; label4.Visible = F; label5.Visible = F; label6.Visible = F;
                comboBox7.Visible = T; comboBox8.Visible = T; label11.Visible = T; label12.Visible = T; label13.Visible = T; label14.Visible = T; label15.Visible = T; label16.Visible = T; textBox4.Visible = T; textBox5.Visible = T; textBox6.Visible = T; textBox7.Visible = T; button3.Visible = T;
            }
            else
            {
                MessageBox.Show("Введите все значения!");
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (!String.IsNullOrEmpty(comboBox7.Text) && !String.IsNullOrEmpty(comboBox8.Text) && !String.IsNullOrEmpty(textBox7.Text) && !String.IsNullOrEmpty(textBox6.Text) && !String.IsNullOrEmpty(textBox5.Text) && !String.IsNullOrEmpty(textBox4.Text))
            {
                comboBox7.Visible = F; comboBox8.Visible = F; label11.Visible = F; label12.Visible = F; label13.Visible = F; label14.Visible = F; label15.Visible = F; label16.Visible = F; textBox4.Visible = F; textBox5.Visible = F; textBox6.Visible = F; textBox7.Visible = F; button3.Visible = F;
                label17.Visible = T; label19.Visible = T; label20.Visible = T; label21.Visible = T; label22.Visible = T; label23.Visible = T; comboBox9.Visible = T; comboBox10.Visible = T; textBox8.Visible = T; textBox9.Visible = T; textBox10.Visible = T; dateTimePicker4.Visible = T; button4.Visible = T;
            }
            else
            {
                MessageBox.Show("Введите все значения!");
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (comboBox9.Text == "Российская Федерация")
            {
                if (!String.IsNullOrEmpty(textBox8.Text) && !String.IsNullOrEmpty(textBox9.Text) && !String.IsNullOrEmpty(textBox10.Text) && !String.IsNullOrEmpty(comboBox10.Text) && !String.IsNullOrEmpty(comboBox9.Text) && !String.IsNullOrEmpty(textBox11.Text))
                {
                    textBox11.Visible = F; label24.Visible = F; label17.Visible = F; label19.Visible = F; label20.Visible = F; label21.Visible = F; label22.Visible = F; label23.Visible = F; comboBox9.Visible = F; comboBox10.Visible = F; textBox8.Visible = F; textBox9.Visible = F; textBox10.Visible = F; dateTimePicker4.Visible = F; button4.Visible = F;
                    label25.Visible = T; label29.Visible = T; label30.Visible = T; label31.Visible = T; comboBox11.Visible = T; comboBox12.Visible = T; comboBox13.Visible = T; comboBox14.Visible = T; button5.Visible = T;
                }
                else
                {
                    MessageBox.Show("Введите все значения!");
                }
            }
            else
            {
                if (!String.IsNullOrEmpty(textBox8.Text) && !String.IsNullOrEmpty(textBox9.Text) && !String.IsNullOrEmpty(textBox10.Text) && !String.IsNullOrEmpty(comboBox10.Text) && !String.IsNullOrEmpty(comboBox9.Text))
                {
                    textBox11.Text = "-";
                    textBox11.Visible = F; label24.Visible = F; label17.Visible = F; label19.Visible = F; label20.Visible = F; label21.Visible = F; label22.Visible = F; label23.Visible = F; comboBox9.Visible = F; comboBox10.Visible = F; textBox8.Visible = F; textBox9.Visible = F; textBox10.Visible = F; dateTimePicker4.Visible = F; button4.Visible = F;
                    label25.Visible = T; label29.Visible = T; label30.Visible = T; label31.Visible = T; comboBox11.Visible = T; comboBox12.Visible = T; comboBox13.Visible = T; comboBox14.Visible = T; button5.Visible = T;
                }
                else
                {
                    MessageBox.Show("Введите все значения!");
                }
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (!String.IsNullOrEmpty(comboBox11.Text) && !String.IsNullOrEmpty(comboBox12.Text) && !String.IsNullOrEmpty(comboBox13.Text) && !String.IsNullOrEmpty(comboBox14.Text) && !String.IsNullOrEmpty(textBox12.Text) && !String.IsNullOrEmpty(textBox15.Text) && !String.IsNullOrEmpty(textBox16.Text) && !String.IsNullOrEmpty(textBox17.Text))//Edit
            {
                label25.Visible = F; label29.Visible = F; label30.Visible = F; label31.Visible = F; comboBox11.Visible = F; comboBox12.Visible = F; comboBox13.Visible = F; comboBox14.Visible = F; button5.Visible = F; textBox12.Visible = F; dateTimePicker3.Visible = F; textBox15.Visible = F; textBox16.Visible = F; textBox17.Visible = F; label28.Visible = F; label29.Visible = F; label32.Visible = F; label33.Visible = F; label34.Visible = F; label35.Visible = F;
                label26.Visible = T; label27.Visible = T; label36.Visible = T; label37.Visible = T; comboBox15.Visible = T; textBox18.Visible = T; textBox19.Visible = T; textBox20.Visible = T; button6.Visible = T;
                button6.Text = "Далее";
                if (comboBox2.Text == "Оригинал")
                {
                    button6.Text = "Проверить";
                }
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            if (!String.IsNullOrEmpty(textBox18.Text) && !String.IsNullOrEmpty(textBox18.Text) && !String.IsNullOrEmpty(textBox19.Text) && !String.IsNullOrEmpty(textBox20.Text) && !String.IsNullOrEmpty(comboBox15.Text))
            {
                label26.Visible = F; label27.Visible = F; label36.Visible = F; label37.Visible = F; comboBox15.Visible = F; textBox18.Visible = F; textBox19.Visible = F; textBox20.Visible = F; button6.Visible = F;
                if (comboBox2.Text == "Оригинал")
                {
                    button6.Text = "Далее";
                    button8.Visible = T;
                }
                else
                {

                    label38.Visible = T; label39.Visible = T; label40.Visible = T; label41.Visible = T; label42.Visible = T; label43.Visible = T; label44.Visible = T; label45.Visible = T; textBox21.Visible = T; textBox22.Visible = T; textBox23.Visible = T; textBox24.Visible = T; dateTimePicker2.Visible = T; textBox26.Visible = T; textBox27.Visible = T; textBox28.Visible = T; button7.Visible = T;
                }
            }
            else
            {
                MessageBox.Show("Введите все значения!");
            }
        }


        private void button7_Click(object sender, EventArgs e)
        {
            if (!String.IsNullOrEmpty(textBox28.Text) && !String.IsNullOrEmpty(textBox27.Text) && !String.IsNullOrEmpty(textBox26.Text) && !String.IsNullOrEmpty(textBox24.Text) && !String.IsNullOrEmpty(textBox23.Text) && !String.IsNullOrEmpty(textBox22.Text) && !String.IsNullOrEmpty(textBox21.Text))
            {
                label38.Visible = F; label39.Visible = F; label40.Visible = F; label41.Visible = F; label42.Visible = F; label43.Visible = F; label44.Visible = F; label45.Visible = F; textBox21.Visible = F; textBox22.Visible = F; textBox23.Visible = F; textBox24.Visible = F; dateTimePicker2.Visible = F; textBox26.Visible = F; textBox27.Visible = F; textBox28.Visible = F; button7.Visible = F;
                button8.Visible = T;

            }
            else
            {
                MessageBox.Show("Введите все значения!");
            }
        }

        private void comboBox2_TextUpdate(object sender, EventArgs e)
        {
            comboBox2.Text = "";
            MessageBox.Show("Пожалуйста выберите что-то из списка вместо того чтобы писать!");
        }

        private void comboBox9_TextChanged(object sender, EventArgs e)
        {
            if (comboBox9.Text == "Российская Федерация")
            {
                textBox11.Visible = T; label24.Visible = T;
            }
            else
            {
                textBox11.Visible = F; label24.Visible = F;
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            if (!String.IsNullOrEmpty(textBox21.Text) && !String.IsNullOrEmpty(textBox22.Text) && !String.IsNullOrEmpty(textBox23.Text) && !String.IsNullOrEmpty(textBox24.Text) && !String.IsNullOrEmpty(textBox26.Text) && !String.IsNullOrEmpty(textBox27.Text) && !String.IsNullOrEmpty(textBox28.Text))
            {
                if (comboBox2.Text == "Оригинал")
                {
                    textBox21.Text = textBox1.Text; textBox22.Text = textBox2.Text;
                    textBox23.Text = textBox3.Text; textBox24.Text = "1";
                    textBox26.Text = textBox8.Text; textBox27.Text = textBox9.Text;
                    textBox28.Text = textBox10.Text;
                }
                Data.AddData(1, textBox1.Text);
                Data.AddData(2, comboBox1.Text);
                Data.AddData(3, comboBox2.Text);
                Data.AddData(4, comboBox3.Text);
                Data.AddData(5, comboBox4.Text);
                Data.AddData(6, comboBox5.Text);
                Data.AddData(7, comboBox6.Text);
                Data.AddData(8, textBox2.Text);
                Data.AddData(9, textBox3.Text);
                Data.AddData(10, dateTimePicker1.Text);
                Data.AddData(11, "");
                Data.AddData(12, comboBox7.Text);
                Data.AddData(13, comboBox8.Text);
                Data.AddData(14, textBox4.Text);
                Data.AddData(15, textBox5.Text);
                Data.AddData(16, textBox6.Text);
                Data.AddData(17, textBox7.Text);
                Data.AddData(18, "");
                Data.AddData(19, textBox8.Text);
                Data.AddData(20, textBox9.Text);
                Data.AddData(21, textBox10.Text);
                Data.AddData(22, dateTimePicker4.Text);
                Data.AddData(23, comboBox10.Text);
                Data.AddData(24, textBox11.Text);
                Data.AddData(25, comboBox9.Text);
                Data.AddData(26, comboBox11.Text);
                Data.AddData(27, comboBox12.Text);
                Data.AddData(28, comboBox13.Text);
                Data.AddData(29, comboBox14.Text);
                Data.AddData(30, textBox12.Text);
                Data.AddData(31, dateTimePicker3.Text);
                Data.AddData(32, textBox15.Text);
                Data.AddData(33, textBox17.Text);
                Data.AddData(34, textBox16.Text);
                Data.AddData(35, textBox18.Text);
                Data.AddData(36, textBox19.Text);
                Data.AddData(37, textBox20.Text);
                Data.AddData(38, comboBox15.Text);
                Data.AddData(39, textBox21.Text);
                Data.AddData(40, textBox22.Text);
                Data.AddData(41, textBox23.Text);
                Data.AddData(42, "");
                Data.AddData(43, dateTimePicker2.Text);
                Data.AddData(44, textBox26.Text);
                Data.AddData(45, textBox27.Text);
                Data.AddData(46, textBox28.Text);
                if (MessageBox.Show("Вы уверены?", "Предупреждение", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1) == DialogResult.Yes)
                {
                    try
                    {
                        using (ExcelHelper helper = new ExcelHelper())
                        {
                            helper.Open(filePath: @"C:\Users\FIOL\Desktop\Coding\otchet_fill\pattern.xlsx");
                            int i = helper.FindLastRow();
                            for (int j = 1; j < 47; j++)
                            {
                                if (j != 18)
                                {
                                    helper.Set(row: i, col: j, data: Data.GetData(j));
                                }
                                if (j == 11 || j == 42)
                                {
                                    helper.Set(row: i, col: j, data: i);
                                }
                            }
                            helper.Save();
                        }
                    }
                    catch (Exception ex) { MessageBox.Show(ex.Message); }

                    if (MessageBox.Show("Продолжить работу?", "", MessageBoxButtons.YesNo, MessageBoxIcon.None, MessageBoxDefaultButton.Button1) == DialogResult.Yes)
                    {
                        label1.Visible = T; label2.Visible = T; label3.Visible = T; label7.Visible = T; label8.Visible = T; label9.Visible = T; label10.Visible = T; dateTimePicker1.Visible = T; comboBox1.Visible = T; comboBox2.Visible = T; comboBox6.Visible = T; textBox1.Visible = T; textBox2.Visible = T; textBox3.Visible = T; button2.Visible = T;
                        button8.Visible = F;
                        Data.ClearData();
                    }
                    else
                    {
                        this.Close();
                    }
                }

            }
            else
            {
                MessageBox.Show("Введите все значения!");
            }
        }

        private void comboBox7_TextUpdate(object sender, EventArgs e)
        {

        }

        private void button9_Click(object sender, EventArgs e)
        {

        }

        private void textBox28_TextChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {

        }

        private void textBox11_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
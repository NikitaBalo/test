using otchet_fill;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Text;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace table_add
{
    public partial class _1page : Form
    {
        public _1page()
        {
            InitializeComponent();
        }
        private List<string> list = new List<string>() {
            "Наименование документа",
            "Вид документа",
            "Статус документа",
            "Подтверждение утраты",
            "Подтверждение обмена",
            "Подтверждение  уничтожения",
            "Уровень образования",
            "Серия документа",
            "Номер документа",
            "Дата выдачи",
            "Регистрационный номер",
            "Код профессии, специальности",
            "Наименование профессии, специальности",
            "Наименование квалификации",
            "Наименование образовательной  программы",
            "Год поступления",
            "Год окончания",
            "Срок обучения, лет",
            "Фамилия получателя",
            "Имя получателя",
            "Отчество получателя",
            "Дата рождения получателя",
            "Пол получателя",
            "СНИЛС",
            "Гражданство получателя (код страны по ОКСМ)",
            "Форма обучения",
            "Форма получения образования на момент прекращения образовательных отношений",
            "Источник финансирования обучения",
            "Наличие договора о целевом обучении",
            "Номер  договора о целевом обучении",
            "Дата заключения договора о целевом обучении",
            "Наименование организации с которой заключён договор о целевом обучении",
            "ОГРН организации с которой заключён договор о целевом обучении",
            "КПП организации с которой заключён договор о целевом обучении",
            "Наименование организации работодателя",
            "ОГРН организации работодателя",
            "КПП организации работодателя",
            "Субъект федерации в котором расположена организация работодатель",
            "Наименование документа об образовании (оригинала)",
            "Серия (оригинала)",
            "Номер (оригинала)",
            "Регистрационный N (оригинала)",
            "Дата выдачи (оригинала)",
            "Фамилия получателя (оригинала)",
            "Имя получателя (оригинала)",
            "Отчество получателя (оригинала)"
        };

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog ofd = new OpenFileDialog();
                ofd.InitialDirectory = $"{Environment.GetFolderPath(Environment.SpecialFolder.Desktop)}";
                ofd.DefaultExt = ".xlsx";
                ofd.Filter = "Файл Excel. Файл формата: .xlsx | *.xlsx";
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    Data.Add1(ofd.FileName);
                }
                else
                {
                    return;
                }
                Form1 frm = new Form1();
                frm.Show();
                Hide();
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                
                button1.Enabled = false;
                FolderBrowserDialog ofd = new FolderBrowserDialog();
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    string _filePath = Path.Combine(ofd.SelectedPath, "Table.xlsx");
                    Data.Add1(_filePath);
                }
                else
                {
                    return;
                }
                using (ExcelHelper helper = new ExcelHelper())
                {
                    helper.Add(filePath: Data.Get1(0).ToString());
                    for (int j = 0; j < 46; j++)
                    helper.Set(1, j+1, list[j]);
                    helper.AutoFit();
                    helper.Save();
                }
                Form1 frm = new Form1();
                frm.Show();
                Hide();
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
        }
    }
}

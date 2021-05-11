using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ResultSheets
{
    public partial class SheetColorSchemeEditor : Form
    {
        private List<SheetColorScheme> schemes;
        private int selectedScheme = 0;
        public int SelectedScheme
        {
            get { return selectedScheme; }
        }

        // Конструктор с параметрами!!!
        public SheetColorSchemeEditor(List<SheetColorScheme> sc, int selIndex)
        {
            InitializeComponent();

            schemes = sc;

            UpdateList();
            listBox1.SelectedIndex = selIndex;
        }

        // Обновить список
        private void UpdateList()
        {
            listBox1.Items.Clear();
            foreach (SheetColorScheme cs in schemes)
            {
                listBox1.Items.Add(cs.Name);
            }

            listBox1.SelectedIndex = selectedScheme;
        }

        // Выбран новый пункт списка
        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (listBox1.SelectedIndex < 0) listBox1.SelectedIndex = 0;
            if (listBox1.SelectedIndex == 0) tb_Name.Enabled = false; else tb_Name.Enabled = true;
            selectedScheme = listBox1.SelectedIndex;
            LoadColorScheme(schemes[listBox1.SelectedIndex]);
        }

        // Новая схема
        private void bt_New_Click(object sender, EventArgs e)
        {
            SheetColorScheme cs = new SheetColorScheme();
            cs.Name = "Noname";
            schemes.Add(GetSchemeFromEditor());
            UpdateList();
            listBox1.SelectedIndex = listBox1.Items.Count - 1;
        }

        // Перезаписать в старую
        private void bt_Save_Click(object sender, EventArgs e)
        {
            schemes[listBox1.SelectedIndex] = GetSchemeFromEditor();
            UpdateList();
        }

        // Записать как новую
        private void bt_Save_New_Click(object sender, EventArgs e)
        {
            schemes.Add(GetSchemeFromEditor());
            UpdateList();
        }

        // Удалить схему
        private void bt_Delete_Click(object sender, EventArgs e)
        {
            int ind = listBox1.SelectedIndex;
            if (listBox1.SelectedIndex > 0)
            {
                schemes.RemoveAt(ind);
                if (selectedScheme >= schemes.Count) selectedScheme = schemes.Count - 1;
                UpdateList();
            }
        }

        // Загрузить схему в редактор
        private void LoadColorScheme(SheetColorScheme scs)
        {
            tb_Name.Text = scs.Name;
            bt_HDBack.BackColor = scs.HeaderBackColor;
            bt_VALBack.BackColor = scs.ValueBackColor;
            bt_TXT.BackColor = scs.TextColor;
            bt_SELBack.BackColor = scs.SelectedBackColor;
            bt_LTBack.BackColor = scs.LightingBackColor;
            bt_LTSELBack.BackColor = scs.LightingSelectedBackColor;
        }

        // Получить схему из редактора
        private SheetColorScheme GetSchemeFromEditor()
        {
            return new SheetColorScheme(
                tb_Name.Text, 
                bt_HDBack.BackColor, 
                bt_VALBack.BackColor, 
                bt_TXT.BackColor,
                bt_SELBack.BackColor, 
                bt_LTBack.BackColor, 
                bt_LTSELBack.BackColor
                );
        }

        //*****************************  Редактор цвета  ********************************
        private Color GetColor(Color initColor)
        {
            Color result = initColor;
            ColorDialog cd = new ColorDialog();
            cd.FullOpen = true;
            cd.CustomColors = new int[] { 
                ColorTranslator.ToOle(bt_HDBack.BackColor), 
                ColorTranslator.ToOle(bt_VALBack.BackColor), 
                ColorTranslator.ToOle(bt_TXT.BackColor),
                ColorTranslator.ToOle(bt_SELBack.BackColor), 
                ColorTranslator.ToOle(bt_LTBack.BackColor), 
                ColorTranslator.ToOle(bt_LTSELBack.BackColor)
            };
            cd.Color = initColor;
            if (cd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                result = cd.Color;
            }

            return result;
        }

        private void bt_HDBack_Click(object sender, EventArgs e)
        {
            bt_HDBack.BackColor = GetColor(bt_HDBack.BackColor);
        }

        private void bt_VALBack_Click(object sender, EventArgs e)
        {
            bt_VALBack.BackColor = GetColor(bt_VALBack.BackColor);
        }

        private void bt_TXT_Click(object sender, EventArgs e)
        {
            bt_TXT.BackColor = GetColor(bt_TXT.BackColor);
        }

        private void bt_SELBack_Click(object sender, EventArgs e)
        {
            bt_SELBack.BackColor = GetColor(bt_SELBack.BackColor);
        }

        private void bt_LTBack_Click(object sender, EventArgs e)
        {
            bt_LTBack.BackColor = GetColor(bt_LTBack.BackColor);
        }

        private void bt_LTSELBack_Click(object sender, EventArgs e)
        {
            bt_LTSELBack.BackColor = GetColor(bt_LTSELBack.BackColor);
        }
        //*******************************************************************************
    }
}

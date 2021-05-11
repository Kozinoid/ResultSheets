using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ProjectFileStreamer;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;

namespace ResultSheets
{
    public partial class ResultViewSheet : UserControl
    {
        private float columnNamePercent = 40.0f;
        private bool resizing = false;
        private string[] ColNames = { "Команда", "Раунд", "Всего" };
        private DataGridViewCell templateHeaderCell;
        private DataGridViewCell templateCell;
        private Font mainFont = new Font("Haettenschweiler", 24f);
        private SheetColorScheme currentCS = new SheetColorScheme();
        private bool rowSelection = true;
        private int currentRound = -1;

        // Количество раундов
        public int RoundCount
        {
            get { return dataGridView1.Columns.Count - 2; }
        }

        // Количество команд
        public int TeamCount
        {
            get { return dataGridView1.Rows.Count; }
        }

        // Режим выделения
        public bool RowSelectioin
        {
            get { return rowSelection; }
            set
            {
                rowSelection = value;
                if (rowSelection) SetRowSelectionMode();
                else SetCellSelectionMode();
            }
        }

        // Текщий раунд
        public int CurrentRound
        {
            get { return currentRound; }
        }

        // Constructor
        public ResultViewSheet()
        {
            InitializeComponent();

            InitializeTabl();
        }

        // ________________________________________  Test function  __________________________________________________
        public void TestFunction()
        {
            
        }

        //  Нулевая таблица
        public void InitializeTabl()
        {
            templateHeaderCell = CreateHeaderCellTemplate(currentCS.HeaderBackColor, currentCS.SelectedBackColor, DataGridViewContentAlignment.MiddleLeft);
            templateCell = CreateHeaderCellTemplate(currentCS.ValueBackColor, currentCS.SelectedBackColor, DataGridViewContentAlignment.MiddleCenter);

            DataGridViewColumn col0 = CreateNewColumn("cl_Team", ColNames[0], templateHeaderCell, DataGridViewTriState.True, DataGridViewContentAlignment.MiddleLeft);
            dataGridView1.Columns.Add(col0);

            DataGridViewColumn col3 = CreateNewColumn("cl_Result", ColNames[2], templateCell, DataGridViewTriState.False, DataGridViewContentAlignment.MiddleCenter);
            dataGridView1.Columns.Add(col3);

            PercentageUpdateLayouts();
        }

        //  Создание шаблона столбца  
        private DataGridViewCell CreateHeaderCellTemplate(Color backColor, Color selColor, DataGridViewContentAlignment al)
        {
            DataGridViewCell tmpCell = new DataGridViewTextBoxCell();

            tmpCell.Style.BackColor = backColor;
            tmpCell.Style.SelectionBackColor = selColor;
            tmpCell.Style.Font = mainFont;
            tmpCell.Style.Alignment = al;

            return tmpCell;
        }

        //  Создание новой колонки
        private DataGridViewColumn CreateNewColumn(string nm, string txt, DataGridViewCell tempCell, DataGridViewTriState resizable, DataGridViewContentAlignment al)
        {
            DataGridViewColumn col = new DataGridViewColumn();
            col.HeaderCell.Style.Font = mainFont;
            col.Name = nm;
            col.HeaderText = txt;
            col.CellTemplate = tempCell;
            col.Resizable = resizable;
            col.HeaderCell.Style.Alignment = al;

            return col;
        }

        // Добавление нового раунда (столбца)
        public void AddNewRound(string roundName)
        {
            string varName = string.Format("cl_Round{0}", dataGridView1.Columns.Count - 1);
            int ind = dataGridView1.Columns.Count - 1;
            DataGridViewColumn col1 = CreateNewColumn(varName, roundName, templateCell, DataGridViewTriState.False, DataGridViewContentAlignment.MiddleCenter);
            dataGridView1.Columns.Insert(dataGridView1.Columns.Count - 1, col1);

            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                dataGridView1.Rows[i].Cells[ind].Value = 0;
            }

            PercentageUpdateLayouts();
        }

        // Удаление последнего раунда (столбца)
        public void DeleteLastRound()
        {
            if (dataGridView1.Columns.Count > 2)
            {
                int ind = dataGridView1.Columns.Count - 2;
                dataGridView1.Columns.RemoveAt(ind);
                if (currentRound == dataGridView1.Columns.Count - 1) currentRound = -1;
                PercentageUpdateLayouts();
            }
        }

        // Добавление новой команды (строки)
        public void AddNewTeam(string teamName)
        {
            int newRow = dataGridView1.Rows.Add(teamName);

            for (int i = 1; i < dataGridView1.Columns.Count; i++)
            {
                dataGridView1.Rows[newRow].Cells[i].Value = 0;
            }

            UpdateColors();
        }

        // Удаление выделенной курсором команды (строки)
        public void DeleteSelectedTeam()
        {
            if (rowSelection)
            {
                DeleteSelectedRows();
            }
            else
            {
                DeleteSelectedTeamNameRows();
            }
        }

        // Удаление команд в режиме выделения по строкам
        private void DeleteSelectedRows()
        {
            if (dataGridView1.SelectedRows.Count > 0)
            {
                foreach(DataGridViewRow row in dataGridView1.SelectedRows)
                {
                    dataGridView1.Rows.Remove(row);
                }
            }
        }

        // Удаление команд в режиме выделения по ячейкам. Удаляем только строки, в которых выделено название команды
        private void DeleteSelectedTeamNameRows()
        {
            if (dataGridView1.SelectedCells.Count > 0)
            {
                foreach(DataGridViewCell cell in dataGridView1.SelectedCells)
                {
                    if (cell.ColumnIndex == 0)
                    {
                        dataGridView1.Rows.Remove(cell.OwningRow);
                    }
                }
            }
        }

        //  Расчет ширины колонок на основании процентного соотношения ширины первой колонки ко всей ширине таблицы
        private void PercentageUpdateLayouts()
        {
            if (dataGridView1.Columns.Count > 1)
            {
                int colNameWidth = (int)((float)dataGridView1.Width * columnNamePercent / 100.0f);
                int colResultWidth = (dataGridView1.Width - colNameWidth) / (dataGridView1.Columns.Count - 1);

                ResizeColumns(colNameWidth, colResultWidth);
            }
        }

        //  Расчет ширины колонок на основании ширины первой колонки
        private void WidthUpdateLayouts()
        {
            if (dataGridView1.Columns.Count > 1)
            {
                int colNameWidth = dataGridView1.Columns[0].Width;
                columnNamePercent = ((float)colNameWidth / (float)dataGridView1.Width) * 100.0f;
                int colResultWidth = (dataGridView1.Width - colNameWidth) / (dataGridView1.Columns.Count - 1);

                ResizeColumns(colNameWidth, colResultWidth);
            }
        }

        // Перерисовка таблицы
        private void ResizeColumns(int nameWidth, int otherWidth)
        {
            resizing = true;
            foreach (DataGridViewColumn cl in dataGridView1.Columns)
            {
                if (cl.Index > 0)
                {
                    cl.Width = otherWidth;
                }
                else
                {
                    cl.Width = nameWidth;
                }
            }
            resizing = false;
        }

        // Изменение размера таблицы
        private void dataGridView1_Resize(object sender, EventArgs e)
        {
            if (resizing) return;

            PercentageUpdateLayouts();
        }

        // Изменение размера колонки названия команды
        private void dataGridView1_ColumnWidthChanged(object sender, DataGridViewColumnEventArgs e)
        {
            if (resizing) return;
            if (e.Column.Index == 0)
            {
                WidthUpdateLayouts();
            }
        }

        // Режим выбора всей строки
        private void SetRowSelectionMode()
        {
            dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        }

        // Режим выбора по ячейкам
        private void SetCellSelectionMode()
        {
            dataGridView1.SelectionMode = DataGridViewSelectionMode.CellSelect;
        }

        // Выбор цветовой схемы
        public void SelectColorScheme(SheetColorScheme cs)
        {
            currentCS = cs;
            UpdateColors();
        }

        // Обновление цветов таблицы
        private void UpdateColors()
        {
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                for (int j = 0; j < dataGridView1.Columns.Count; j++)
                {
                    dataGridView1.Rows[i].Cells[j].Style.BackColor = (j == 0) ? currentCS.HeaderBackColor : 
                        (j == currentRound) ? currentCS.LightingBackColor : currentCS.ValueBackColor;
                    dataGridView1.Rows[i].Cells[j].Style.SelectionBackColor = (j == currentRound) ? currentCS.LightingSelectedBackColor : currentCS.SelectedBackColor;
                    dataGridView1.Rows[i].Cells[j].Style.ForeColor = currentCS.TextColor;
                }
            }
        }

        // Выбор ткущего раунда
        private void dataGridView1_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.ColumnIndex > 0 && e.ColumnIndex < dataGridView1.Columns.Count - 1)
            {
                currentRound = e.ColumnIndex;
                UpdateColors();
            }
        }

        // Имя раунда
        private void dataGridView1_ColumnHeaderMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.ColumnIndex > 0)
            {
                EnterNameForm enf = new EnterNameForm();
                enf.NameText = dataGridView1.Columns[e.ColumnIndex].HeaderText;
                if (enf.ShowDialog() == DialogResult.OK)
                {
                    dataGridView1.Columns[e.ColumnIndex].HeaderText = enf.NameText;
                }
            }
        }

        // Нажата клавиша
        private void dataGridView1_KeyDown(object sender, KeyEventArgs e)
        {
            switch (e.KeyCode)
            {
                case Keys.Add:
                    AddSubCellValue(true, e.Shift);
                    break;

                case Keys.Oemplus:
                    AddSubCellValue(true, e.Shift);
                    break;

                case Keys.Subtract:
                    AddSubCellValue(false, e.Shift);
                    break;

                case Keys.OemMinus:
                    AddSubCellValue(false, e.Shift);
                    break;

                case Keys.Right:
                    NextRound();
                    break;

                case Keys.Left:
                    PrevRound();
                    break;
            }
        }

        // NextRound
        private void NextRound()
        {
            if (currentRound < 1) currentRound = 1;
            else if (currentRound < dataGridView1.Columns.Count - 2) currentRound++;
            UpdateColors();
        } 

        // PrevRound
        private void PrevRound()
        {
            if (currentRound < 1) currentRound = dataGridView1.Columns.Count - 2;
            else if (currentRound > 1) currentRound--;
            UpdateColors();
        }

        // Изменение результатов выделенных ячеек
        private void AddSubCellValue(bool add, bool ctrl)
        {
            List<DataGridViewCell> selList = new List<DataGridViewCell>();
            selList.Clear();
            int min = int.MaxValue;
            if (rowSelection)
            {
                if (ctrl)
                {
                    foreach (DataGridViewRow row in dataGridView1.Rows)
                    {
                        if (!row.Selected)
                        {
                            selList.Add(row.Cells[currentRound]);
                            if ((int)row.Cells[currentRound].Value < min) min = (int)row.Cells[currentRound].Value;
                        }
                    }
                }
                else
                {
                    if (currentRound >= 0 && dataGridView1.SelectedRows.Count > 0)
                    {
                        foreach (DataGridViewRow row in dataGridView1.SelectedRows)
                        {
                            selList.Add(row.Cells[currentRound]);
                            if ((int)row.Cells[currentRound].Value < min) min = (int)row.Cells[currentRound].Value;
                        }
                    }
                }
            }
            else
            {
                foreach (DataGridViewCell cell in dataGridView1.SelectedCells)
                {
                    int ind = cell.ColumnIndex;
                    if (ind > 0 && ind < dataGridView1.Columns.Count - 1)
                    {
                        selList.Add(cell);
                        if ((int)cell.Value < min) min = (int)cell.Value;
                    }
                }
            }

            if (selList.Count > 0)
            {
                if (add)
                {
                    foreach (DataGridViewCell cell in selList)
                    {
                        int intval = (int)cell.Value;
                        intval++;
                        cell.Value = intval;
                    }
                }
                else
                {
                    if (min > 0)
                    {
                        foreach (DataGridViewCell cell in selList)
                        {
                            int intval = (int)cell.Value;
                            intval--;
                            cell.Value = intval;
                        }
                    }
                }
            }

            UpdateResults();
        }

        //  Подсчет результатов
        private void UpdateResults()
        {
            for (int j = 0; j < dataGridView1.Rows.Count; j++)
            {
                int sum = 0;
                for (int i = 1; i < dataGridView1.Columns.Count - 1; i++)
                {
                    sum += (int)dataGridView1.Rows[j].Cells[i].Value;
                }
                dataGridView1.Rows[j].Cells[dataGridView1.Columns.Count - 1].Value = sum;
            }

            //extScreen.CalculateLayouts();
        }

        // Редактор названия команд
        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 0)
            {
                dataGridView1.BeginEdit(true);
            }
        }

        // Закончить редактирование названия команд
        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            dataGridView1.EndEdit();
        }

        // Сохранение результатов
        public void SaveSheet(ProjFileStream pfs, bool templ)
        {
            int rc = dataGridView1.Rows.Count;
            int cc = dataGridView1.Columns.Count;
            pfs.WriteInt(rc);
            pfs.WriteInt(cc);

            for (int i = 1; i < cc - 1; i++)
            {
                pfs.WriteString(dataGridView1.Columns[i].HeaderText);
            }
            for (int i = 0; i < rc; i++)
            {
                pfs.WriteString((string)dataGridView1.Rows[i].Cells[0].Value);
                for (int j = 1; j < cc - 1; j++)
                {
                    if (templ)
                    {
                        pfs.WriteInt((int)0);
                    }
                    else
                    {
                        pfs.WriteInt((int)dataGridView1.Rows[i].Cells[j].Value);
                    }
                }
            }
        }

        // Загрузка результатов
        public void LoadSheet(ProjFileStream pfs)
        {
            int rc = pfs.ReadInt();
            int cc = pfs.ReadInt();
            Clear();

            for (int j = 1; j < cc - 1; j++)
            {
                AddNewRound(pfs.ReadString());
            }

            for (int i = 0; i < rc; i++)
            {
                string name = pfs.ReadString();
                AddNewTeam(name);

                for (int j = 1; j < cc - 1; j++)
                {
                    dataGridView1.Rows[i].Cells[j].Value = pfs.ReadInt();
                }
            }
            UpdateResults();
        }

        // Очистка таблицы
        public void Clear()
        {
            dataGridView1.Rows.Clear();
            dataGridView1.Columns.Clear();
            InitializeTabl();
        }

        // Export To Exel
        public void ExportToExel()
        {
            //Объявляем приложение
            Excel.Application ex = new Microsoft.Office.Interop.Excel.Application();

            //Отобразить Excel
            //ex.Visible = true;

            //Добавить рабочую книгу
            Excel.Workbook workBook = ex.Workbooks.Add(Type.Missing);

            //Отключить отображение окон с сообщениями
            ex.DisplayAlerts = false;

            //Получаем первый лист документа (счет начинается с 1)
            Excel.Worksheet sheet = (Excel.Worksheet)ex.Worksheets.get_Item(1);

            for (int i = 0; i < dataGridView1.Columns.Count; i++)
            {
                sheet.Cells[1, i + 1] = dataGridView1.Columns[i].HeaderText;
                for (int j = 0; j < dataGridView1.Rows.Count; j++)
                    sheet.Cells[j + 2, i + 1] = dataGridView1.Rows[j].Cells[i].Value;
            }
            ex.Application.ActiveWorkbook.SaveAs(Path.Combine(Application.StartupPath, "Screen.xlsx"), Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            if (ex.Visible) ex.Visible = false;
            ex.Workbooks.Close();
            ex.Quit();
            ex = null;
        }

        // Sort Results
        public void SortResults()
        {
            dataGridView1.Sort(dataGridView1.Columns[dataGridView1.Columns.Count - 1], ListSortDirection.Descending);
        }

        // Renger To JPG
        public Bitmap RengerToJPG()
        {
            return RengerToJPG(1920, 1080);
        }

        // Renger To JPG (overloaded)
        public Bitmap RengerToJPG(int width, int height)
        {
            Bitmap bmp = new Bitmap(width, height);
            Graphics bitmapGraphics = Graphics.FromImage(bmp);

            bitmapGraphics.Clear(Color.Black);
            DataGridView dataPointer = dataGridView1;
            if (dataPointer != null)
            {
                float oldFontSize = dataPointer.Columns[0].HeaderCell.Style.Font.Size;
                float bmpk = (float)dataPointer.Size.Width / (float)width;
                float bmpFontSize = oldFontSize / bmpk;
                Font bmpFnt = new System.Drawing.Font(dataPointer.Columns[0].HeaderCell.Style.Font.Name, bmpFontSize,
                    dataPointer.Columns[0].HeaderCell.Style.Font.Style, dataPointer.Columns[0].HeaderCell.Style.Font.Unit, dataPointer.Columns[0].HeaderCell.Style.Font.GdiCharSet,
                    dataPointer.Columns[0].HeaderCell.Style.Font.GdiVerticalFont);

                int bmpTeamNamesWidth = (int)((float)dataPointer.Columns[0].HeaderCell.Size.Width / bmpk);
                int bmpTeamNamesHeight = (int)((float)dataPointer.Columns[0].HeaderCell.Size.Height / bmpk);
                int bmpResValuesWidth = (int)((float)dataPointer.Columns[1].HeaderCell.Size.Width / bmpk);
                int bmpTeamValuesHeight = (int)((float)(height - bmpTeamNamesHeight) / dataPointer.Rows.Count);
                DrawTabl(bitmapGraphics, bmpTeamNamesWidth, bmpTeamNamesHeight, bmpResValuesWidth, bmpTeamValuesHeight, bmpFnt);
            }

            Bitmap scrBMP = new Bitmap(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height);
            Graphics scrGraphics = Graphics.FromImage(scrBMP);
            scrGraphics.Clear(Color.Black);

            float kx = (float)scrBMP.Width / (float)bmp.Width;
            float ky = (float)scrBMP.Height / (float)bmp.Height;
            float k = Math.Min(kx, ky);
            int newWidth = (int)((float)bmp.Width * k);
            int newHeight = (int)((float)bmp.Height * k);
            int newX = (scrBMP.Width - newWidth) / 2;
            int newY = (scrBMP.Height - newHeight) / 2;

            scrGraphics.DrawImage(bmp, newX, newY, newWidth, newHeight);

            return scrBMP;
        }

        //************************************  DRAW SHEET  *********************************************
        private void DrawTabl(Graphics gr, int tNamesWidth, int tNamesHeight, int rValWidth, int tValHeight, Font f)
        {
            DrawHeaders(gr, 0, 0, tNamesWidth, rValWidth, tNamesHeight, f);
            DrawTeams(gr, 0, tNamesHeight, tNamesWidth, rValWidth, tValHeight, f);
        }


        private void DrawHeaders(Graphics gr, int x, int y, int hdWidth, int valWidth, int height, Font fnt)
        {
            DataGridView dataPointer = dataGridView1;
            gr.FillRectangle(new SolidBrush(currentCS.HeaderBackColor), new Rectangle(x, y, hdWidth, height));
            StringFormat sf1 = new StringFormat();
            sf1.Alignment = StringAlignment.Center;
            sf1.LineAlignment = StringAlignment.Center;
            gr.DrawString(dataPointer.Columns[0].HeaderCell.Value.ToString(), fnt, new SolidBrush(currentCS.TextColor), new RectangleF(x, y, hdWidth, height), sf1);
            gr.DrawRectangle(new Pen(new SolidBrush(currentCS.TextColor)), new Rectangle(x, y, hdWidth, height));

            int curX = hdWidth;
            for (int i = 1; i < dataPointer.Columns.Count; i++)
            {
                gr.FillRectangle(new SolidBrush(currentCS.HeaderBackColor), new Rectangle(curX, y, valWidth, height));
                gr.DrawString(dataPointer.Columns[i].HeaderCell.Value.ToString(), fnt, new SolidBrush(currentCS.TextColor), new RectangleF(curX, y, valWidth, height), sf1);
                gr.DrawRectangle(new Pen(new SolidBrush(currentCS.TextColor)), new Rectangle(curX, y, valWidth, height));
                curX += valWidth;
            }
        }


        private void DrawTeams(Graphics gr, int x, int y, int hdWidth, int valWidth, int height, Font fnt)
        {
            DataGridView dataPointer = dataGridView1;
            StringFormat sf1 = new StringFormat();
            sf1.Alignment = StringAlignment.Center;
            sf1.LineAlignment = StringAlignment.Center;
            StringFormat sf2 = new StringFormat();
            sf2.Alignment = StringAlignment.Near;
            sf2.LineAlignment = StringAlignment.Center;
            sf2.FormatFlags = StringFormatFlags.NoWrap;
            int curY = y;
            for (int j = 0; j < dataPointer.Rows.Count; j++)
            {
                gr.FillRectangle(new SolidBrush(currentCS.HeaderBackColor), new Rectangle(x, curY, hdWidth, height));
                if (dataPointer.Rows[j].Cells[0].Value != null)
                {
                    gr.DrawString(dataPointer.Rows[j].Cells[0].Value.ToString(), fnt, new SolidBrush(currentCS.TextColor), new RectangleF(x, curY, hdWidth, height), sf2);
                }
                gr.DrawRectangle(new Pen(new SolidBrush(currentCS.TextColor)), new Rectangle(x, curY, hdWidth, height));

                int curX = hdWidth;
                for (int i = 1; i < dataPointer.Columns.Count; i++)
                {
                    gr.FillRectangle(new SolidBrush(currentCS.ValueBackColor), new Rectangle(curX, curY, valWidth, height));
                    gr.DrawString(dataPointer.Rows[j].Cells[i].Value.ToString(), fnt, new SolidBrush(currentCS.TextColor), new RectangleF(curX, curY, valWidth, height), sf1);
                    gr.DrawRectangle(new Pen(new SolidBrush(currentCS.TextColor)), new Rectangle(curX, curY, valWidth, height));
                    curX += valWidth;
                }
                curY += height;
            }
        }

        //***********************************************************************************************
    }

    public class SheetColorScheme
    {

        public string Name { get; set; }
        public Color HeaderBackColor { get; set; }
        public Color ValueBackColor { get; set; }
        public Color TextColor { get; set; }
        public Color SelectedBackColor { get; set; }
        public Color LightingBackColor { get; set; }
        public Color LightingSelectedBackColor { get; set; }

        public SheetColorScheme()
        {
            Name = "Default";
            HeaderBackColor = Color.Coral;
            ValueBackColor = ColorTranslator.FromWin32(0x76C4ED);//76C4ED
            TextColor = Color.Black;
            SelectedBackColor = Color.Brown;
            LightingBackColor = ColorTranslator.FromWin32(0xCBE8F8);
            LightingSelectedBackColor = ColorTranslator.FromWin32(0x6060CF);
        }

        public SheetColorScheme(string name, Color clHDBack, Color clValBack, Color clTxt, Color clSelBack, Color clLtValBack, Color clLtSelBack)
        {
            Name = name;
            HeaderBackColor = clHDBack;
            ValueBackColor = clValBack;
            TextColor = clTxt;
            SelectedBackColor = clSelBack;
            LightingBackColor = clLtValBack;
            LightingSelectedBackColor = clLtSelBack;
        }

        public static void WriteToStream(ProjFileStream writer, SheetColorScheme scheme)
        {
            writer.WriteString(scheme.Name);
            writer.WriteColor(scheme.HeaderBackColor);
            writer.WriteColor(scheme.ValueBackColor);
            writer.WriteColor(scheme.TextColor);
            writer.WriteColor(scheme.SelectedBackColor);
            writer.WriteColor(scheme.LightingBackColor);
            writer.WriteColor(scheme.LightingSelectedBackColor);
        }

        public static SheetColorScheme ReadFromStream(ProjFileStream reader)
        {
            SheetColorScheme resultScheme = new SheetColorScheme();
            resultScheme.Name = reader.ReadString();
            resultScheme.HeaderBackColor = reader.ReadColor();
            resultScheme.ValueBackColor = reader.ReadColor();
            resultScheme.TextColor = reader.ReadColor();
            resultScheme.SelectedBackColor = reader.ReadColor();
            resultScheme.LightingBackColor = reader.ReadColor();
            resultScheme.LightingSelectedBackColor = reader.ReadColor();
            return resultScheme;
        }
    }
}

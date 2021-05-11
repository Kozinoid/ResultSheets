using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using ProjectFileStreamer;


namespace ResultSheets
{
    public partial class Form1 : Form
    {
        private const int port = 8060;
        private List<SheetColorScheme> schemes = new List<SheetColorScheme>();
        private bool saved = true;
        private string pathToFile = "";
        private string tempPath;

        private LocalNetServer localServer;

        // Constructor
        public Form1()
        {
            InitializeComponent();
            localServer = new LocalNetServer(port);
            MakeTemplate();
            Text += " (" + localServer.GetConnectionIP() + ")";
        }

        // Загрузка исходных настроек
        private void Form1_Load(object sender, EventArgs e)
        {
            ColorSchemesRead();
            UpdateColorSchemesComboBox();
            toolStripComboBox1.SelectedIndex = 0;
        }

        // Сохранение настроек
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            localServer.Disconnect();

            ColorSchemesWrite();

             if (!saved)
             {
                 NeedToSaveMenu();
             }
        }

        // Загрузка цветовых схем
        private void ColorSchemesRead()
        {
            schemes.Clear();
            string colorFileName = Path.Combine(Application.StartupPath, "Colors");
            if (File.Exists(colorFileName))
            {
                ProjFileStream reader = new ProjFileStream(colorFileName, FileMode.Open, FileAccess.Read);
                
                int cnt = reader.ReadInt();
                for (int i = 0; i < cnt; i++)
                {
                    schemes.Add(SheetColorScheme.ReadFromStream(reader));
                }

                reader.Close();
            }
            else
            {
                schemes.Add(new SheetColorScheme());
            }
        }

        // Обновить список цветовых схем
        private void UpdateColorSchemesComboBox()
        {
            toolStripComboBox1.Items.Clear();
            foreach (SheetColorScheme cs in schemes)
            {
                toolStripComboBox1.Items.Add(cs.Name);
            }
        }

        // Сохранение цветовых схем
        private void ColorSchemesWrite()
        {
            string colorFileName = Path.Combine(Application.StartupPath, "Colors");

            ProjFileStream writer = new ProjFileStream(colorFileName, FileMode.Create, FileAccess.Write);

            int cnt = schemes.Count;
            writer.WriteInt(cnt);
            for (int i = 0; i < cnt; i++)
            {
                SheetColorScheme.WriteToStream(writer, schemes[i]);
            }

            writer.Close();
        }

        // Template
        private void MakeTemplate()
        {
            tempPath = Path.Combine(Path.GetDirectoryName(Application.ExecutablePath), "temp");

            if (File.Exists(tempPath))
            {
                ProjFileStream reader = new ProjFileStream(tempPath, FileMode.Open, FileAccess.Read);

                resultViewSheet1.LoadSheet(reader);
            }
            else
            {
                for (int i = 0; i < 2; i++)
                {
                    resultViewSheet1.AddNewRound(("Раунд " + (resultViewSheet1.RoundCount + 1).ToString()));
                }
                for (int i = 0; i < 3; i++)
                {
                    resultViewSheet1.AddNewTeam(string.Format("Team {0}", resultViewSheet1.TeamCount.ToString()));
                }
            }
        }

        // MENU: TEST
        private void testToolStripMenuItem_Click(object sender, EventArgs e)
        {
            resultViewSheet1.TestFunction();
        }

        // МЕНЮ: Выбор цветовой схемы
        private void toolStripComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            resultViewSheet1.SelectColorScheme(schemes[toolStripComboBox1.SelectedIndex]);
            saved = false;
        }

        // МЕНЮ: Добавить раунд
        private void addColumnToolStripMenuItem_Click(object sender, EventArgs e)
        {
            resultViewSheet1.AddNewRound(("Раунд " + (resultViewSheet1.RoundCount + 1).ToString()));
            saved = false;
        }

        // МЕНЮ: Удалить раунд
        private void deleteColumnToolStripMenuItem_Click(object sender, EventArgs e)
        {
            resultViewSheet1.DeleteLastRound();
            saved = false;
        }

        // МЕНЮ: Добавить команду
        private void addRowToolStripMenuItem_Click(object sender, EventArgs e)
        {
            resultViewSheet1.AddNewTeam(string.Format("Team {0}", resultViewSheet1.TeamCount.ToString()));
            saved = false;
        }

        // МЕНЮ: Удалить команду
        private void deleteRowToolStripMenuItem_Click(object sender, EventArgs e)
        {
            resultViewSheet1.DeleteSelectedTeam();
            saved = false;
        }

        // МЕНЮ: Режимы выделения
        private void rowSelectionToolStripMenuItem_Click(object sender, EventArgs e)
        {
            resultViewSheet1.RowSelectioin = rowSelectionToolStripMenuItem.Checked;
        }

        // МЕНЮ: Редактор схем
        private void colorSchemeListEditorToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SheetColorSchemeEditor scse = new SheetColorSchemeEditor(schemes, toolStripComboBox1.SelectedIndex);
            scse.ShowDialog();
            UpdateColorSchemesComboBox();

            toolStripComboBox1.SelectedIndex = (scse.SelectedScheme < 0)?0:scse.SelectedScheme;

            saved = false;
        }

        // МЕНЮ: SORT
        private void sortToolStripMenuItem_Click(object sender, EventArgs e)
        {
            resultViewSheet1.SortResults();
        }

        //***********************************************  FILE  **************************************************
        private void newProjectToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (!saved)
            {
                NeedToSaveMenu();
            }
            resultViewSheet1.Clear();
        }


        private void openProjectToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (!saved)
            {
                NeedToSaveMenu();
            }

            GetNameAndLoad();
        }


        private void GetNameAndLoad()
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Таблицы|*.shf";
            if (ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                pathToFile = ofd.FileName;

                LoadingProcess();
            }
        }


        private void LoadingProcess()
        {
            ProjFileStream reader = new ProjFileStream(pathToFile, FileMode.Open, FileAccess.Read);

            resultViewSheet1.LoadSheet(reader);

            reader.Close();
        }


        private void saveProjectToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (!File.Exists(pathToFile))
            {
                GetNameAndSave();
            }
            else
            {
                SavingProcess();
            }
        }


        private void saveProjectAsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            GetNameAndSave();
        }


        private void GetNameAndSave()
        {
            string newname = GetNewFileName();
            if (newname != "")
            {
                pathToFile = newname;
                SavingProcess();
            }
        }


        private void SavingProcess()
        {
            ProjFileStream writer = new ProjFileStream(pathToFile, FileMode.Create, FileAccess.Write);

            resultViewSheet1.SaveSheet(writer, false);

            writer.Close();
            saved = true;
        }


        private void NeedToSaveMenu()
        {
            if (MessageBox.Show("Current file is not saved! Do you want to save?", "File is not saved", MessageBoxButtons.YesNo)
                == System.Windows.Forms.DialogResult.Yes)
            {
                string newname = "";
                if (!File.Exists(pathToFile))
                {
                    newname = GetNewFileName();
                }

                if (newname != "")
                {
                    pathToFile = newname;
                    SavingProcess();
                }
            }
        }


        private string GetNewFileName()
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Таблицы|*.shf";
            if (sfd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                return sfd.FileName;
            }
            else
            {
                return "";
            }
        }


        private void saveAsTemplateToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ProjFileStream writer = new ProjFileStream(tempPath, FileMode.Create, FileAccess.Write);
            resultViewSheet1.SaveSheet(writer, true);
        }

        //********************************************  EXPORT  ************************************************
        private void renderAsJPGToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Bitmap bmp = resultViewSheet1.RengerToJPG();
            bmp.Save(Path.Combine(Application.StartupPath, "Screen.jpg"));
        }


        private void exportToXLSToolStripMenuItem_Click(object sender, EventArgs e)
        {
            resultViewSheet1.ExportToExel();
        }

    }
}

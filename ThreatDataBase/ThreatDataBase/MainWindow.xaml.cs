using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Net;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace ThreatDataBase
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        WebClient wc = new WebClient();
        Excel.Application ObjWorkExcel = new Excel.Application();
        static Dictionary<int,ThreatInfo> list = new Dictionary<int,ThreatInfo>();
        static Dictionary<int, ThreatInfo> updatedlist = new Dictionary<int, ThreatInfo>();
        
        static List<ThreatShortInfo> shortlist = new List<ThreatShortInfo>();
        static List<ThreatShortInfo> bufferlist = new List<ThreatShortInfo>();
        static List<int> updatedIds = new List<int>();
        static bool isViewAllopen = false;
        static bool isViewOneopen = false;
        static bool isUpdateopen = false;
        static int page = 0;

        public MainWindow()
        {
            InitializeComponent();
            HideElements();
            StartMessage.Visibility = Visibility.Visible;
            if (File.Exists("C:\\ThreatBase.xlsx"))
            {
               
                StartMessage.Text = "На вашем компьютере существует загруженная база угроз ИБ";
               
                FromExcelToList("C:\\ThreatBase.xlsx",list);


            }
            else
            {
                StartMessage.Text = "На Вашем компьютере отсутствует база данных по УБИ";
                FirstUpdate.Visibility = Visibility.Visible;
            }
        }

        private void FirstUpdate_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                wc.DownloadFile("https://bdu.fstec.ru/files/documents/thrlist.xlsx", "C:\\ThreatBase.xlsx");
               MessageBox.Show( "База данных загружена!");
                StartMessage.Text = "На вашем компьютере существует загруженная база угроз ИБ";
                FirstUpdate.Visibility = Visibility.Collapsed;
                FromExcelToList("C:\\ThreatBase.xlsx",list);
                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void FromExcelToList(string path, Dictionary<int, ThreatInfo> list)
        {
            Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(path, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];
            int numberofrows = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
            for (int i = 3; i <= numberofrows; i++)
            {

                int id =Convert.ToInt32( ObjWorkSheet.Cells[i, 1].Text.ToString());
                string name = ObjWorkSheet.Cells[i, 2].Text.ToString();
                string description = ObjWorkSheet.Cells[i, 3].Text.ToString();
                string source = ObjWorkSheet.Cells[i, 4].Text.ToString();
                string target = ObjWorkSheet.Cells[i, 5].Text.ToString();
                string breachOfConfid;
                if (ObjWorkSheet.Cells[i, 6].Text.ToString() == "1")
                {
                     breachOfConfid = "да";
                }
                else
                {
                    breachOfConfid = "нет";
                }
                string integrityViolation;
                if (ObjWorkSheet.Cells[i, 7].Text.ToString() == "1")
                {
                    integrityViolation = "да";
                }
                else
                {
                    integrityViolation = "нет";
                }
                string accessibilityViolation;
                if (ObjWorkSheet.Cells[i, 8].Text.ToString() == "1")
                {
                    accessibilityViolation = "да";
                }
                else
                {
                    accessibilityViolation = "нет";
                }
                ThreatInfo ti = new ThreatInfo(id, name, description, source, target, breachOfConfid, integrityViolation, accessibilityViolation);
                list.Add(id,ti);
                ThreatShortInfo shortti = new ThreatShortInfo("УБИ." + id, name);
                shortlist.Add(shortti);
               

            }
            ObjWorkBook.Close(false, Type.Missing, Type.Missing); 
            ObjWorkExcel.Quit(); 
            GC.Collect(); 
           
        }

        private void ViewAll_Click(object sender, RoutedEventArgs e)
        {
            if (list.Count != 0)
            {
                HideElements();
               
                isViewOneopen = false;
                isUpdateopen = false;
                if (isViewAllopen == false)
                {
                    ThreatsList.Visibility = Visibility.Visible;
                    Prev.Visibility = Visibility.Visible;
                    Next.Visibility = Visibility.Visible;
                    Pages.Visibility = Visibility.Visible;
                    isViewAllopen = true;
                    bufferlist = new List<ThreatShortInfo>();
                    for (int i = 0; i < 20; i++)
                    {
                        bufferlist.Add(shortlist[i]);
                    }
                    ThreatsList.ItemsSource = bufferlist;
                    Pages.Text = "1 - 20";
                }
                else
                {
                    ThreatsList.Visibility = Visibility.Collapsed;
                    isViewAllopen = false;
                    StartMessage.Visibility = Visibility.Visible;
                }
            }
            else
            {
                MessageBox.Show("Пожалуйста, произведите первичную загрузку базы данных!");
            }
            
        }

        private void EnterIdButton_Click(object sender, RoutedEventArgs e)
        {
            
                EnterId.Visibility = Visibility.Visible;
                EnterIdButton.Visibility = Visibility.Visible;
            View.Visibility = Visibility.Collapsed;
            string s = EnterId.Text;
                int id;
                List<ThreatField> fields = new List<ThreatField>();

                if (!Int32.TryParse(s, out id))
                {
                    MessageBox.Show("Введите корректный идентификатор");
                    EnterId.Text = "";
                }
                else
                {
                    if (!list.ContainsKey(id))
                    {
                        MessageBox.Show("УБИ с данным модификатором отсутствует!");
                    }
                    else
                    {
                        fields.Add(new ThreatField() { FieldName = "Идентификатор угрозы", Field = list[id].Id.ToString() });
                        fields.Add(new ThreatField() { FieldName = "Наименование угрозы", Field = list[id].Name });
                        fields.Add(new ThreatField() { FieldName = "Описание угрозы", Field = list[id].Description });
                        fields.Add(new ThreatField() { FieldName = "Источник угрозы", Field = list[id].Source });
                        fields.Add(new ThreatField() { FieldName = "Объект воздействия угрозы", Field = list[id].Target });
                        fields.Add(new ThreatField() { FieldName = "Нарушение конфиденциальности", Field = list[id].BreachOfConfid });
                        fields.Add(new ThreatField() { FieldName = "Нарушение целостности", Field = list[id].IntegrityViolation });
                        fields.Add(new ThreatField() { FieldName = "Нарушение доступности", Field = list[id].AccessibilityViolation });
                    View.Visibility = Visibility.Visible;
                    View.ItemsSource = fields;


                    }
                }
            
        }
        private void HideElements()
        {
            FirstUpdate.Visibility = Visibility.Collapsed;
            StartMessage.Visibility = Visibility.Collapsed;
            ThreatsList.Visibility = Visibility.Collapsed;
            EnterId.Visibility = Visibility.Collapsed;
            EnterIdMessage.Visibility = Visibility.Collapsed;
            View.Visibility = Visibility.Collapsed;
            EnterIdButton.Visibility = Visibility.Collapsed;
            Prev.Visibility = Visibility.Collapsed;
            Next.Visibility = Visibility.Collapsed;
            Pages.Visibility = Visibility.Collapsed;
            UpdateButton.Visibility = Visibility.Collapsed;
            UpdatedThreat.Visibility = Visibility.Collapsed;
            UpdatedThreatId.Visibility = Visibility.Collapsed;
            UpdateMessage.Visibility = Visibility.Collapsed;
            UpdateStatus.Visibility = Visibility.Collapsed;
            
            page = 0;
        }

        private void ViewOne_Click(object sender, RoutedEventArgs e)
        {
            if (list.Count != 0)
            {
                HideElements();
                isViewAllopen = false;
                isUpdateopen = false;

                EnterId.Text = "";
                if (isViewOneopen == false)
                {
                    EnterId.Visibility = Visibility.Visible;
                    EnterIdButton.Visibility = Visibility.Visible;
                    EnterIdMessage.Visibility = Visibility.Visible;

                    isViewOneopen = true;
                }
                else
                {
                    EnterId.Visibility = Visibility.Collapsed;
                    EnterIdButton.Visibility = Visibility.Collapsed;
                    EnterIdMessage.Visibility = Visibility.Collapsed;
                    View.Visibility = Visibility.Collapsed;
                    isViewOneopen = false;
                    StartMessage.Visibility = Visibility.Visible;
                }
            }
            else
            {
                MessageBox.Show("Пожалуйста, произведите первичную загрузку базы данных!");
            }
        }

        private void Prev_Click(object sender, RoutedEventArgs e)
        {
            if (page > 0)
            {
                page--;
                bufferlist = new List<ThreatShortInfo>();
                if (page * 20 + 20 <= shortlist.Count - 1)
                {
                    for (int i = page * 20; i < page * 20 + 20; i++)
                    {
                        bufferlist.Add(shortlist[i]);
                    }
                    Pages.Text = (page * 20 + 1) + " - " + (page * 20 + 20);
                }
                else
                {
                    for (int i = page * 20; i < shortlist.Count; i++)
                    {
                        bufferlist.Add(shortlist[i]);
                    }
                    Pages.Text = (page * 20 + 1) + " - " + shortlist.Count;
                }
                ThreatsList.ItemsSource = bufferlist;
            }
        }

        private void Next_Click(object sender, RoutedEventArgs e)
        {
            if ((page+1) * 20 <= shortlist.Count - 1)
            {
                page++;
                bufferlist = new List<ThreatShortInfo>();
                if (page * 20 + 20 <= shortlist.Count - 1)
                {
                    for (int i = page * 20; i < page * 20 + 20; i++)
                    {
                        bufferlist.Add(shortlist[i]);
                    }
                    Pages.Text = (page * 20 + 1) + " - " + (page * 20 + 20);
                }
                else
                {
                    for (int i = page * 20; i < shortlist.Count; i++)
                    {
                        bufferlist.Add(shortlist[i]);
                    }
                    Pages.Text = (page * 20 + 1) + " - " + shortlist.Count;
                }
                ThreatsList.ItemsSource = bufferlist;
            }
        }

        private void Update_Click(object sender, RoutedEventArgs e)
        {
            if (list.Count != 0)
            {
                HideElements();
                isViewAllopen = false;
                isViewOneopen = false;
                
                if (isUpdateopen == false)
                {
                    UpdateButton.Visibility = Visibility.Visible;
                    isUpdateopen = true;
                }
                else
                {
                    UpdateButton.Visibility = Visibility.Collapsed;
                    isUpdateopen = false;
                    StartMessage.Visibility = Visibility.Visible;
                }
            }
            else
            {
                MessageBox.Show("Пожалуйста, произведите первичную загругку базы данных!");
            }
           
        }

        private void UpdateButton_Click(object sender, RoutedEventArgs e)
        {
            UpdateStatus.Visibility = Visibility.Collapsed;
            UpdateMessage.Visibility = Visibility.Collapsed;
            UpdatedThreat.Visibility = Visibility.Collapsed;
            UpdatedThreatId.Visibility = Visibility.Collapsed;
            if (updatedlist.Count != 0)
            {
                list = updatedlist;
                updatedlist.Clear();
            }
            if (list.Count != 0)
            {
                try
                {


                    wc.DownloadFile("https://bdu.fstec.ru/files/documents/thrlist.xlsx", "C:\\ThreatBase.xlsx");
                    updatedIds.Clear();

                    int count = 0;
                    FromExcelToList("C:\\ThreatBase.xlsx", updatedlist);
                    foreach (var item in list)
                    {
                        if (!item.Value.Equals(updatedlist[item.Key]) || !updatedlist.ContainsKey(item.Key))
                        {

                            count++;
                            updatedIds.Add(item.Key);

                        }
                    }
                    foreach (var item in updatedlist)
                    {
                        if (!list.ContainsKey(item.Key))
                        {
                            count++;
                            updatedIds.Add(item.Key);
                        }
                    }
                    UpdateStatus.Visibility = Visibility.Visible;
                    UpdateStatus.Text = "База данных загружена успешно!";
                    UpdateMessage.Visibility = Visibility.Visible;
                    UpdateMessage.Text = $"Всего обновлено записей: {count}";

                    if (updatedIds.Count != 0)
                    {
                        UpdatedThreatId.Visibility = Visibility.Visible;
                        UpdatedThreatId.Text = "УБИ." + updatedIds[0].ToString();
                       
                        if (list.ContainsKey(updatedIds[0]))
                        {

                            List<UpdatedThreatField> updatedfields = new List<UpdatedThreatField>();
                            if (list[updatedIds[0]].Name != updatedlist[updatedIds[0]].Name)
                            {
                                updatedfields.Add(new UpdatedThreatField() { FieldName = "Идентификатор угрозы", Field = list[updatedIds[0]].Name, UpdatedField = updatedlist[updatedIds[0]].Name });
                            }
                            if (list[updatedIds[0]].Description != updatedlist[updatedIds[0]].Description)
                            {
                                updatedfields.Add(new UpdatedThreatField() { FieldName = "Описание угрозы", Field = list[updatedIds[0]].Description, UpdatedField = updatedlist[updatedIds[0]].Description });
                            }
                            if (list[updatedIds[0]].Source != updatedlist[updatedIds[0]].Source)
                            {
                                updatedfields.Add(new UpdatedThreatField() { FieldName = "Источник угрозы", Field = list[updatedIds[0]].Source, UpdatedField = updatedlist[updatedIds[0]].Source });
                            }
                            if (list[updatedIds[0]].Target != updatedlist[updatedIds[0]].Target)
                            {
                                updatedfields.Add(new UpdatedThreatField() { FieldName = "Объект воздействия угрозы", Field = list[updatedIds[0]].Target, UpdatedField = updatedlist[updatedIds[0]].Target });
                            }
                            if (list[updatedIds[0]].BreachOfConfid != updatedlist[updatedIds[0]].BreachOfConfid)
                            {
                                updatedfields.Add(new UpdatedThreatField() { FieldName = "Нарушение конфиденциальности", Field = list[updatedIds[0]].BreachOfConfid, UpdatedField = updatedlist[updatedIds[0]].BreachOfConfid });
                            }
                            if (list[updatedIds[0]].IntegrityViolation != updatedlist[updatedIds[0]].IntegrityViolation)
                            {
                                updatedfields.Add(new UpdatedThreatField() { FieldName = "Нарушение целостности", Field = list[updatedIds[0]].IntegrityViolation, UpdatedField = updatedlist[updatedIds[0]].IntegrityViolation });
                            }
                            if (list[updatedIds[0]].AccessibilityViolation != updatedlist[updatedIds[0]].AccessibilityViolation)
                            {
                                updatedfields.Add(new UpdatedThreatField() { FieldName = "Нарушение доступности", Field = list[updatedIds[0]].AccessibilityViolation, UpdatedField = updatedlist[updatedIds[0]].AccessibilityViolation });
                            }
                            UpdatedThreat.Visibility = Visibility.Visible;
                            UpdatedThreat.ItemsSource = updatedfields;
                            UpdatedThreatId.Text += "   " + updatedfields.Count;

                        }
                    }

                }
                catch (Exception ex)
                {
                    UpdateStatus.Visibility = Visibility.Visible;
                    UpdateStatus.Text = "Ошибка!";
                    UpdateMessage.Visibility = Visibility.Visible;
                    UpdateMessage.Text = ex.Message;
                }
            }
            else
            {
                MessageBox.Show("Пожалуйста, произведите первичную загругку базы данных!");
            }
        }
    }

}

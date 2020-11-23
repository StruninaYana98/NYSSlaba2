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
using System.Threading;
using System.Collections.ObjectModel;

namespace ThreatDataBase
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        WebClient wc = new WebClient();
        Excel.Application ObjWorkExcel = new Excel.Application();
        static Dictionary<int, ThreatInfo> list = new Dictionary<int, ThreatInfo>();
        static Dictionary<int, ThreatInfo> updatedlist = new Dictionary<int, ThreatInfo>();
        static ObservableCollection<UpdatedThreatField> updatedfields = new ObservableCollection<UpdatedThreatField>();
        static ObservableCollection<ThreatShortInfo> shortlist = new ObservableCollection<ThreatShortInfo>();
        static List<ThreatShortInfo> bufferlist = new List<ThreatShortInfo>();
        static List<int> updatedIds = new List<int>();
        static bool isViewAllopen = false;
        static bool isViewOneopen = false;
        static bool isUpdateopen = false;
        static int page = 0;
        string path = Directory.GetCurrentDirectory() + "\\ThreatBase.xlsx";
        public MainWindow()
        {
            InitializeComponent();
            HideElements();
           
           
            StartMessage.Visibility = Visibility.Visible;
            Img.Visibility = Visibility.Visible;
            if (File.Exists(Directory.GetCurrentDirectory() + "\\SavedThreatBase.xlsx"))
            {

                StartMessage.Text = $"На вашем компьютере существует загруженная база угроз ИБ\n";

                FromExcelToList(Directory.GetCurrentDirectory() + "\\SavedThreatBase.xlsx", list);

                StartMessage.Text += $"Всего записей в базе данных: {list.Count}";
                if(File.Exists(Directory.GetCurrentDirectory() + "\\updatedate.txt"))
                {
                    UpdateDate.Text = "Последнее обновление: " + File.ReadAllText(Directory.GetCurrentDirectory() + "\\updatedate.txt");
                }

            }
            else
            {
                StartMessage.Text = $"На Вашем компьютере отсутствует база данных по УБИ\nПожалуйста, загрузите базу данных\n";
                FirstUpdate.Visibility = Visibility.Visible;
            }
        }

       
        private void FirstUpdate_Click(object sender, RoutedEventArgs e)
        {

            try
            {
                wc.DownloadFile("https://bdu.fstec.ru/files/documents/thrlist.xlsx", path);
                FromExcelToList(path, list);
                FirstUpdate.Visibility = Visibility.Collapsed;
                
                StartMessage.Text = $"На вашем компьютере существует загруженная база угроз ИБ\n";
                StartMessage.Text += $"Всего записей в базе данных: {list.Count}";
                
                File.WriteAllText(Directory.GetCurrentDirectory() + "\\updatedate.txt", DateTime.Now.ToString());
                UpdateDate.Text = "Последнее обновление: " + File.ReadAllText(Directory.GetCurrentDirectory() + "\\updatedate.txt");


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        private void FromExcelToList(string p, Dictionary<int, ThreatInfo> l)
        {


            Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(p, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];
            int numberofrows = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
            shortlist.Clear();
            for (int i = 3; i <= numberofrows; i++)
            {

                int id = Convert.ToInt32(ObjWorkSheet.Cells[i, 1].Text.ToString());
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
                l.Add(id, ti);
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
                    Rect2.Visibility = Visibility.Visible;

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
                    Pages.Text = "1 - 20 из " + shortlist.Count;
                }
                else
                {
                    ThreatsList.Visibility = Visibility.Collapsed;
                    Rect2.Visibility = Visibility.Collapsed;
                    isViewAllopen = false;
                    StartMessage.Visibility = Visibility.Visible;
                    Img.Visibility = Visibility.Visible;
                    Img.Visibility = Visibility.Visible;

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
                if (updatedlist.Count == 0 && !list.ContainsKey(id) || updatedlist.Count != 0 && !updatedlist.ContainsKey(id))
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

            UpdateMessage.Visibility = Visibility.Collapsed;
            UpdateStatus.Visibility = Visibility.Collapsed;
            Rect1.Visibility = Visibility.Collapsed;
            Rect2.Visibility = Visibility.Collapsed;
            Rect3.Visibility = Visibility.Collapsed;
            Img.Visibility = Visibility.Collapsed;
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
                    Rect3.Visibility = Visibility.Visible;
                    isViewOneopen = true;
                }
                else
                {
                    EnterId.Visibility = Visibility.Collapsed;
                    EnterIdButton.Visibility = Visibility.Collapsed;
                    EnterIdMessage.Visibility = Visibility.Collapsed;
                    Rect3.Visibility = Visibility.Collapsed;
                    View.Visibility = Visibility.Collapsed;
                    isViewOneopen = false;
                    StartMessage.Visibility = Visibility.Visible;

                    Img.Visibility = Visibility.Visible;
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
                    Pages.Text = (page * 20 + 1) + " - " + (page * 20 + 20) + " из " + shortlist.Count;
                }
                else
                {
                    for (int i = page * 20; i < shortlist.Count; i++)
                    {
                        bufferlist.Add(shortlist[i]);
                    }
                    Pages.Text = (page * 20 + 1) + " - " + shortlist.Count + " из " + shortlist.Count;
                }
                ThreatsList.ItemsSource = bufferlist;
            }
        }

        private void Next_Click(object sender, RoutedEventArgs e)
        {
            if ((page + 1) * 20 <= shortlist.Count - 1)
            {
                page++;
                bufferlist = new List<ThreatShortInfo>();
                if (page * 20 + 20 <= shortlist.Count - 1)
                {
                    for (int i = page * 20; i < page * 20 + 20; i++)
                    {
                        bufferlist.Add(shortlist[i]);
                    }
                    Pages.Text = (page * 20 + 1) + " - " + (page * 20 + 20) + " из " + shortlist.Count;
                }
                else
                {
                    for (int i = page * 20; i < shortlist.Count; i++)
                    {
                        bufferlist.Add(shortlist[i]);
                    }
                    Pages.Text = (page * 20 + 1) + " - " + shortlist.Count + " из " + shortlist.Count;
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
                    isUpdateopen = true;
                    if (updatedfields.Count != 0)
                    {
                        UpdatedThreat.Visibility = Visibility.Visible;
                    }
                    UpdateButton.Visibility = Visibility.Visible;
                    Rect1.Visibility = Visibility.Visible;
                    UpdateStatus.Visibility = Visibility.Visible;
                    UpdateMessage.Visibility = Visibility.Visible;

                }
                else
                {
                    Rect1.Visibility = Visibility.Collapsed;
                    isUpdateopen = false;
                    StartMessage.Visibility = Visibility.Visible;
                    Img.Visibility = Visibility.Visible;
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

            updatedfields.Clear();
            min = 60;


            try
            {
                 wc.DownloadFile("https://bdu.fstec.ru/files/documents/thrlist.xlsx", Directory.GetCurrentDirectory() + "\\UpdatedThreatBase.xlsx");
                updatedIds.Clear();
                int count = 0;
                FromExcelToList(Directory.GetCurrentDirectory() + "\\UpdatedThreatBase.xlsx", updatedlist);


                foreach (var item in list)
                {
                    if (!updatedlist.ContainsKey(item.Key) || !item.Value.Equals(updatedlist[item.Key]))
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



                    for (int i = 0; i < updatedIds.Count; i++)
                    {
                        UpdatedThreatField threatField = new UpdatedThreatField();
                        threatField.Id = "УБИ." + updatedIds[i].ToString();

                        threatField.FieldName.Name = "Наименование угрозы";
                        threatField.FieldName.Description = "Описание угрозы";
                        threatField.FieldName.Source = "Источник угрозы";

                        threatField.FieldName.Target = "Объект воздействия угрозы";
                        threatField.FieldName.BreachOfConfid = "Нарушение конфиденциальности";
                        threatField.FieldName.IntegrityViolation = "Нарушение целостности";
                        threatField.FieldName.AccessibilityViolation = "Нарушение доступности";

                        if (list.ContainsKey(updatedIds[i]) && updatedlist.ContainsKey(updatedIds[i]))
                        {


                            threatField.Field.Name = "нет изменений";
                            threatField.UpdatedField.Name = "нет изменений";

                            threatField.Field.Description = "нет изменений";
                            threatField.UpdatedField.Description = "нет изменений";

                            threatField.Field.Source = "нет изменений";
                            threatField.UpdatedField.Source = "нет изменений";

                            threatField.Field.Target = "нет изменений";
                            threatField.UpdatedField.Target = "нет изменений";

                            threatField.Field.BreachOfConfid = "нет изменений";
                            threatField.UpdatedField.BreachOfConfid = "нет изменений";


                            threatField.Field.IntegrityViolation = "нет изменений";
                            threatField.UpdatedField.IntegrityViolation = "нет изменений";

                            threatField.Field.AccessibilityViolation = "нет изменений";
                            threatField.UpdatedField.AccessibilityViolation = "нет изменений";

                            if (list[updatedIds[i]].Name != updatedlist[updatedIds[i]].Name)
                            {


                                threatField.Field.Name = list[updatedIds[i]].Name;
                                threatField.UpdatedField.Name = updatedlist[updatedIds[i]].Name;
                            }
                            if (list[updatedIds[i]].Description != updatedlist[updatedIds[i]].Description)
                            {

                                threatField.Field.Description = list[updatedIds[i]].Description;
                                threatField.UpdatedField.Description = updatedlist[updatedIds[i]].Description;
                            }
                            if (list[updatedIds[i]].Source != updatedlist[updatedIds[i]].Source)
                            {


                                threatField.Field.Source = list[updatedIds[i]].Source;
                                threatField.UpdatedField.Source = updatedlist[updatedIds[i]].Source;
                            }
                            if (list[updatedIds[i]].Target != updatedlist[updatedIds[i]].Target)
                            {

                                threatField.Field.Target = list[updatedIds[i]].Target;
                                threatField.UpdatedField.Target = updatedlist[updatedIds[i]].Target;
                            }
                            if (list[updatedIds[i]].BreachOfConfid != updatedlist[updatedIds[i]].BreachOfConfid)
                            {

                                threatField.Field.BreachOfConfid = list[updatedIds[i]].BreachOfConfid;
                                threatField.UpdatedField.BreachOfConfid = updatedlist[updatedIds[i]].BreachOfConfid;
                            }
                            if (list[updatedIds[i]].IntegrityViolation != updatedlist[updatedIds[i]].IntegrityViolation)
                            {


                                threatField.Field.IntegrityViolation = list[updatedIds[i]].IntegrityViolation;
                                threatField.UpdatedField.IntegrityViolation = updatedlist[updatedIds[i]].IntegrityViolation;
                            }
                            if (list[updatedIds[i]].AccessibilityViolation != updatedlist[updatedIds[i]].AccessibilityViolation)
                            {

                                threatField.Field.AccessibilityViolation = list[updatedIds[i]].AccessibilityViolation;
                                threatField.UpdatedField.AccessibilityViolation = updatedlist[updatedIds[i]].AccessibilityViolation;
                            }


                        }
                        else if (!list.ContainsKey(updatedIds[i]))
                        {
                            threatField.Field.Name = " - ";
                            threatField.Field.Description = " - ";
                            threatField.Field.Source = " - ";
                            threatField.Field.Target = " - ";
                            threatField.Field.BreachOfConfid = " - ";
                            threatField.Field.IntegrityViolation = " - ";
                            threatField.Field.AccessibilityViolation = " - ";


                            threatField.UpdatedField.Name = updatedlist[updatedIds[i]].Name;
                            threatField.UpdatedField.Description = updatedlist[updatedIds[i]].Description;
                            threatField.UpdatedField.Source = updatedlist[updatedIds[i]].Source;
                            threatField.UpdatedField.Target = updatedlist[updatedIds[i]].Target;
                            threatField.UpdatedField.BreachOfConfid = updatedlist[updatedIds[i]].BreachOfConfid;
                            threatField.UpdatedField.IntegrityViolation = updatedlist[updatedIds[i]].IntegrityViolation;
                            threatField.UpdatedField.AccessibilityViolation = updatedlist[updatedIds[i]].AccessibilityViolation;
                        }
                        else if (!updatedlist.ContainsKey(updatedIds[i]))
                        {


                            threatField.Field.Name = list[updatedIds[i]].Name;
                            threatField.Field.Description = list[updatedIds[i]].Description;
                            threatField.Field.Source = list[updatedIds[i]].Source;
                            threatField.Field.Target = list[updatedIds[i]].Target;
                            threatField.Field.BreachOfConfid = list[updatedIds[i]].BreachOfConfid;
                            threatField.Field.IntegrityViolation = list[updatedIds[i]].IntegrityViolation;
                            threatField.Field.AccessibilityViolation = list[updatedIds[i]].AccessibilityViolation;

                            threatField.UpdatedField.Name = "Запись удалена!";
                            threatField.UpdatedField.Description = "Запись удалена!";
                            threatField.UpdatedField.Source = "Запись удалена!";
                            threatField.UpdatedField.Target = "Запись удалена!";
                            threatField.UpdatedField.BreachOfConfid = "Запись удалена!";
                            threatField.UpdatedField.IntegrityViolation = "Запись удалена!";
                            threatField.UpdatedField.AccessibilityViolation = "Запись удалена!";

                        }
                        updatedfields.Add(threatField);


                    }
                    UpdatedThreat.Visibility = Visibility.Visible;

                    UpdatedThreat.ItemsSource = updatedfields;
                   
                }
                else
                {
                    UpdatedThreat.Visibility = Visibility.Collapsed;
                }
                File.Delete(path);
                File.Copy(Directory.GetCurrentDirectory() + "\\UpdatedThreatBase.xlsx", path);
                File.Delete(Directory.GetCurrentDirectory() + "\\UpdatedThreatBase.xlsx");
                File.WriteAllText(Directory.GetCurrentDirectory() + "\\updatedate.txt", DateTime.Now.ToString());
                UpdateDate.Text = "Последнее обновление: " + File.ReadAllText(Directory.GetCurrentDirectory() + "\\updatedate.txt");
                StartMessage.Text = $"На вашем компьютере существует загруженная база угроз ИБ\n";
                StartMessage.Text += $"Всего записей в базе данных: {list.Count}";

            }
            catch (Exception ex)
            {
               
                UpdateStatus.Text = "Ошибка!";
                MessageBox.Show("Невозможно обновить базу данных!");
                UpdateStatus.Visibility = Visibility.Visible;
                UpdateMessage.Visibility = Visibility.Visible;
                UpdateMessage.Text = ex.Message;
            }
            if (updatedlist.Count != 0)
            {
                list.Clear();
                foreach (var item in updatedlist)
                {
                    list.Add(item.Key, item.Value);
                }


            }
            updatedlist.Clear();
           
           

        }

        static int min = 60;

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            var timer = new System.Windows.Threading.DispatcherTimer();
            timer.Interval = new TimeSpan(0, 0, 1);
            timer.IsEnabled = true;
            timer.Tick += (o, t) => 
            {
                if (list.Count!=0)
                {
                    if (min != 0)
                    {

                        Timer.Text = $"\nСледующее автоматическое обновление через {min} минут";
                    }
                    if (min == 0)
                    {
                        UpdateButton_Click(sender, e);
                        UpdateStatus.Visibility = Visibility.Collapsed;
                        UpdateMessage.Visibility = Visibility.Collapsed;
                        UpdatedThreat.Visibility = Visibility.Collapsed;

                    }
                    min--;
                }
            };
            timer.Start();
        }

        private void Save_Click(object sender, RoutedEventArgs e)
        {
            if (File.Exists(path))
            {
                if (File.Exists(Directory.GetCurrentDirectory() + "\\SavedThreatBase.xlsx"))
                {
                    File.Delete(Directory.GetCurrentDirectory() + "\\SavedThreatBase.xlsx");
                }
                File.Copy(path, Directory.GetCurrentDirectory() + "\\SavedThreatBase.xlsx");
                File.Delete(path);
                MessageBox.Show("База данных успешно сохраненa!");
            }else if(File.Exists(Directory.GetCurrentDirectory() + "\\SavedThreatBase.xlsx"))
            {
                MessageBox.Show("Текущая версия базы данных уже сохранена");
            }
        }

        private void Window_Closed(object sender, EventArgs e)
        {
            File.Delete(path);
        }
    }

}

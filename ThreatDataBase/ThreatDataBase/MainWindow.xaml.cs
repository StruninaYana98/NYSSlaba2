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

        static Dictionary<int, ThreatInfo> list = new Dictionary<int, ThreatInfo>();//хранит полную информацию

        static Dictionary<int, ThreatInfo> updatedlist = new Dictionary<int, ThreatInfo>();//хранит обновленную полную информацию
        static List<int> updatedIds = new List<int>();//список id обновленных записей
        static ObservableCollection<UpdatedThreatField> updatedfields = new ObservableCollection<UpdatedThreatField>();//для вывода обновленной информации в таблицу

        static ObservableCollection<ThreatShortInfo> shortlist = new ObservableCollection<ThreatShortInfo>();//вся краткая информация для таблицы
        static ObservableCollection<ThreatShortInfo> bufferlist = new ObservableCollection<ThreatShortInfo>();//краткая информация для одной страницы

        //флаги для смены режимов
        static bool isViewAllopen = false;
        static bool isViewOneopen = false;
        static bool isUpdateopen = false;

        static int page = 0;//номер страницы таблицы с краткой информацией


        static int min = 60;//счетчик автообновления
        public MainWindow()
        {
            InitializeComponent();
            HideElements();


            StartMessage.Visibility = Visibility.Visible;
            Img.Visibility = Visibility.Visible;

            //-----------Загрузка БД из файла SavedThreatBase, если он существует--------------

            if (File.Exists(Directory.GetCurrentDirectory() + "\\SavedThreatBase.xlsx"))
            {
                FromExcelToList(Directory.GetCurrentDirectory() + "\\SavedThreatBase.xlsx", list);
                StartMessage.Text = $"На вашем компьютере существует загруженная база угроз ИБ\n";
                StartMessage.Text += $"Всего записей в базе данных: {list.Count}";
                if (File.Exists(Directory.GetCurrentDirectory() + "\\savedupdatedate.txt"))
                {
                    UpdateDate.Text = "Последнее обновление: " + File.ReadAllText(Directory.GetCurrentDirectory() + "\\savedupdatedate.txt");
                }
                Timer.Text = $"\nСледующее автоматическое обновление через {min} минут";

            }
            //--------------Кнопка для первичной загрузки, если файла SavedThreatBase нет---------------
            else
            {
                StartMessage.Text = $"На Вашем компьютере отсутствует база данных по УБИ\nПожалуйста, загрузите базу данных\n";
                FirstUpdate.Visibility = Visibility.Visible;
            }
        }


        private void FirstUpdate_Click(object sender, RoutedEventArgs e)   // Кнопка первичной загрузки
        {

            try
            {
                wc.DownloadFile("https://bdu.fstec.ru/files/documents/thrlist.xlsx", Directory.GetCurrentDirectory() + "\\ThreatBase.xlsx");
                FromExcelToList(Directory.GetCurrentDirectory() + "\\ThreatBase.xlsx", list);
                FirstUpdate.Visibility = Visibility.Collapsed;

                StartMessage.Text = $"На вашем компьютере существует загруженная база угроз ИБ\n";
                StartMessage.Text += $"Всего записей в базе данных: {list.Count}";

                File.WriteAllText(Directory.GetCurrentDirectory() + "\\updatedate.txt", DateTime.Now.ToString()); // Запись даты-времени обновления
                UpdateDate.Text = "Последнее обновление: " + File.ReadAllText(Directory.GetCurrentDirectory() + "\\updatedate.txt");
                Timer.Text = $"\nСледующее автоматическое обновление через {min} минут";

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        private void FromExcelToList(string p, Dictionary<int, ThreatInfo> l) // Запись информации их excel в списки
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
                l.Add(id, ti);  // Запись полной информации в список
                ThreatShortInfo shortti = new ThreatShortInfo("УБИ." + id, name);
                shortlist.Add(shortti); // Запись краткой информации в список для вывода в таблицу 


            }
            ObjWorkBook.Close(false, Type.Missing, Type.Missing);
            ObjWorkExcel.Quit();
            GC.Collect();


        }

        private void ViewAll_Click(object sender, RoutedEventArgs e) // Кнопка "Просмотр всех УБИ"
        {
            if (list.Count != 0)
            {
                HideElements();
                //-------Закрытие остальных страниц-------
                isViewOneopen = false;
                isUpdateopen = false;

                if (isViewAllopen == false)
                {
                    //------Открытие страницы-----
                    isViewAllopen = true;

                    //------Показ всех UI  элементов страницы----
                    Rect2.Visibility = Visibility.Visible;
                    ThreatsList.Visibility = Visibility.Visible;
                    Prev.Visibility = Visibility.Visible;
                    Next.Visibility = Visibility.Visible;
                    Pages.Visibility = Visibility.Visible;

                    //--------Вывод первых 20 (или менее) записей в таблицу---------
                    bufferlist.Clear();
                    if (shortlist.Count >= 20)
                    {
                        for (int i = 0; i < 20; i++)
                        {
                            bufferlist.Add(shortlist[i]);
                        }
                    }
                    else
                    {
                        for (int i = 0; i < shortlist.Count; i++)
                        {
                            bufferlist.Add(shortlist[i]);
                        }
                    }
                    ThreatsList.ItemsSource = bufferlist;
                    Pages.Text = $"1 - {bufferlist.Count} из " + shortlist.Count;
                }
                else
                {
                    //-------Закрытие страницы-------
                    isViewAllopen = false;

                    StartMessage.Visibility = Visibility.Visible;
                    Img.Visibility = Visibility.Visible;

                }
            }
            else
            {
                MessageBox.Show("Пожалуйста, произведите первичную загрузку базы данных!");
            }

        }

        private void EnterIdButton_Click(object sender, RoutedEventArgs e) // Кнопка для ввода ID записи для показа полной информации
        {
            View.Visibility = Visibility.Collapsed;
            string s = EnterId.Text;
            int id;
            ObservableCollection<ThreatField> fields = new ObservableCollection<ThreatField>(); // Список для вывода полной информации по одной записи в таблицу

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
                    EnterId.Text = "";
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
        private void HideElements() // Скрытие всех ненужных элементов
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

        private void ViewOne_Click(object sender, RoutedEventArgs e) // Кнопка "Просмотр полной информации по УБИ
        {
            if (list.Count != 0)
            {
                HideElements();

                //-------Зактырие остальных страниц---------
                isViewAllopen = false;
                isUpdateopen = false;

                EnterId.Text = "";
                if (isViewOneopen == false)

                {
                    isViewOneopen = true;

                    //--------Показ всех элементов страницы---------
                    EnterId.Visibility = Visibility.Visible;
                    EnterIdButton.Visibility = Visibility.Visible;
                    EnterIdMessage.Visibility = Visibility.Visible;
                    Rect3.Visibility = Visibility.Visible;

                }
                else
                {

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

        private void Prev_Click(object sender, RoutedEventArgs e) // Кнопка пагинации "назад"
        {
            if (page > 0)
            {
                page--;
                bufferlist.Clear();
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

        private void Next_Click(object sender, RoutedEventArgs e)  // Кнопка пагинации "вперед"
        {
            if ((page + 1) * 20 <= shortlist.Count - 1)
            {
                page++;
                bufferlist.Clear();
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

        private void Update_Click(object sender, RoutedEventArgs e) // Кнопка "Обновить базу данных"
        {
            if (list.Count != 0)
            {
                HideElements();

                //-------Закрытие остальных страниц--------
                isViewAllopen = false;
                isViewOneopen = false;

                if (isUpdateopen == false)
                {
                    isUpdateopen = true;

                    //-------Показывает таблицу с последними обновлениями, если они были--------
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

        private void UpdateButton_Click(object sender, RoutedEventArgs e) // Кнопка "Обновить"
        {
            updatedfields.Clear();

            //--------Отсчет для автообновления начинается заново--------
            min = 60;
            Timer.Text = $"\nСледующее автоматическое обновление через {min} минут";

            try
            {
                //------Скачиваем в другой файл, чтобы в случае ошибки предыдущее успешное обновление осталось--------
                wc.DownloadFile("https://bdu.fstec.ru/files/documents/thrlist.xlsx", Directory.GetCurrentDirectory() + "\\UpdatedThreatBase.xlsx");

                updatedIds.Clear();
                int count = 0;
                FromExcelToList(Directory.GetCurrentDirectory() + "\\UpdatedThreatBase.xlsx", updatedlist);

                //-----------Находим ID всех измененных записей------
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

                        //---------Если была изменена существующая запись---------
                        if (list.ContainsKey(updatedIds[i]) && updatedlist.ContainsKey(updatedIds[i]))
                        {

                            threatField.Fields.Name = "нет изменений";
                            threatField.UpdatedFields.Name = "нет изменений";

                            threatField.Fields.Description = "нет изменений";
                            threatField.UpdatedFields.Description = "нет изменений";

                            threatField.Fields.Source = "нет изменений";
                            threatField.UpdatedFields.Source = "нет изменений";

                            threatField.Fields.Target = "нет изменений";
                            threatField.UpdatedFields.Target = "нет изменений";

                            threatField.Fields.BreachOfConfid = "нет изменений";
                            threatField.UpdatedFields.BreachOfConfid = "нет изменений";


                            threatField.Fields.IntegrityViolation = "нет изменений";
                            threatField.UpdatedFields.IntegrityViolation = "нет изменений";

                            threatField.Fields.AccessibilityViolation = "нет изменений";
                            threatField.UpdatedFields.AccessibilityViolation = "нет изменений";

                            if (list[updatedIds[i]].Name != updatedlist[updatedIds[i]].Name)
                            {
                                threatField.Fields.Name = list[updatedIds[i]].Name;
                                threatField.UpdatedFields.Name = updatedlist[updatedIds[i]].Name;
                            }
                            if (list[updatedIds[i]].Description != updatedlist[updatedIds[i]].Description)
                            {
                                threatField.Fields.Description = list[updatedIds[i]].Description;
                                threatField.UpdatedFields.Description = updatedlist[updatedIds[i]].Description;
                            }
                            if (list[updatedIds[i]].Source != updatedlist[updatedIds[i]].Source)
                            {
                                threatField.Fields.Source = list[updatedIds[i]].Source;
                                threatField.UpdatedFields.Source = updatedlist[updatedIds[i]].Source;
                            }
                            if (list[updatedIds[i]].Target != updatedlist[updatedIds[i]].Target)
                            {
                                threatField.Fields.Target = list[updatedIds[i]].Target;
                                threatField.UpdatedFields.Target = updatedlist[updatedIds[i]].Target;
                            }
                            if (list[updatedIds[i]].BreachOfConfid != updatedlist[updatedIds[i]].BreachOfConfid)
                            {
                                threatField.Fields.BreachOfConfid = list[updatedIds[i]].BreachOfConfid;
                                threatField.UpdatedFields.BreachOfConfid = updatedlist[updatedIds[i]].BreachOfConfid;
                            }
                            if (list[updatedIds[i]].IntegrityViolation != updatedlist[updatedIds[i]].IntegrityViolation)
                            {
                                threatField.Fields.IntegrityViolation = list[updatedIds[i]].IntegrityViolation;
                                threatField.UpdatedFields.IntegrityViolation = updatedlist[updatedIds[i]].IntegrityViolation;
                            }
                            if (list[updatedIds[i]].AccessibilityViolation != updatedlist[updatedIds[i]].AccessibilityViolation)
                            {
                                threatField.Fields.AccessibilityViolation = list[updatedIds[i]].AccessibilityViolation;
                                threatField.UpdatedFields.AccessibilityViolation = updatedlist[updatedIds[i]].AccessibilityViolation;
                            }
                        }
                        //---------Если была добавлена новая запись---------
                        else if (!list.ContainsKey(updatedIds[i]))
                        {
                            threatField.Fields.Name = " - ";
                            threatField.Fields.Description = " - ";
                            threatField.Fields.Source = " - ";
                            threatField.Fields.Target = " - ";
                            threatField.Fields.BreachOfConfid = " - ";
                            threatField.Fields.IntegrityViolation = " - ";
                            threatField.Fields.AccessibilityViolation = " - ";

                            threatField.UpdatedFields.Name = updatedlist[updatedIds[i]].Name;
                            threatField.UpdatedFields.Description = updatedlist[updatedIds[i]].Description;
                            threatField.UpdatedFields.Source = updatedlist[updatedIds[i]].Source;
                            threatField.UpdatedFields.Target = updatedlist[updatedIds[i]].Target;
                            threatField.UpdatedFields.BreachOfConfid = updatedlist[updatedIds[i]].BreachOfConfid;
                            threatField.UpdatedFields.IntegrityViolation = updatedlist[updatedIds[i]].IntegrityViolation;
                            threatField.UpdatedFields.AccessibilityViolation = updatedlist[updatedIds[i]].AccessibilityViolation;
                        }
                        //---------Если существующая запись была удалена-----------
                        else if (!updatedlist.ContainsKey(updatedIds[i]))
                        {
                            threatField.Fields.Name = list[updatedIds[i]].Name;
                            threatField.Fields.Description = list[updatedIds[i]].Description;
                            threatField.Fields.Source = list[updatedIds[i]].Source;
                            threatField.Fields.Target = list[updatedIds[i]].Target;
                            threatField.Fields.BreachOfConfid = list[updatedIds[i]].BreachOfConfid;
                            threatField.Fields.IntegrityViolation = list[updatedIds[i]].IntegrityViolation;
                            threatField.Fields.AccessibilityViolation = list[updatedIds[i]].AccessibilityViolation;

                            threatField.UpdatedFields.Name = "Запись удалена!";
                            threatField.UpdatedFields.Description = "Запись удалена!";
                            threatField.UpdatedFields.Source = "Запись удалена!";
                            threatField.UpdatedFields.Target = "Запись удалена!";
                            threatField.UpdatedFields.BreachOfConfid = "Запись удалена!";
                            threatField.UpdatedFields.IntegrityViolation = "Запись удалена!";
                            threatField.UpdatedFields.AccessibilityViolation = "Запись удалена!";

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

                //--------Перезаписываем файл ThreatBase и дату-время для запоминания последнего обновления до конца текущего сеанса программы-----------
                File.Delete(Directory.GetCurrentDirectory() + "\\ThreatBase.xlsx");
                File.Copy(Directory.GetCurrentDirectory() + "\\UpdatedThreatBase.xlsx", Directory.GetCurrentDirectory() + "\\ThreatBase.xlsx");
                File.Delete(Directory.GetCurrentDirectory() + "\\UpdatedThreatBase.xlsx");
                File.WriteAllText(Directory.GetCurrentDirectory() + "\\updatedate.txt", DateTime.Now.ToString());

                //---------Обновляем данные---------
                if (updatedlist.Count != 0)
                {
                    list.Clear();
                    foreach (var item in updatedlist)
                    {
                        list.Add(item.Key, item.Value);
                    }
                }

                UpdateDate.Text = "Последнее обновление: " + File.ReadAllText(Directory.GetCurrentDirectory() + "\\updatedate.txt");
                StartMessage.Text = $"На вашем компьютере существует загруженная база угроз ИБ\n";
                StartMessage.Text += $"Всего записей в базе данных: {list.Count}";

            }
            catch (Exception ex)
            {
                UpdateStatus.Visibility = Visibility.Visible;
                UpdateMessage.Visibility = Visibility.Visible;
                UpdateStatus.Text = "Ошибка!";
                UpdateMessage.Text = ex.Message;
                MessageBox.Show("Невозможно обновить базу данных!");
            }
            updatedlist.Clear();
        }



        private void Window_Loaded(object sender, RoutedEventArgs e) // Запуск таймера автообновления при загрузке окна
        {
            var timer = new System.Windows.Threading.DispatcherTimer();
            timer.Interval = new TimeSpan(0, 1, 0);
            timer.IsEnabled = true;

            timer.Tick += (o, t) =>
            {
                if (list.Count != 0)
                {
                    min--;
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

                }
            };
            timer.Start();

        }

        private void Save_Click(object sender, RoutedEventArgs e) // Кнопка "Сохранить базу УБИ"
        {
            //------Если есть обновление, сохраняем его---------
            if (File.Exists(Directory.GetCurrentDirectory() + "\\ThreatBase.xlsx"))
            { 
                File.Delete(Directory.GetCurrentDirectory() + "\\SavedThreatBase.xlsx");
                File.Copy(Directory.GetCurrentDirectory() + "\\ThreatBase.xlsx", Directory.GetCurrentDirectory() + "\\SavedThreatBase.xlsx");

                File.Delete(Directory.GetCurrentDirectory() + "\\savedupdatedate.txt");
                File.Copy(Directory.GetCurrentDirectory() + "\\updatedate.txt", Directory.GetCurrentDirectory() + "\\savedupdatedate.txt");

                MessageBox.Show("База данных успешно сохраненa!");
            }
            else if (File.Exists(Directory.GetCurrentDirectory() + "\\SavedThreatBase.xlsx"))
            {
                MessageBox.Show("Текущая версия базы данных уже сохранена");
            }
        }

        private void Window_Closed(object sender, EventArgs e)
        {
            File.Delete(Directory.GetCurrentDirectory() + "\\ThreatBase.xlsx");
            File.Delete(Directory.GetCurrentDirectory() + "\\updatedate.txt");
        }
    }

}

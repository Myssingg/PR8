using System;
using System.Linq;
using System.Runtime.Remoting.Contexts;
using System.Windows;
using System.Windows.Controls;

using Word = Microsoft.Office.Interop.Word;

namespace BorDer_PR8
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        BorDer_PR8Entities db = new BorDer_PR8Entities();


        Stationery selectedItem;
        int Count = -1;
        int Discount = -1;
        double Summ = 0; //Сумма покупки
        double Result = 0; 


        public MainWindow()
        {
            InitializeComponent();
            var Stationaries = db.Stationeries.AsNoTracking().ToList();
            comboBox_stationaries.ItemsSource = Stationaries;
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            if ((sender as Button).Name == "btn_Result") //Рассчитать
            {
                try
                {
                    double summTaken = double.Parse(textBox_SummTaken.Text);

                    int Discount = int.Parse(textBox_Discount.Text);
                    if (Discount < 0 || Discount > 100)
                    {
                        throw new Exception("Скидка должна быть АДЕКВАТНОЙ");
                    }
                    Result = Summ - (Summ * Discount / 100);
                    double Change = summTaken - Result;

                    if (Change < 0)
                    {
                        throw new Exception("Сумма покупки не должна превышать принятую сумму! (С учетом скидок)");
                    }
                    textBox_Change.Text = Change.ToString();

                    this.Discount = Discount;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            else //Сохранить
            {
                using (BorDer_PR8Entities tempDB = new BorDer_PR8Entities())
                {

                    if (selectedItem != null && Count > 0 && Discount >= 0 && Result > 0)
                    {
                        var Sale = new Sale
                        {
                            Date = DateTime.Now,
                            IdStationery = selectedItem.Id,
                            Count = Count,
                            Discount = Discount,
                            Summary = Result
                        };
                        MessageBox.Show(Sale.ToString());
                        tempDB.Sales.Add(Sale);
                        try
                        {
                            tempDB.SaveChanges();
                            btn_Check.Visibility = Visibility.Visible;
                        }
                        catch (Exception ex)
                        {
                            btn_Check.Visibility = Visibility.Collapsed;
                            MessageBox.Show("Ошибка сохранения \n" + ex.Message);
                        }


                    }
                    else
                    {
                        MessageBox.Show("Не заполнены поля, либо заполнены неверно!");
                    }
                }
            }

        }
        private void comboBox_stationaries_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            selectedItem = comboBox_stationaries.SelectedItem as Stationery;
            fillSumm();
        }

        private void textBox_Count_TextChanged(object sender, TextChangedEventArgs e)
        {
            fillSumm();
        }

        private void fillSumm()
        {
            var selected = comboBox_stationaries.SelectedItem as Stationery;

            var itemPrice = selected.Price;

            try
            {
                if (!String.IsNullOrEmpty(textBox_Count.Text))
                {
                    var Count = int.Parse(textBox_Count.Text);
                    if (Count < 0) { MessageBox.Show("Количество товаров не может быть меньше нуля!"); return; }
                    Summ = itemPrice * Count;
                    this.Count = Count;
                    textBox_Summ.Text = (Summ).ToString();
                }
            }
            catch {}

        }

        private void Repwo(string subToReplace, string text, Word.Document worddoc)
        {
            var range = worddoc.Content;
            range.Find.ClearFormatting();
            range.Find.Execute(FindText: subToReplace, ReplaceWith: text);
        }

        private void saveWord(int number, string Name)
        {
            var WordApp = new Word.Application();
            WordApp.Visible = false;

            var Worddoc = WordApp.Documents.Open(Environment.CurrentDirectory + @"\Чек.docx");

            Repwo("{id}", number.ToString(), Worddoc);
            Repwo("{date}", DateTime.Now.ToString(), Worddoc);
            Repwo("{name}", Name, Worddoc);
            Repwo("{result}", Result.ToString(), Worddoc);
            Repwo("{count}", Count.ToString(), Worddoc);
            Repwo("{discount}", Discount + "%", Worddoc);
            Repwo("{summ}", Summ.ToString(), Worddoc);
            Worddoc.SaveAs2(Environment.CurrentDirectory + $@"\Чек_{number.ToString()}.docx");
            MessageBox.Show("Билет сохранен!");
        }

        private void btn_Check_Click(object sender, RoutedEventArgs e)
        {
            var lastWrite = db.Sales.OrderByDescending(s => s.Id).FirstOrDefault();
            saveWord(lastWrite.Id++, selectedItem.Name);
        }
    }
}

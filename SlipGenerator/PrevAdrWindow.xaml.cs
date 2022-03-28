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
using System.Windows.Shapes;

namespace SlipGenerator
{
    /// <summary>
    /// Interaction logic for PrevAdrWindow.xaml
    /// </summary>
    public partial class PrevAdrWindow : Window
    {
        private readonly Settings _settings = Settings.Default;

        public PrevAdrWindow()
        {
            
            InitializeComponent();
            var arr = _settings.Adr.Split(',').ToList();
            if (arr.Count > 1)
            {
                var items = new List<ListBoxItem>();
                arr.ForEach(str =>
                {
                    var item = new ListBoxItem
                    {
                        Content = str
                    };
                    items.Add(item);
                    item.MouseDoubleClick += AdrListBoxItemClick;
                });
                AdrListBox.ItemsSource = items;
            }

            void AdrListBoxItemClick(object sender, RoutedEventArgs e)
            {
                var item = AdrListBox.SelectedItem as ListBoxItem;
                _settings.SelAdr = item?.Content.ToString();
                _settings.Save();
                GetWindow(this)?.Close();
            }

        }
        void BtnAdrRemoveClick(object sender, RoutedEventArgs e)
        {
            _settings.Adr = "";
            _settings.Save();
            GetWindow(this)?.Close();
        }

    }
}
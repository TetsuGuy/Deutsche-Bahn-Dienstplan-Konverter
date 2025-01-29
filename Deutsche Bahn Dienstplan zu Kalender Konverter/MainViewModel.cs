using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;

namespace Deutsche_Bahn_Dienstplan_zu_Kalender_Konverter
{
    internal class MainViewModel : INotifyPropertyChanged
    {
        private string _greeting;
        public string Greeting
        {
            get => _greeting;
            set
            {
                _greeting = value;
                OnPropertyChanged(nameof(Greeting));
                Console.WriteLine("success");
            }
        }

    public MainViewModel()
        {
            Greeting = "Hello, MVVM!";
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected void OnPropertyChanged(string propertyName) {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}

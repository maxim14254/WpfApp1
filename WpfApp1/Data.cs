using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

namespace WpfApp1
{
    public class Data : INotifyPropertyChanged
    {
        string name;
        double distance;
        double angle;
        double width;
        double hegth;
        string isDefect;

        public Data(string name, string distance, string angle, string width, string hegth, string isDefect)
        {
            this.name = name;
            this.distance = double.Parse(distance);
            this.angle = double.Parse(angle);
            this.width = double.Parse(width);
            this.hegth = double.Parse(hegth);
            this.isDefect = isDefect;
        }

        public string Name
        {
            get
            {
                return name;
            }
            set
            {
                name = value;
                OnPropertyChanged("Name");
                OnPropertyChanged("AllSelectString");
            }
        }
        public double Distance
        {
            get
            {
                return distance;
            }
            set
            {
                distance = Validation(value, 20.0);
                OnPropertyChanged("Distance");
                OnPropertyChanged("AllSelectString");
            }
        }

        public double Angle
        {
            get
            {
                return angle;
            }
            set
            {
                angle = Validation(value, 12.0);
                OnPropertyChanged("Angle");
                OnPropertyChanged("AllSelectString");
            }
        }

        public double Width
        {
            get
            {
                return width;
            }
            set
            {
                width = Validation(value, 20.0);
                OnPropertyChanged("Width");
                OnPropertyChanged("AllSelectString");
            }
        }

        public double Hegth
        {
            get
            {
                return hegth;
            }
            set
            {
                hegth = Validation(value, 12.0);
                OnPropertyChanged("Hegth");
                OnPropertyChanged("AllSelectString");
            }
        }

        public string IsDefect
        {
            get
            {
                return isDefect;
            }
            set
            {
                isDefect = value;
                OnPropertyChanged(value);
                OnPropertyChanged("AllSelectString");
            }
        }

        public string AllSelectString
        {
            get
            {
                return $"{name}  {distance}  {angle}  {width}  {hegth}  {isDefect}";
            }
            set
            {
                OnPropertyChanged("AllSelectString");
            }
        }

        double Validation(double value, double maxValue)
        {
            if (value < 0)
                return 0;
            if (value > maxValue)
                return maxValue;
            return value;
        }

        public event PropertyChangedEventHandler PropertyChanged;

        public bool OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
                return true;
            }
            return false;
        }
    }
}

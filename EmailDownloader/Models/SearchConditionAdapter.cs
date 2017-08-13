using MahApps.Metro.Controls;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;

namespace EmailDownloader.Models
{
    internal class SearchConditionAdapter : IEquatable<SearchConditionAdapter>, INotifyPropertyChanged
    {
        public readonly MethodInfo Method;

        private List<Control> paramEditors;
        public ObservableCollection<Control> ParamEditors => new ObservableCollection<Control>(paramEditors);

        private List<ParameterInfo> paramsInfo;

        public event PropertyChangedEventHandler PropertyChanged;

        public SearchConditionAdapter(MethodInfo conditionGenerateMethod)
        {
            Method = conditionGenerateMethod;
            paramEditors = new List<Control>();
            paramsInfo = new List<ParameterInfo>();
            foreach (var p in Method.GetParameters())
            {
                var type = p.ParameterType;

                if (type.Equals(typeof(long)))
                {
                    var c = new NumericUpDown()
                    {
                        HasDecimals = false,
                        Maximum = long.MaxValue,
                        Minimum = long.MinValue,
                        Value = 1
                    };

                    paramEditors.Add(c);
                }
                else if (type.Equals(typeof(uint)))
                {
                    var c = new NumericUpDown()
                    {
                        HasDecimals = false,
                        Maximum = uint.MaxValue,
                        Minimum = uint.MinValue,
                        Value = 1
                    };

                    paramEditors.Add(c);
                }
                else if (type.Equals(typeof(string)))
                {
                    paramEditors.Add(new TextBox());
                }
                else if (type.Equals(typeof(DateTime)))
                {
                    paramEditors.Add(new DatePicker());
                }

                paramsInfo.Add(p);
            }
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(this.ParamEditors)));
        }

        public object[] GetValues()
        {
            var row = new object[paramEditors.Count];
            for(int i = 0; i < row.Length; i++)
            {
                switch (paramEditors[i])
                {
                    case NumericUpDown nud:
                        if (paramsInfo[i].ParameterType.Equals(typeof(long)))
                            row[i] = long.Parse(nud.Invoke(() => nud.Value.ToString()));
                        else
                            row[i] = uint.Parse(nud.Invoke(() => nud.Value.ToString()));
                        break;
                    case TextBox tb:
                        row[i] = tb.Invoke(() => tb.Text);
                        break;
                    case DatePicker dtp:
                        row[i] = dtp.Invoke(() => dtp.SelectedDate);
                        break;
                }
            }
            return row;
        }

        public bool Equals(SearchConditionAdapter other) => Method.Equals(other.Method);
    }
}

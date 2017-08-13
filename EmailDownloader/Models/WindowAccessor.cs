using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace EmailDownloader.Models
{
    internal static class WindowAccessor
    {
        public static Window Main => Application.Current.MainWindow;
    }
}

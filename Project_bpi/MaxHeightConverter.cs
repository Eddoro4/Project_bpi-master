using System;
using System.Globalization;
using System.Windows.Data;

namespace Project_bpi
{
    public class MaxHeightConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is double actualHeight && actualHeight > 0)
            {
                // Резервируем место под:
                double header = 80;       // Заголовок
                double topMargin = 20;    // StackPanel.Margin.Top
                double bottomMargin = 20; // StackPanel.Margin.Bottom
                double textBlockMargin = 100; // Margin у Border с текстом (сверху+снизу = 10+10)
                double editButtonArea = 50;  // Кнопка "Редактировать" + отступ
                double minContentHeight = 100;

                double reserved = header + topMargin + bottomMargin + textBlockMargin + editButtonArea;
                double available = actualHeight - reserved;

                return Math.Max(minContentHeight, available);
            }
            return 350.0; // fallback
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
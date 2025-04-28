using Microsoft.IdentityModel.Tokens;
using System.Globalization;

namespace TechReportToDB.Converters
{
    internal static class CC
    {
        public static string? ConvertStringToDateTimeString(string? str)
        {
            if (str.IsNullOrEmpty()) return null;

            try
            {
                return Convert.ToDateTime(str).ToShortDateString();
            }
            catch (Exception)
            {

                return str ?? null;
            }
        }

        public static double? ConvertStringToDouble(string? str)
        {
            if (string.IsNullOrWhiteSpace(str)) return null;

            try
            {
                string decimalSeparator = CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator;

                str = str.Replace(",", decimalSeparator).Replace(".", decimalSeparator);

                if (double.TryParse(str, NumberStyles.Any, CultureInfo.CurrentCulture, out double result))
                {
                    return result;
                }
            }
            catch
            {
                return null;
            }

            return null;
        }

        public static int? ConvertStringToInt(string? str)
        {
            if (str.IsNullOrEmpty()) return null;

            try
            {
                return Convert.ToInt32(str);
            }
            catch (Exception)
            {

                return null;
            }
        }

        public static bool TryExtractDateFromSheetName(string sheetName, out DateTime date)
        {

            string[] dateFormats = { "dd.MM.yyyy", "yyyy-MM-dd", "MM-dd-yyyy" };

            foreach (string format in dateFormats)
            {
                if (DateTime.TryParseExact(sheetName, format, CultureInfo.InvariantCulture, DateTimeStyles.None, out date))
                {
                    return true;
                }
            }

            // Если не удается найти дату в названии листа
            date = DateTime.MinValue;
            return false;
        }
    }
}

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Reflection;
using ClosedXML.Excel;

namespace ExportExcel.Models
{
    public static class ExcelMethods
    {
        public static string UppercaseWords(string value)
        {
            string v = value.ToLower();

            return v.First().ToString().ToUpper() + v.Substring(1);
        }

        public static string[] CalculateExcelColumnName(int NumberOfColumns)
        {
            string[] sa = new string[] { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z" };
            string[] result = new string[NumberOfColumns];
            string s = string.Empty;
            int i, j, k, l;
            i = j = k = -1;
            for (l = 0; l < NumberOfColumns; ++l)
            {
                s = string.Empty;
                ++k;
                if (k == 26)
                {
                    k = 0;
                    ++j;
                    if (j == 26)
                    {
                        j = 0;
                        ++i;
                    }
                }
                if (i >= 0) s += sa[i];
                if (j >= 0) s += sa[j];
                if (k >= 0) s += sa[k];
                result[l] = s;
            }
            return result;
        }

        public static string FormatCell(string value)
        {
            var _format = "";
            if (String.IsNullOrEmpty(value))
                return _format;

            #region DateTime
            DateTime datetime;
            var isDate = DateTime.TryParse(value, out datetime);
            if (isDate)
            {
                _format = "dd/MM/yyyy hh:mm:ss";
            }
            #endregion

            return _format;
        }

        public static XLDataType GetDataTypeOfColumnValue(string value)
        {
            var _XLDataType = XLDataType.Text;

            if (String.IsNullOrEmpty(value))
                return _XLDataType;

            #region DateTime
            DateTime datetime;
            var isDate = DateTime.TryParse(value, out datetime);
            if (isDate)
            {
                _XLDataType = XLDataType.DateTime;
            }
            #endregion

            return _XLDataType;
        }

        /// <summary>
        /// Gets Excel value type used to store specified property
        /// </summary>
        /// <param name="property"></param>
        /// <returns></returns>
        public static XLDataType GetXLDataType(PropertyInfo property)
        {
            Type type = GetExcelDataType(property.PropertyType);
            XLDataType result;

            if (type == typeof(DateTime))
            {
                result = XLDataType.DateTime;
            }
            else if (type == typeof(decimal) || type == typeof(int) || type == typeof(long))
            {
                result = XLDataType.Number;
            }
            else if (type == typeof(TimeSpan))
            {
                result = XLDataType.TimeSpan;
            }
            else
            {
                result = XLDataType.Text;
            }

            return result;
        }

        /// <summary>
        /// Gets actual data type used to store the specified property in Excel
        /// </summary>
        /// <param name="type"></param>
        /// <returns></returns>
        public static Type GetExcelDataType(Type type)
        {
            Type underlyingType = Nullable.GetUnderlyingType(type) ?? type;
            Type result = underlyingType;

            if (underlyingType != typeof(DateTime)
                && underlyingType != typeof(TimeSpan)
                && underlyingType != typeof(decimal)
                && underlyingType != typeof(int)
                && underlyingType != typeof(long))
            {
                // enum, string & other objects
                result = typeof(string);
            }

            return result;
        }


        /// <summary>
        /// Converter a lista do tipo de modelo para um Datatable
        /// </summary>
        /// <typeparam name="T">O tipo de modelo</typeparam>
        /// <param name="data">Lista do tipo de modelo</param>
        /// <returns>retorna um datatable</returns>
        public static DataTable ToDataTable<T>(this IList<T> data)
        {
            if (data is null) return null;

            PropertyDescriptorCollection properties = TypeDescriptor.GetProperties(typeof(T));

            DataTable datatable = new DataTable();

            foreach (PropertyDescriptor prop in properties)
                datatable.Columns.Add(prop.Name, Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType);

            foreach (T item in data)
            {
                DataRow row = datatable.NewRow();
                foreach (PropertyDescriptor prop in properties)
                    row[prop.Name] = prop.GetValue(item) ?? DBNull.Value;

                datatable.Rows.Add(row);
            }

            return datatable;
        }

        /// <summary>
        /// Converter a lista do tipo de modelo para um DataSet
        /// </summary>
        /// <typeparam name="T">O tipo de modelo</typeparam>
        /// <param name="data">Lista do tipo de modelo</param>
        /// <returns>retorna um dataset</returns>
        public static DataSet ToDataSet<T>(this IList<T> data)
        {
            DataTable dt = data.ToDataTable<T>();
            if (dt is null) return null;

            DataSet ds = new DataSet();
            ds.Tables.Add(dt);
            return ds;
        }
    }
}

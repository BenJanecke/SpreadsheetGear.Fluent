using System;
using System.Collections.Generic;
using System.Linq;
using SpreadsheetGear;

namespace SpreadsheetGear.Fluent
{
    public static class Fluent
    {

        #region WorkSheets 

        public static IWorksheet SetMarginToNarrow(this IWorksheet ws)
        {
            ws.PageSetup.TopMargin = 0.75;
            ws.PageSetup.LeftMargin = 0.25;
            ws.PageSetup.HeaderMargin = 0.30000001192092896;
            ws.PageSetup.FooterMargin = 0.30000001192092896;
            ws.PageSetup.RightMargin = 0.25;
            ws.PageSetup.BottomMargin = 0.75;
          
            return ws;
        }

        public static IWorksheet SetLayout(this IWorksheet ws, bool landscape)
        {
           ws.PageSetup.Orientation = landscape ? PageOrientation.Landscape : PageOrientation.Portrait;
            
           return ws;
        }

        public static IWorksheet SetColWidths(this IWorksheet ws, params int[] cols)
        {
            var k = 0;
            foreach (var col in cols)
            {
                ws.FluentCells(0, k).ColumnWidth = col;
                k++;
            }  
            return ws;
        }

        public static IWorksheet FitColsToPages(this IWorksheet ws, int pages)
        {
            ws.PageSetup.FitToPagesWide = pages;

            return ws;
        }

        #endregion WorkSheets

        #region Cells
        public static IRange FluentCells(this IWorksheet worksheet, int row, int col)
        {
            return worksheet.Cells[row,col];
        }


        public static IRange ToggleAutoFilter(this IRange range)
        {
            range.AutoFilter();
            return range;
        }

        public static IRange FluentCells(this IWorksheet worksheet, int startingRow, int startingCol, int finalRow, int finalCol)
        {
            return worksheet.Cells[startingRow, startingCol, finalRow, finalCol];
        }

        public static IRange SetValue(this IRange range, object value, bool autoNumberFormat = false)
        {
            if (value != null)
            {
            //Dumps the object to cell range 
            if (!autoNumberFormat)
            {
                range.Value = value.ToString();
                   
                    range.SetNumberFormat(NumberFormat.Text);
                return range;
            }
                range.Value = value.ToString();
                range.NumberFormat = value.GetNumberFormat();
            }
            return range;
        }

        public static IRange SetNumberFormat(this IRange range, NumberFormat numberFormat)
        {
            range.NumberFormat = numberFormat.GetNumberFormat();
         
            return range;
        }

        public static IRange Merge(this IRange range, bool merge)
        {
            if (merge)
            {
                range.Merge();
            }
            else
            {
                range.UnMerge();
            }
            return range;
        }

        

        public static  IRange SetAlignment(this IRange range, VAlign vertical, HAlign horizontal)
        {
            range.VerticalAlignment = vertical;
            range.HorizontalAlignment = horizontal;
            range.Style.IncludeAlignment = true;
            return range;
        }

        public static IRange SetWidth(this IRange range, double width)
        {
            range.ColumnWidth = width;
            return range;
        }

        public static IRange SetHeight(this IRange range, double height)
        {
            range.RowHeight = height;
            return range;
        }

        public static IRange SetWrapText(this IRange range)
        {
            range.WrapText = true;
            return range;
        }

        public static IRange SetFontSize(this IRange range, int size)
        {
            range.Font.Size = size;
            return range;
        }


        #endregion Cells

        #region Styles 
        
        public static Color ToSpreadsheetGearColor(this System.Drawing.Color color)
        {
            return Color.FromArgb(color.ToArgb());
        }


        public static IRange SetStyle(this IRange range, IStyle style)
        {
            range.Style = style;
            return range;
        }

        public static IRange SetBorders(this IRange range, BordersIndex borders, LineStyle style, BorderWeight weight, Color color)
        {
            range.Borders[borders].LineStyle = style;
            range.Borders[borders].Weight = weight;
            range.Borders[borders].Color = color;
            range.Style.IncludeBorder = true;

            return range;
        }

        public static IRange SetBold(this IRange range, bool bold)
        {
            range.Font.Bold = bold;
            return range;
        }



        public static IRange SetBorders(this IRange range, LineStyle style, BorderWeight weight, Color color)
        {
            range.Borders.LineStyle = style;
            range.Borders.Weight = weight;
            range.Borders.Color = color;
            range.Style.IncludeBorder = true;

            return range;
        }
        #endregion Styles 

        #region Types

        public enum NumberFormat
        {
            General,
            Number,
            Currency,
            Accounting,
            ShortDate,
            LongDate,
            Time,
            Percentage,
            PercentageTrunc,
            Fraction,
            Scientific,
            Text,
            Days,
            Months,
            ShortNumber,
            Clean
        }

        private static readonly Dictionary<NumberFormat, string> NumberFormats = new Dictionary<NumberFormat, string>
                                                                                 {
                                                                                     {NumberFormat.Accounting, "R # ##0.00;[Red]R -# ##0.00"},
                                                                                     {NumberFormat.Currency, "R ### ### ##0.00;[Red]R -### ### ##0.00"},
                                                                                     {NumberFormat.Days, "0\" Days\""},
                                                                                     {NumberFormat.Fraction, ""},
                                                                                     {NumberFormat.General, ""},
                                                                                     {NumberFormat.LongDate, "[$-1C09]dd mmmm yyyy"},
                                                                                     {NumberFormat.Months, "0\" Months\""},
                                                                                     {NumberFormat.Number, "# ##0.00;[Red]-# ##0.00"},
                                                                                     {NumberFormat.Percentage, "0.00%"},
                                                                                     {NumberFormat.PercentageTrunc, "0%"},
                                                                                     {NumberFormat.Scientific, "0.00E+00"},
                                                                                     {NumberFormat.ShortDate, "dd/mm/yyyy"},
                                                                                     {NumberFormat.ShortNumber, "# ##0;[Red]-# ##0"},
                                                                                     {NumberFormat.Text, "@"},
                                                                                     {NumberFormat.Time, ""},
                                                                                     {NumberFormat.Clean,"#"}
                                                                                 };




        private static string GetNumberFormat(this NumberFormat key)
        {
            if (!NumberFormats.ContainsKey(key))
            {
                throw new InvalidOperationException(string.Format("Type {0} doesn't have a matching Type configured", key));
            }
            return NumberFormats[key];
        }

        private static string GetNumberFormat(this object value)
        {
            //Attempts to determine the type of the value and returns a appropriate number format
            if (value is DateTime)
            {
                return NumberFormat.ShortDate.GetNumberFormat();
            }
            if (value is double?)
            {
                return NumberFormat.ShortDate.GetNumberFormat();
            }
            return "";
        }
        #endregion Types    

    }
}
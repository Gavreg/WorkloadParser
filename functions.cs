using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using RtfDomParser;

namespace ConsoleApp11
{
    enum Mode { SEARCHING, AUD_PARCE, VA_PARCE }

    public static class functions
    {
        public static string GetCellText(RTFDomTableCell? cell)
        {
            if (cell == null) return "";
            if (cell.Elements.Count == 0) return "";
            if (cell.Elements[0].Elements.Count == 0) return "";
            string ret="";
            for (int i = 0; i < cell.Elements[0].Elements.Count; ++i)
            {
                var e = cell.Elements[0].Elements[i];
                ret += e.InnerText.Replace("?", "");

            }
            return ret;
        }

        public static string GetRowText(RTFDomTableRow? row)
        {
            if (row == null) return "";
            string ret = "|";

            for (int i=0; i<row.Elements.Count; ++i)
            {
                ret += GetCellText(row.Elements[i] as RTFDomTableCell);
                i += (row.Elements[i] as RTFDomTableCell).RowSpan;
            }
            return ret;
            
        }


        public static void PrintRow(RTFDomTableRow? row)
        {
            bool first = true;

            List<string> cell_strings = new List<string>();
            List<int> cell_lengths = new List<int>();

            for (int i = 0; i < row.Elements.Count; ++i)
            {
                var cell = row.Elements[i] as RTFDomTableCell;
                cell_strings.Add(GetCellText(cell));
                cell_lengths.Add(Math.Max(i.ToString().Length, GetCellText(cell).Length));
            }

            Console.Write("┌");
            first = true;
            for (int i = 0; i < cell_strings.Count; ++i)
            {
                var length = cell_strings[i].Length;
                
                if (length == 0) continue;
                
                if (!first) Console.Write('┬');
                Console.Write(new String('─', cell_lengths[i]));
                first = false;
     
            }
            Console.WriteLine('┐');


            Console.Write("│");
            first = true;
            for (int i = 0; i < cell_strings.Count; ++i)
            {
                var length = cell_strings[i].Length;

                if (length == 0) continue;

                if (!first) Console.Write('│');
                
                if (i.ToString().Length < cell_lengths[i])
                {
                    string number_str = i.ToString().PadRight(cell_lengths[i]/2);
                    Console.Write(number_str);
                    Console.Write(new String(' ', cell_lengths[i] - number_str.Length));
                }
                else
                {
                    Console.Write(i.ToString());
                }

                first = false;

            }
            Console.WriteLine('│');

            Console.Write("├");
            first = true;
            for (int i = 0; i < cell_strings.Count; ++i)
            {
                var length = cell_strings[i].Length;

                if (length == 0) continue;

                if (!first) Console.Write('┼');
                Console.Write(new String('─', cell_lengths[i]));
                first = false;

            }
            Console.WriteLine('┤');


            Console.Write("│");
            first = true;

            for (int i = 0; i < cell_strings.Count; ++i)
            {
                var length = cell_strings[i].Length;
                if (length == 0) continue;

                if (!first) Console.Write("│");
                if (length < cell_lengths[i])
                {
                    Console.Write(cell_strings[i].ToString());
                    Console.Write(new String(' ', cell_lengths[i] - length));
                }
                else
                {
                    Console.Write(cell_strings[i].ToString());
                }    
                first= false;
            }
            Console.WriteLine("│");

            Console.Write("└");
            first = true;
            for (int i = 0; i < cell_strings.Count; ++i)
            {
                var length = cell_strings[i].Length;

                if (length == 0) continue;

                if (!first) Console.Write('┴');
                Console.Write(new String('─', cell_lengths[i]));
                first = false;

            }
            Console.WriteLine('┘');

        }
    }




    class Discipline_Row
    {
        public enum Fields
        {
            ID, 
            NAME, 
            SEM, 
            STREAM, 
            GROUPS, 
            SUBGROUPS, 
            STUDENTS, 
            LEK,
            PR,
            LAB,
            KUR,
            IND,
            SUM,

            _COUNT
        }

        public static readonly Dictionary<Fields, int> fields_map = new Dictionary<Fields, int>()
        {
            { Fields.ID, 0 },
            { Fields.NAME, 4 },
            { Fields.SEM, 9 },
            { Fields.STREAM, 12 },
            { Fields.GROUPS, 15 },
            { Fields.SUBGROUPS, 19 },
            { Fields.STUDENTS, 22 },
            { Fields.LEK, 25 },
            { Fields.PR, 32 },
            { Fields.LAB, 39 },
            { Fields.KUR, 50 },
            { Fields.IND, 50 },
            { Fields.SUM, 54},
        };

        string [] data = new string[(int)Fields._COUNT];

        string this[int index]
        {
            get { return data[index]; }
            set { data[index] = value; }
        }

        public string this[Fields index] 
        {
            get { return data[(int)index]; } 
            set { data[(int)index] = value; }
        }

        public  int? id =>  Int32.TryParse(data[(int)Fields.ID], out var id) ? id: null;
        public  string name => data[(int)Fields.NAME];
        public  int? sem => Int32.TryParse(data[(int)Fields.SEM], out var s) ? s : null;
        public  string stream => data[(int)Fields.STREAM];
        public  int? group => Int32.TryParse(data[(int)Fields.GROUPS], out var g) ? g : null;
        public  int? sub_group => Int32.TryParse(data[(int)Fields.SUBGROUPS], out var sg) ? sg : null;
        public  int? student => Int32.TryParse(data[(int)Fields.STUDENTS], out var st) ? st : null;
        public  double? lections => Double.TryParse(data[(int)Fields.LEK], out var l) ? l : null;
        public  double? pr => Double.TryParse(data[(int)Fields.PR], out var p) ? p : null;
        public  double? labs => Double.TryParse(data[(int)Fields.LAB], out var lab) ? lab : null;
        public  double? ind => Double.TryParse(data[(int)Fields.IND], out var indv) ? indv : null;
        public double? kur => Double.TryParse(data[(int)Fields.KUR], out var k) ? k : null;
        public double? sum => Double.TryParse(data[(int)Fields.SUM], out var su) ? su : null;

        public override string ToString()
        {
           //return  new string ($"{id} {name} {sem} {stream} {group} {sub_group} {student} {bud_student} {part} {lections} {pr} {labs} {ind} {kur} {sum}");
           return data.Aggregate((a, b) => a + " " + b);
        }


        // навайбкодил
        public void PrintFormatted()
        {
            // Вывод одной строки через общий метод, который выравнивает столбцы по всем строкам
            PrintTable(new[] { this });
        }

        public static void PrintTable(IEnumerable<Discipline_Row> rowsEnumerable)
        {
            var rows = rowsEnumerable?.ToArray() ?? Array.Empty<Discipline_Row>();
            if (rows.Length == 0) return;

            // Получаем список полей в порядке объявления, исключая служебный _COUNT
            var fields = Enum.GetValues(typeof(Fields)).Cast<Fields>().Where(f => f != Fields._COUNT).ToArray();

            // Вычисляем ширину каждой колонки (максимум между именем поля и значением во всех строках)
            var widths = new int[fields.Length];
            for (int i = 0; i < fields.Length; ++i)
            {
                var name = fields[i].ToString();
                int maxValLen = 0;
                foreach (var r in rows)
                {
                    var v = r.data[(int)fields[i]] ?? string.Empty;
                    if (v.Length > maxValLen) maxValLen = v.Length;
                }
                widths[i] = Math.Max(name.Length, maxValLen);
            }

            // Верхняя граница
            Console.Write('┌');
            for (int i = 0; i < fields.Length; ++i)
            {
                if (i > 0) Console.Write('┬');
                Console.Write(new string('─', widths[i] + 2));
            }
            Console.WriteLine('┐');

            // Заголовок
            Console.Write('│');
            for (int i = 0; i < fields.Length; ++i)
            {
                var name = fields[i].ToString();
                Console.Write(' ' + name.PadRight(widths[i]) + ' ');
                Console.Write('│');
            }
            Console.WriteLine();

            // Разделитель между заголовком и значениями
            Console.Write('├');
            for (int i = 0; i < fields.Length; ++i)
            {
                if (i > 0) Console.Write('┼');
                Console.Write(new string('─', widths[i] + 2));
            }
            Console.WriteLine('┤');

            // Строки данных
            for (int rIndex = 0; rIndex < rows.Length; ++rIndex)
            {
                var r = rows[rIndex];
                Console.Write('│');
                for (int i = 0; i < fields.Length; ++i)
                {
                    var val = r.data[(int)fields[i]] ?? string.Empty;
                    Console.Write(' ' + val.PadRight(widths[i]) + ' ');
                    Console.Write('│');
                }
                Console.WriteLine();

                // Если это не последняя строка, вывести разделитель между строками
                if (rIndex + 1 < rows.Length)
                {
                    Console.Write('├');
                    for (int i = 0; i < fields.Length; ++i)
                    {
                        if (i > 0) Console.Write('┼');
                        Console.Write(new string('─', widths[i] + 2));
                    }
                    Console.WriteLine('┤');
                }
            }

            // Нижняя граница
            Console.Write('└');
            for (int i = 0; i < fields.Length; ++i)
            {
                if (i > 0) Console.Write('┴');
                Console.Write(new string('─', widths[i] + 2));
            }
            Console.WriteLine('┘');
        }

     
    }

    class VARow
    {
        public enum Fields
        {
            ID,
            NAME,
            SEM,
            STREAM,
            GROUPS,
            STUDENTS,
            HOURS_FORMULA,
            HOURS,

            _COUNT
        }
        public static readonly Dictionary<Fields, int> fields_map =  new Dictionary<Fields, int>()
        { 
            { Fields.ID, 0 },
            { Fields.NAME, 4 },
            { Fields.SEM, 7 },
            { Fields.STREAM, 11 },
            { Fields.GROUPS, 13 },
            { Fields.STUDENTS, 18},
            { Fields.HOURS_FORMULA, 22 },
            { Fields.HOURS, 30 },           
        };

        string[] data = new string[(int)Fields._COUNT];

        public string this[int index] 
        { 
            get { return data[index]; }
            set { data[index] = value; }
        }
        public string this[Fields index]
        {
            get { return data[(int)index]; }
            set { data[(int)index] = value; }
        }

        public override string ToString()
        {
            return data.Aggregate((a, b) => a + " " + b);
        }   

        public void PrintFormatted()
        {
            PrintTable(new[] { this });
        }

        public static void PrintTable(IEnumerable<VARow> rowsEnumerable)
        {
            var rows = rowsEnumerable?.ToArray() ?? Array.Empty<VARow>();
            if (rows.Length == 0) return;

            var fields = Enum.GetValues(typeof(Fields)).Cast<Fields>().Where(f => f != Fields._COUNT).ToArray();

            var widths = new int[fields.Length];
            for (int i = 0; i < fields.Length; ++i)
            {
                var name = fields[i].ToString();
                int maxValLen = 0;
                foreach (var r in rows)
                {
                    var v = r.data[(int)fields[i]] ?? string.Empty;
                    if (v.Length > maxValLen) maxValLen = v.Length;
                }
                widths[i] = Math.Max(name.Length, maxValLen);
            }

            // Верхняя граница
            Console.Write('┌');
            for (int i = 0; i < fields.Length; ++i)
            {
                if (i > 0) Console.Write('┬');
                Console.Write(new string('─', widths[i] + 2));
            }
            Console.WriteLine('┐');

            // Заголовок
            Console.Write('│');
            for (int i = 0; i < fields.Length; ++i)
            {
                var name = fields[i].ToString();
                Console.Write(' ' + name.PadRight(widths[i]) + ' ');
                Console.Write('│');
            }
            Console.WriteLine();

            // Разделитель
            Console.Write('├');
            for (int i = 0; i < fields.Length; ++i)
            {
                if (i > 0) Console.Write('┼');
                Console.Write(new string('─', widths[i] + 2));
            }
            Console.WriteLine('┤');

            // Строки данных
            for (int rIndex = 0; rIndex < rows.Length; ++rIndex)
            {
                var r = rows[rIndex];
                Console.Write('│');
                for (int i = 0; i < fields.Length; ++i)
                {
                    var val = r.data[(int)fields[i]] ?? string.Empty;
                    Console.Write(' ' + val.PadRight(widths[i]) + ' ');
                    Console.Write('│');
                }
                Console.WriteLine();

                if (rIndex + 1 < rows.Length)
                {
                    Console.Write('├');
                    for (int i = 0; i < fields.Length; ++i)
                    {
                        if (i > 0) Console.Write('┼');
                        Console.Write(new string('─', widths[i] + 2));
                    }
                    Console.WriteLine('┤');
                }
            }

            // Нижняя граница
            Console.Write('└');
            for (int i = 0; i < fields.Length; ++i)
            {
                if (i > 0) Console.Write('┴');
                Console.Write(new string('─', widths[i] + 2));
            }
            Console.WriteLine('┘');
        }

    }

    static class ExcellMapper
    {
        public enum A_Fields
        {
            ID = 1,
            NAME = 2,
            SEM = 3,
            STREAM = 4,
            STREAM_FULL = 5,
            GROUPS = 6,
            SUBGROUPS = 7,
            
            LEC_COUNT = 8,
            LEC_HOURS = 9,
            LEC_HOURS_BUDGET = 10,

            PR_COUNT = 11,
            PR_HOURS = 12,
            PR_BUDGET = 13,
            PR_IND = 14,
            

            LAB_COUNT = 15,
            LAB_HOURS = 16,
            LAB_BUDGET = 17,
            LAB_IND = 18,

            KUR_PER_MAN = 19,
            KUR = 20,

            IND_PER_MAN = 21,
            IND = 22,
            

            SUM = 23


        };

        private static Dictionary<A_Fields, string> columnnames = new Dictionary<A_Fields, string> 
        {
            {A_Fields.ID, "id"},
            {A_Fields.NAME, "наим"},
            {A_Fields.SEM, "сем" },
            {A_Fields.STREAM, "пот"  },
            {A_Fields.STREAM_FULL, "пот_год"  },
            {A_Fields.GROUPS, "гр"  },
            {A_Fields.SUBGROUPS, "подгр" },
            
            {A_Fields.LEC_COUNT, "лек кол" },
            {A_Fields.LEC_HOURS_BUDGET, "лек бюдж" },
            {A_Fields.LEC_HOURS, "лек час" },

            {A_Fields.PR_COUNT, "сем кол" },
            {A_Fields.PR_BUDGET, "сем бюдж" },
            {A_Fields.PR_HOURS, "сем час" },
            {A_Fields.PR_IND, "сем инд" },

            {A_Fields.LAB_COUNT, "лаб кол" },
            {A_Fields.LAB_HOURS, "лаб час" },
            {A_Fields.LAB_BUDGET, "лаб бюдж" },
            {A_Fields.LAB_IND, "лаб инд" },

            {A_Fields.KUR, "кур" },
            {A_Fields.KUR_PER_MAN, "кур/чел" },

            {A_Fields.IND, "инд" },
            {A_Fields.IND_PER_MAN, "инд/чел" },

            {A_Fields.SUM, "всего" }
        };
        
        public static string ColumnName(A_Fields col)
        {
            return columnnames[col];
        }
      
    }
}

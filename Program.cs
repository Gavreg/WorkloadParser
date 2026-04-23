// See https://aka.ms/new-console-template for more information
using System.Globalization;
using System.Reflection.Metadata;
using ConsoleApp11;
using RtfDomParser;
using OfficeOpenXml;


string inputfile="";

for (int i = 0; i<args.Length; i++)
{
    switch (args[i])
    {
        case "-i":
            {
                inputfile = args[++i];
                 

                break;
            }
    }
}

FileInfo fi = new FileInfo(inputfile);

ExcelPackage.LicenseContext = LicenseContext.NonCommercial;


const int curYear = 25;

RTFDomDocument doc = new RTFDomDocument();
var nagr = File.ReadAllText(inputfile);
doc.LoadRTFText(nagr);


var rows = new List<RTFDomTableRow>();



var discipline_rows = new List<Discipline_Row>();
var va_rows = new List<VARow>();


for (int i = 0; i < doc.Elements.Count; ++i)
{
    var rootable = doc.Elements[i];
    for (int t = 0; t < rootable.Elements.Count; ++t)
    {
        var el = rootable.Elements[t];
        var row = el as RTFDomTableRow;
        functions.PrintRow(row);
    }
}

Console.WriteLine(new String('=', 10));

for (int i = 0; i < doc.Elements.Count; ++i)
{
    var rootable = doc.Elements[i];
    for (int t = 0; t < rootable.Elements.Count; ++t)
    {


        var el = rootable.Elements[t];
        var row = el as RTFDomTableRow;

        string row_string = functions.GetRowText(row);

        

        //ищем строку с аудиторной нагрузкой
        if (row.Elements.Count >= 54 &&
                Int32.TryParse(functions.GetCellText(row.Elements[Discipline_Row.fields_map[Discipline_Row.Fields.ID]] as RTFDomTableCell), out _) &&
                !string.IsNullOrEmpty(functions.GetCellText(row.Elements[Discipline_Row.fields_map[Discipline_Row.Fields.NAME]] as RTFDomTableCell)) &&
                Double.TryParse(functions.GetCellText(row.Elements[Discipline_Row.fields_map[Discipline_Row.Fields.SUM]] as RTFDomTableCell), out _)
            )
        {
            
            var d_row = new Discipline_Row();
            discipline_rows.Add(d_row);

            foreach (var x in Discipline_Row.fields_map)
            {
                string data = functions.GetCellText(row.Elements[x.Value] as RTFDomTableCell);
                if (data != "")
                    d_row[x.Key] += " " + data;
            }

            d_row.PrintFormatted();

        }
        //ищем строку с внеаудиторной нагрузкой
        if (row.Elements.Count >= 30 &&
                Int32.TryParse(functions.GetCellText(row.Elements[VARow.fields_map[VARow.Fields.ID]] as RTFDomTableCell), out _) &&
                !string.IsNullOrEmpty(functions.GetCellText(row.Elements[VARow.fields_map[VARow.Fields.NAME]] as RTFDomTableCell)) &&
                Double.TryParse(functions.GetCellText(row.Elements[VARow.fields_map[VARow.Fields.HOURS]] as RTFDomTableCell), out _)
            )

        {
            

            var varow = new VARow();
            va_rows.Add(varow);

            foreach (var x in VARow.fields_map)
            {
                string data = functions.GetCellText(row.Elements[x.Value] as RTFDomTableCell);
                if (data != "")
                    varow[x.Key] += " " + data;
            }

            varow.PrintFormatted();
        }

    }
}

Console.WriteLine(new String('=', 10));


using StreamWriter sw = new StreamWriter($"{Path.GetFileNameWithoutExtension(fi.Name)}.txt",false);
using MemoryStream ms = new MemoryStream(sw);



sw.WriteLine($"Всего аудиторной:");
double sum_os = Math.Round(discipline_rows.Where(x =>  x.sem.Value % 2 != 0).Sum(x => x.sum.Value),1);
double sum_ves = Math.Round(discipline_rows.Where(x => x.sem.Value % 2 == 0).Sum(x => x.sum.Value),1);
double sum = Math.Round(sum_os + sum_ves, 1);
sw.WriteLine($"Осень: {sum_os}");
sw.WriteLine($"Весна: {sum_ves}");
sw.WriteLine($"Всего: {sum}");

sw.WriteLine($"");

var excel = new ExcelPackage();
var ws = excel.Workbook.Worksheets.Add("Аудиторная");


for (int i=1; i<30; ++i)
{
    ws.Cells[1, i].Value = i;
}

var ___i = 1;
ws.Cells[2, ___i++].Value = "id";
ws.Cells[2, ___i++].Value = "наим";
ws.Cells[2, ___i++].Value = "сем";
ws.Cells[2, ___i++].Value = "пот";
ws.Cells[2, ___i++].Value = "пот2";
ws.Cells[2, ___i++].Value = "гр";
ws.Cells[2, ___i++].Value = "подгр";

ws.Cells[2, ___i++].Value = "лек кол";
ws.Cells[2, ___i++].Value = "лек час";
ws.Cells[2, ___i++].Value = "лек бю";

ws.Cells[2, ___i++].Value = "сем кол";
ws.Cells[2, ___i++].Value = "сем час";
ws.Cells[2, ___i++].Value = "сем студ";
ws.Cells[2, ___i++].Value = "сем инд";

ws.Cells[2, ___i++].Value = "лаб кол";
ws.Cells[2, ___i++].Value = "лаб час";
ws.Cells[2, ___i++].Value = "лаб студ";
ws.Cells[2, ___i++].Value = "лаб инд";

ws.Cells[2, ___i++].Value = "кур/чел";
ws.Cells[2, ___i++].Value = "кур час";

ws.Cells[2, ___i++].Value = "инд/чел";
ws.Cells[2, ___i++].Value = "инд час";

ws.Cells[2, ___i++].Value = "всего";




int _row = 3; 

foreach (var dr in discipline_rows)
{
    sw.WriteLine($"{dr.id} {dr.name}");
    sw.WriteLine($"   Семестр {dr.sem}");
    sw.WriteLine($"   Поток {dr.stream}");
    sw.WriteLine($"   Кол-во групп {dr.group} ({dr.sub_group} подогр.)");
    sw.WriteLine($"   Студентов {dr.student}");

    var __i = 1;
    int _sem = dr.sem.Value;
    ws.Cells[_row, __i++].Value = dr.id;
    ws.Cells[_row, __i++].Value = dr.name;
    ws.Cells[_row, __i++].Value = dr.sem;
    ws.Cells[_row, __i++].Value = dr.stream;
    ws.Cells[_row, __i++].Value = dr.stream + "-" + Math.Round(1.0*curYear - ( (_sem - 1))%8 / 2).ToString();
    ws.Cells[_row, __i++].Value = dr.group;
    ws.Cells[_row, __i++].Value = dr.student;

   

    if (dr.lections.HasValue)
    {
        double lek = dr.lections.Value ;
        
        sw.WriteLine($"   ЛЕКЦИИ -> Кол-во      {lek / 2}    ");
        sw.WriteLine($"             Часов-всего {lek}");
        sw.WriteLine($"             Индив.      0,3");

        ws.Cells[_row, __i++].Value = lek / 2;
        ws.Cells[_row, __i++].Value = lek;
        ws.Cells[_row, __i++].Value = lek;
    }
    else    
        __i+=3;


    if (dr.pr.HasValue)
    {
        double pr = dr.pr.Value / dr.group.Value;
        double pr_all = dr.pr.Value;
        double ind = 0;
        if (dr.sem < 5)
        {
            ind = 0.5 * pr / 8.0;
        }
        else
        {
            ind = 0.2 * pr / 4.0;
        }
        sw.WriteLine($"   СЕМИНАРЫ -> Кол-во      {pr / 2}   ");
        sw.WriteLine($"               Часов-всего {pr_all}");
        sw.WriteLine($"               Индив.      {Math.Round(ind, 1, MidpointRounding.ToZero)}");

        ws.Cells[_row, __i++].Value = pr / 2;
        ws.Cells[_row, __i++].Value = pr_all;
        ws.Cells[_row, __i++].Value = pr_all;
        ws.Cells[_row, __i++].Value = Math.Round(ind, 1, MidpointRounding.ToZero);

    }
    else
        __i += 4;

    if (dr.labs.HasValue)
    { 
        double labs = Math.Round(dr.labs.Value / dr.sub_group.Value, 1);

        double labs_all = dr.labs.Value;
        double ind = 0;
        if (dr.sem < 5)
        {
            ind = 0.3 * labs / 4.0;
        }
        else
        {
            ind = 0.2 * labs / 4.0;
        }
        sw.WriteLine($"   ЛАБОРАТОРНЫЕ -> Кол-во      {labs / 4} ");
        sw.WriteLine($"                   Часов-всего {labs_all}");
        sw.WriteLine($"                   Индив.      {Math.Round(ind,1, MidpointRounding.ToZero)}");

        ws.Cells[_row, __i++].Value = labs / 4;
        ws.Cells[_row, __i++].Value = labs_all;
        ws.Cells[_row, __i++].Value = labs_all;
        ws.Cells[_row, __i++].Value = Math.Round(ind, 1, MidpointRounding.ToZero);
    }
    else
        __i += 4;

    if (dr.kur.HasValue)
    {
        double kur = dr.kur.Value;
        double hours_per_man = kur / dr.student.Value;
        sw.WriteLine($"   КУРСОВЫЕ ->     На человека: {hours_per_man} ");
        sw.WriteLine($"                   Часов-всего  {kur}");

        ws.Cells[_row, __i++].Value = hours_per_man;
        ws.Cells[_row, __i++].Value = kur;
    }
    else
        __i += 2;

    if (dr.ind > 0)
    {
        double ind = dr.ind.Value;
        double ind_per_man = ind / dr.student.Value;
        sw.WriteLine($"   ИНДИВИДУАЛЬНАЯ -> На человека: {ind_per_man}");
        sw.WriteLine($"                     Часов-всего  {ind}");

        ws.Cells[_row, __i++].Value = ind_per_man;
        ws.Cells[_row, __i++].Value = ind;
    }
    else
        __i += 2;
    sw.WriteLine($"   ВСЕГО ЗА КУРС ->  {dr.sum}");
    ws.Cells[_row, __i++].Value = dr.sum;

    sw.WriteLine("--------------");
    sw.WriteLine("");

    _row++;

}

ws = excel.Workbook.Worksheets.Add("Внеаудиторная");

_row = 3;

va_rows = va_rows.Where(x => (x[VARow.Fields.STREAM] != null) && !x[VARow.Fields.STREAM].Contains("АСП")).ToList();

___i = 1;

ws.Cells[2, ___i++].Value = "id";
ws.Cells[2, ___i++].Value = "имя";
ws.Cells[2, ___i++].Value = "сем";
ws.Cells[2, ___i++].Value = "пот";
ws.Cells[2, ___i++].Value = "пот2";
ws.Cells[2, ___i++].Value = "груп";
ws.Cells[2, ___i++].Value = "студ";
ws.Cells[2, ___i++].Value = "ф-ла";
ws.Cells[2, ___i++].Value = "часы";

foreach (var _varow in va_rows)
{
    int __i = 1;

    ws.Cells[_row, __i++].Value = _varow[VARow.Fields.ID];
    ws.Cells[_row, __i++].Value = _varow[VARow.Fields.NAME];
    ws.Cells[_row, __i++].Value = _varow[VARow.Fields.SEM];
    ws.Cells[_row, __i++].Value = _varow[VARow.Fields.STREAM];
    //  --->>>   dr.stream + "-" + Math.Round(1.0*curYear - ( (_sem - 1))%8 / 2).ToString();
    int _sem = 0;
    if ( Int32.TryParse(_varow[VARow.Fields.SEM],out _sem) )
    {
        ws.Cells[_row, __i++].Value = _varow[VARow.Fields.STREAM] + "-" + Math.Round(1.0 * curYear - ((_sem - 1)) % 8 / 2).ToString();
    }
    ws.Cells[_row, __i++].Value = _varow[VARow.Fields.GROUPS];
    ws.Cells[_row, __i++].Value = _varow[VARow.Fields.STUDENTS];
    ws.Cells[_row, __i++].Value = _varow[VARow.Fields.HOURS_FORMULA];
    ws.Cells[_row, __i++].Value = _varow[VARow.Fields.HOURS];

    _row++;

}

sw.Close();





for(bool f = true; f;)
{
    try
    {
        excel.SaveAs($"{Path.GetFileNameWithoutExtension(fi.Name)}.xlsx");
        f = false;
    }

    catch (Exception e)
    {        

        Console.WriteLine(e.ToString());
        Console.WriteLine("");
        Console.WriteLine("нажмите клавишу для повтора...");
        Console.ReadKey();
    }
}


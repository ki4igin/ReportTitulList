using AutoCompleteConsole;
using AutoCompleteConsole.StringProvider;
using ReportTitulList;

AppSettings setting = AppSettings.Read();
Request request = Acc.CreateRequest(new());
Selector selector = Acc.CreateSelector(new());

string fileName = request.ReadLine(new("Введите имя файла"), setting.FileName) + ".docx";
using Report report = new(fileName);

start:
string year = selector.Run(new("Выберите год", setting.Years));
int semester =
    int.Parse(selector.Run(new("Выберите семестр", setting.NumberOfSemesters.Select(i => $"{i}").ToArray())));
string group = selector.Run(new("Выберите группу", setting.Groups));
string discipline = selector.Run(new("Выберите дисциплину", setting.Disciplines));
string workType = selector.Run(new("Выберите тип отчетов", setting.WorkTypes));
int workCount = int.Parse(request.ReadLine(new("Введите количество работ", "Количество должно быть целым"),
    s => int.TryParse(s, out int _), "8"));

int cnt = 0;
List<string> names = new();
while (cnt < 2)
{
    string str = request.ReadLine(new("Введите ФИО", "", "", "Для выхода нажмите два раза Enter"), _ => true, "");
    if (str.Length > 3)
    {
        names.Add(str);
        cnt = 0;
    }
    else
    {
        cnt++;
    }
}

ReportSettings reportSettings = new(year, semester, group, discipline, workType, workCount, names.ToArray());
report.Add(reportSettings);

Acc.WriteLine(
    "\nСгенерирован отчет\n" +
    $"{year}, {semester} семестр\n" +
    $"{group}\n" +
    $"{discipline}\n" +
    $"{workType}");
Acc.WriteLine("\nДля продолжения нажмите два раза Enter, для выхода любую другую кнопку");

if (Console.ReadKey().Key == ConsoleKey.Enter)
    if (Console.ReadKey().Key == ConsoleKey.Enter)
        goto start;

AppSettings.Save(setting);
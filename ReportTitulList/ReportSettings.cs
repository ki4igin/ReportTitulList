namespace ReportTitulList;
public record ReportSettings(
    string Year,
    int Semester,
    string Group,
    string Discipline,
    string WorkType,
    int WorkCount,
    string[] Names
);
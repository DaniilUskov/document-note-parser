using DocToExcelNoteParser.Workers;

class Program
{
    static void Main()
    {
        Console.WriteLine("Чтение docx файла...");

        var footNotesCollector = new FootNotesCollector();

        Console.WriteLine("Считывание сносок...");

        var footNotes = footNotesCollector.GetFootNotes();

        var excelCreator = new ExcelCreator();

        Console.WriteLine("Создание Excel файла");

        excelCreator.GenerateExcel(footNotes);

        Console.WriteLine("Excel файл создан");
    }
}
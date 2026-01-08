using NPOI.XWPF.UserModel;

public class Program
{
    public static void Main()
    {
        // Specifică calea unde vrei să salvezi fișierul generat
        const string docxFilePath = "portofoliu-de-test.docx";
        const string csvFilePath = "input.csv";

        creazaDocumentWord(docxFilePath, null);

        Console.WriteLine("Documentul a fost creat cu succes la " + docxFilePath);

        citesteColectieCSV(csvFilePath);
    }

    public static XWPFDocument creazaDocumentWord(string filePath, Portofoliu portofoliu)
    {
        // Creează un document nou
        var document = new XWPFDocument();

        // Adauga titlul
        scrieTitlu(document);

        // Salveaza fisierul
        using(var fileSave = new FileStream(filePath, FileMode.Create))
        {
            document.Write(fileSave);
        }

        return document;
    }

    // Metoda din cadrul functiei XWPFDocument creazaDocumentWord(string filePath, Portofoliu portofoliu)
    // care are rolul de a scrie paragraful cu titlul documentului mereu asa cum se regaseste in exemplu 
    public static void scrieTitlu(XWPFDocument document)
    {
        var title = document.CreateParagraph();
        title.Alignment = ParagraphAlignment.CENTER;

        var preTitleRun = title.CreateRun();
        preTitleRun.SetText("ANEXĂ la Convenţia-cadru privind efectuarea stagiului de practică în cadrul programelor de studii universitare de licenţă/masterat");
        preTitleRun.AddCarriageReturn();
        preTitleRun.AddCarriageReturn();

        var titleRun = title.CreateRun();
        titleRun.SetText("PORTOFOLIU DE PRACTICĂ");
        titleRun.FontSize = 18;
        titleRun.IsBold = true;
        titleRun.AddCarriageReturn();
        
        var postTitleRun = title.CreateRun();
        postTitleRun.SetText("aferent Convenţiei-cadru privind efectuarea stagiului de practică în cadrul programelor de studii universitare de licenţă și masterat");
    }

    public static List<Portofoliu> citesteColectieCSV(string csvFilePath)
    {
        List<Portofoliu> colectiePortofolii = new List<Portofoliu>();

        StreamReader reader = new StreamReader(csvFilePath);

        string? line;

        while((line = reader.ReadLine()) != null)
        {
            string[] fields = line.Split('|');

            // Creaza stagiul din datele citite
            int idStagiu = Convert.ToInt32(fields[0]);
            int durata = Convert.ToInt32(fields[1]);
            string calendar = fields[2];
            DateOnly dataStart = DateOnly.Parse(fields[3]);
            DateOnly dataFinal = DateOnly.Parse(fields[4]);
            int orePeZi = Convert.ToInt32(fields[5]);
            TimeOnly oraStart = TimeOnly.Parse(fields[6]);
            TimeOnly oraFinal = TimeOnly.Parse(fields[7]);
            string adresa = fields[8];

            PerioadaStagiu perioadaStagiu = new PerioadaStagiu(dataStart, dataFinal, orePeZi, oraStart, oraFinal);

            // StagiuDePractica stagiuDePractica = new StagiuDePractica(idStagiu, durata, calendar, perioadaStagiu, adresa, );
                
            // for(int i=0;i<fields.Length;i++)
            // {
                // Console.WriteLine(fields[i]);
            // }
        }

        return colectiePortofolii;
    }
}
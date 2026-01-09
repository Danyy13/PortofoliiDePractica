using System.Globalization;
using NPOI.Util;
using NPOI.XWPF.Model;
using NPOI.XWPF.UserModel;

public enum TipRubrica
{
    Simplu,
    Lista,
    Tabel
}

public class Program
{
    public static void Main()
    {
        // Specifică calea unde vrei să salvezi fișierul generat
        const string docxFilePath = "portofoliu-de-test.docx";
        const string csvFilePath = "input.csv";

        List<Portofoliu> colectiePortofolii = citesteColectiePortofoliiDinCSV(csvFilePath); 

        creazaDocumentWord(docxFilePath, colectiePortofolii[0]);

        Console.WriteLine("Documentul a fost creat cu succes la " + docxFilePath);
    }

    // public static CT_AbstractNum creazaSablonNumerotareCifreArabe()
    // {
    //     CT_AbstractNum sablonCifreArabe = new CT_AbstractNum();
    //     sablonCifreArabe.abstractNumId = "0";
    //     CT_Lvl level = new CT_Lvl();
    //     level.ilvl = "0";

    //     level.numFmt = new CT_NumFmt();
    //     level.numFmt.val = ST_NumberFormat.decimalEnclosedFullstop;

    //     if(sablonCifreArabe.lvl == null)
    //     {
    //         sablonCifreArabe.lvl = new List<CT_Lvl>();
    //     }
    //     sablonCifreArabe.Add(level);

    //     return sablonCifreArabe;
    // }

    public static XWPFDocument creazaDocumentWord(string filePath, Portofoliu portofoliu)
    {
        // Creează un document nou
        var document = new XWPFDocument();

        // Adauga titlul
        scrieTitlu(document);

        // Creaza lista numerotata cu rubricile
        XWPFNumbering numbering = document.CreateNumbering();
        string abstractNumId = numbering.AddAbstractNum();
        string numId = numbering.AddNum("1.");

        string titluRubrica, descriereRubrica;

        // 1
        titluRubrica = "Durata totală a pregătirii practice: ";
        descriereRubrica = portofoliu.stagiuDePractica.durata.ToString() + "h";
        scrieRubrica(document, titluRubrica, descriereRubrica, numId, numbering, TipRubrica.Simplu);

        // 2
        titluRubrica = "Calendarul pregătirii: ";
        descriereRubrica = portofoliu.stagiuDePractica.calendar;
        scrieRubrica(document, titluRubrica, descriereRubrica, numId, numbering, TipRubrica.Lista);

        // 3
        titluRubrica = "Perioada stagiului, timpul de lucru şi orarul (de precizat zilele de pregătire practică în cazul timpului de lucru parţial): ";
        descriereRubrica = portofoliu.stagiuDePractica.perioadaStagiu.ToString();
        scrieRubrica(document, titluRubrica, descriereRubrica, numId, numbering, TipRubrica.Simplu);

        // 4
        titluRubrica = "Adresa unde se va derula stagiul de pregătire practică: ";
        descriereRubrica = portofoliu.stagiuDePractica.adresa;
        scrieRubrica(document, titluRubrica, descriereRubrica, numId, numbering, TipRubrica.Simplu);

        // 5
        titluRubrica = "Deplasarea în afara locului unde este repartizat practicantul vizează următoarele locaţii: ";
        descriereRubrica = portofoliu.stagiuDePractica.locatiiDeplasare;
        scrieRubrica(document, titluRubrica, descriereRubrica, numId, numbering, TipRubrica.Simplu);

        // 6
        titluRubrica = "Condiţii de primire a studentului/masterandului în stagiul de practică: ";
        descriereRubrica = portofoliu.stagiuDePractica.listaConditii;
        scrieRubrica(document, titluRubrica, descriereRubrica, numId, numbering, TipRubrica.Lista);

        // 7
        titluRubrica = "Modalităţi prin care se asigură complementaritatea între pregătirea dobândită de studentul practicant în instituţia de învăţământ superior şi în cadrul stagiului de practică: ";
        descriereRubrica = portofoliu.stagiuDePractica.modalitatiComplementaritate;
        scrieRubrica(document, titluRubrica, descriereRubrica, numId, numbering, TipRubrica.Lista);

        // 8
        titluRubrica = "Numele şi prenumele cadrului didactic care asigură supravegherea pedagogică a practicantului pe perioada stagiului de practică: ";
        descriereRubrica = portofoliu.profesorCoordonator.nume + " " + portofoliu.profesorCoordonator.prenume;
        scrieRubrica(document, titluRubrica, descriereRubrica, numId, numbering, TipRubrica.Simplu);
        
        // 9
        titluRubrica = "Drepturi şi responsabilităţi ale cadrului didactic din unitatea de învăţământ - organizator al practicii, pe perioada stagiului de practică: ";
        descriereRubrica = portofoliu.profesorCoordonator.responsabilitati;
        scrieRubrica(document, titluRubrica, descriereRubrica, numId, numbering, TipRubrica.Lista);

        // 10
        titluRubrica = "Numele şi prenumele tutorelui desemnat de întreprindere care va asigura respectarea condiţiilor de pregătire şi dobândirea de către practicant a competenţelor profesionale planificate pentru perioada stagiului de practică: ";
        descriereRubrica = portofoliu.tutore.nume + " " + portofoliu.tutore.prenume;
        scrieRubrica(document, titluRubrica, descriereRubrica, numId, numbering, TipRubrica.Simplu);
        
        // 9
        titluRubrica = "Drepturi şi responsabilităţi ale tutorelui de practică desemnat de partenerul de practică: ";
        descriereRubrica = portofoliu.tutore.responsabilitati;
        scrieRubrica(document, titluRubrica, descriereRubrica, numId, numbering, TipRubrica.Lista);     

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

/*
    Functie: scrieRubrica
    Parametri: XPWFDocument document - documentul in care este scrisa rubrica
    Scop: Scrie o rubrica simpla de genul "{numar}. {titluRubrica}: {descriereRubrica}", unde titlul rubricii este scris ca in documentul
    dat exemplu, iar descrierile sunt datele citite din csv. Rubrica face parte dintr-o lista numerotata
*/
    public static void scrieRubrica(XWPFDocument document, string titluRubrica, string descriereRubrica, string numId, XWPFNumbering numbering, TipRubrica tipRubrica)
    {
        XWPFParagraph rubrica = document.CreateParagraph();
        rubrica.Alignment = ParagraphAlignment.LEFT;
        rubrica.SetNumID(numId, "0");

        XWPFRun runTitlu = rubrica.CreateRun();
        runTitlu.IsBold = true;
        runTitlu.SetText(titluRubrica);
        
        XWPFRun runDescriere = rubrica.CreateRun(); // descrierea se va scrie diferit in functie de rubrica
        switch(tipRubrica)
        {
            case TipRubrica.Simplu:
                runDescriere.SetText(descriereRubrica);
                break;
            case TipRubrica.Lista:
                // runDescriere.AddCarriageReturn();
                string[] lines = descriereRubrica.Split(';'); // separa liniile listei
                
                // creaza lista nenumerotata avand caractere '-'
                string numIdSubLista = numbering.AddNum("-");

                foreach(string line in lines)
                {
                    XWPFParagraph listItem = document.CreateParagraph();
                    listItem.SetNumID(numIdSubLista, "1");

                    XWPFRun listItemRun = listItem.CreateRun();
                    listItemRun.SetText(line);      
                }

                // runDescriere.SetText(descriereRubrica);
                break;

            default:
                break;
        }
    }

    public static List<Portofoliu> citesteColectiePortofoliiDinCSV(string csvFilePath)
    {
        List<Portofoliu> colectiePortofolii = new List<Portofoliu>();

        StreamReader reader = new StreamReader(csvFilePath);
        
        string headerLine = reader.ReadLine();
        var columns = headerLine.Split('|')
                                .Select((name, index) => new { name, index })
                                .ToDictionary(x => x.name.Trim(), x => x.index);
        string? line;

        while((line = reader.ReadLine()) != null)
        {
            string[] fields = line.Split('|');

            // Creaza stagiul din datele citite
            int idStagiu = Convert.ToInt32(fields[columns["idStagiu"]]);
            int durata = Convert.ToInt32(fields[columns["durata"]]);
            string calendar = fields[columns["calendar"]];
            DateOnly dataStart = DateOnly.ParseExact(fields[columns["dataStart"]], "dd/MM/yyyy", CultureInfo.InvariantCulture);
            DateOnly dataFinal = DateOnly.ParseExact(fields[columns["dataFinal"]], "dd/MM/yyyy", CultureInfo.InvariantCulture);
            int orePeZi = Convert.ToInt32(fields[columns["orePeZi"]]);
            TimeOnly oraStart = TimeOnly.Parse(fields[columns["oraStart"]]);
            TimeOnly oraFinal = TimeOnly.Parse(fields[columns["oraFinal"]]);
            string adresa = fields[columns["adresa"]];
            string deplasari = fields[columns["deplasari"]];
            string conditiiDePrimire = fields[columns["conditiiDePrimire"]];
            string modalitatiComplementaritate = fields[columns["modalitatiComplementaritate"]];            
            string modalitatiEvaluare = fields[columns["modalitatiEvaluare"]];
            int idProfesor = Convert.ToInt32(fields[columns["idProfesor"]]);
            string numeProfesor = fields[columns["numeProfesor"]];
            string prenumeProfesor = fields[columns["prenumeProfesor"]];
            string functieProfesor = fields[columns["functieProfesor"]];
            DateTime dataSemnaturiiProfesor = DateTime.ParseExact(fields[columns["dataSemnaturiiProfesor"]], "dd/MM/yyyy", CultureInfo.InvariantCulture);
            string responsabilitatiProfesor = fields[columns["responsabilitatiProfesor"]];
            int idTutore = Convert.ToInt32(fields[columns["idTutore"]]);
            string numeTutore = fields[columns["numeTutore"]];
            string prenumeTutore = fields[columns["prenumeTutore"]];
            string functieTutore = fields[columns["functieTutore"]];
            DateTime dataSemnaturiiTutore = DateTime.ParseExact(fields[columns["dataSemnaturiiTutore"]], "dd/MM/yyyy", CultureInfo.InvariantCulture);
            string responsabilitatiTutore = fields[columns["responsabilitatiTutore"]];
            int idStudent = Convert.ToInt32(fields[columns["idStudent"]]);
            string numeStudent = fields[columns["numeStudent"]];
            string prenumeStudent = fields[columns["prenumeStudent"]];
            string functieStudent = fields[columns["functieStudent"]];
            DateTime dataSemnaturiiStudent = DateTime.ParseExact(fields[columns["dataSemnaturiiStudent"]], "dd/MM/yyyy", CultureInfo.InvariantCulture);

            List<Competenta> listaCompetente = new List<Competenta>();
            string competente = fields[columns["competente"]];
            var listaCompetenteInitial = competente.Split(';').ToList();
            for(int i=0;i<listaCompetenteInitial.Count;i++)
            {
                string[] compFields = listaCompetenteInitial[i].Split('#');
                Competenta competenta = new Competenta(compFields[0], compFields[1], compFields[2], compFields[3]);
                listaCompetente.Add(competenta);
            }

            PerioadaStagiu perioadaStagiu = new PerioadaStagiu(dataStart, dataFinal, orePeZi, oraStart, oraFinal);
            StagiuDePractica stagiuDePractica = new StagiuDePractica(idStagiu, durata, calendar, perioadaStagiu, adresa, deplasari, conditiiDePrimire, modalitatiComplementaritate, listaCompetente, modalitatiEvaluare);

            ProfesorCoordonator profesorCoordonator = new ProfesorCoordonator(idProfesor, numeProfesor, prenumeProfesor, functieProfesor, dataSemnaturiiProfesor, responsabilitatiProfesor);
            Tutore tutore = new Tutore(idTutore, numeTutore, prenumeTutore, functieTutore, dataSemnaturiiTutore, responsabilitatiTutore);
            StudentPracticant studentPracticant = new StudentPracticant(idStudent, numeStudent, prenumeStudent, functieStudent, dataSemnaturiiStudent);

            Portofoliu portofoliu = new Portofoliu(stagiuDePractica, studentPracticant, tutore, profesorCoordonator);
            colectiePortofolii.Add(portofoliu);
        }

        return colectiePortofolii;
    }
}
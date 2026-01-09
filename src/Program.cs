using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using System.Drawing;
using System.Globalization;
using System.Reflection.Metadata;
using System.Security.Cryptography.X509Certificates;
using NPOI.HSSF.Record;
using NPOI.Util;
using NPOI.XWPF.Model;
using NPOI.XWPF.UserModel;

public enum TipRubrica
{
    Simplu,
    Lista,
    Default
}

public class Program
{
    public static void Main()
    {
        // Specifică calea unde vrei să salvezi fișierul generat
        // const string docxFilePath = "portofoliu-de-test.docx";
        const string csvFilePath = "../inputs/input.csv";

        List<Portofoliu> colectiePortofolii = citesteColectiePortofoliiDinCSV(csvFilePath); 
        
        foreach(Portofoliu portofoliu in colectiePortofolii)
        {
            string docxFilePath = "../portofolii/" + portofoliu.studentPracticant.nume + " " + portofoliu.studentPracticant.prenume + " - " + portofoliu.companie + " - portofoliu.docx";
            
            creazaDocumentWord(docxFilePath, portofoliu);

            Console.WriteLine("Documentul a fost creat cu succes la " + docxFilePath);
        }
    }

    public static XWPFDocument creazaDocumentWord(string filePath, Portofoliu portofoliu)
    {
        // Creează un document nou
        var document = new XWPFDocument();

        // insereaza logo UPT in header
        // XWPFParagraph logo = document.CreateParagraph();
        // XWPFRun logoRun = logo.CreateRun();
        // logo.Alignment = ParagraphAlignment.RIGHT;
        // string imagePath = "../assets/logoupt.jpg";

        // if(File.Exists(imagePath))
        // {
        //     using(FileStream imageData = new FileStream(imagePath, FileMode.Open, FileAccess.Read))
        //     {
        //         int width = Units.ToEMU(100);
        //         int height = Units.ToEMU(100);

        //         logoRun.AddPicture(imageData, (int)PictureType.JPEG, "logoupt.jpg", width, height);
        //     }
        // } else
        // {
        //     Console.WriteLine("Fisierul nu exista la calea specificata");
        // }

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
        
        // 11
        titluRubrica = "Drepturi şi responsabilităţi ale tutorelui de practică desemnat de partenerul de practică: ";
        descriereRubrica = portofoliu.tutore.responsabilitati;
        scrieRubrica(document, titluRubrica, descriereRubrica, numId, numbering, TipRubrica.Lista);
        
        // 12 - nu exista un raspuns in exemplu

        // 13 - scrie doar titlul rubricii si apoi creaza tabelul
        titluRubrica = "Definirea competențelor care vor fi dobândite pe perioada stagiului de practică: ";
        descriereRubrica = "";
        scrieRubrica(document, titluRubrica, descriereRubrica, numId, numbering, TipRubrica.Default);
        // creaza tabelul de competente
        scrieTabelCompetente(document, portofoliu);

        // 14
        titluRubrica = "Modalităţi de evaluare a pregătirii profesionale dobândite de practicant pe perioada stagiului de pregătire practică: ";
        descriereRubrica = portofoliu.stagiuDePractica.modalitatiEvaluare;
        scrieRubrica(document, titluRubrica, descriereRubrica, numId, numbering, TipRubrica.Simplu);

        // Paragraf de final
        XWPFParagraph paragrafFinal = document.CreateParagraph();
        XWPFRun runFinal = paragrafFinal.CreateRun();
        runFinal.SetText("Evaluarea practicantului pe perioada stagiului de pregătire practică se va face de către tutore.");

        // creaza ultimul tabel - tabel de semnaturi
        scrieTabelSemnaturi(document, portofoliu.profesorCoordonator, portofoliu.tutore, portofoliu.studentPracticant);

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
                string[] lines = descriereRubrica.Split(';'); // separa liniile listei
                
                // creaza lista nenumerotata avand caractere '-'
                string numIdSubLista = numbering.AddNum("-");

                // scrie elementele listei
                foreach(string line in lines)
                {
                    XWPFParagraph listItem = document.CreateParagraph();
                    listItem.SetNumID(numIdSubLista, "1");

                    XWPFRun listItemRun = listItem.CreateRun();
                    listItemRun.SetText(line);      
                }
                break;

            default:
                break;
        }
    }

    public static void scrieTabelCompetente(XWPFDocument document, Portofoliu portofoliu)
    {   
        // date pentru randuri si coloane
        int rows = portofoliu.stagiuDePractica.listaCompetente.Count + 1; // numarul de randuri este egal cu numarul de competente + randul de header
        int cols = 6; // numar de atribute pentru fiecare compententa

        // date pentru latimea celulelor
        const int totalWidth = 5000;
        // int[] colWidths = { 600, 3000, 1500, 1500, 1500, 900 };

        XWPFTable table = document.CreateTable(rows, cols);
        table.Width = totalWidth;

        // text cap tabel
        string[] tableHeaderStrings =
        {
            "Nr. Crt.", "Competenta", "Modul de pregatire", "Locul de Munca", "Activitati Planificate", "Observatii"
        };

        // scrie cap tabel
        var headerRow = table.GetRow(0);
        for(int i=0;i<cols;i++)
        {
            // table.SetColumnWidth(i, Convert.ToUInt64(colWidths[i]));
            var headerCell = headerRow.GetCell(i);

            // seteazaLatimeCelula(headerCell, colWidths[i]);
            headerCell.SetText(tableHeaderStrings[i]);
        }

        // completare date
        for(int i=0;i<portofoliu.stagiuDePractica.listaCompetente.Count;i++) // incepem de la 1 pentru ca ignoram randul header
        {
            Competenta compententa = portofoliu.stagiuDePractica.listaCompetente[i];
            var row = table.GetRow(i + 1);

            row.GetCell(0).SetText((i + 1).ToString());
            row.GetCell(1).SetText(compententa.descriere);
            row.GetCell(2).SetText(compententa.modPregatire);
            row.GetCell(3).SetText(compententa.locDeMunca);
            row.GetCell(4).SetText(compententa.activitati);
            row.GetCell(5).SetText(compententa.observatii);
        }
    }

    public static void scrieTabelSemnaturi(XWPFDocument document, ProfesorCoordonator profesorCoordonator, Tutore tutore, StudentPracticant studentPracticant)
    {
        // date pentru randuri si coloane
        string[] rowsText = { "", "Nume si prenume", "Functia", "Data", "Semnatura" };
        string[] columnsText = { "", "Cadru didactic supervizor", "Tutore", "Practicant" }; // prima celula este goala

        int rows = rowsText.Length;     
        int cols = columnsText.Length;        

        // date pentru latimea celulelor
        const int totalWidth = 5000;

        XWPFTable table = document.CreateTable(rows, cols);
        table.Width = totalWidth;

        // scrie cap tabel
        var headerRow = table.GetRow(0);
        for(int i=0;i<cols;i++)
        {
            XWPFTableCell headerCell = headerRow.GetCell(i); 
            seteazaLatimeCelula(headerCell, totalWidth / cols);

            var paragraph = headerCell.AddParagraph();
            var run = paragraph.CreateRun();
            run.IsBold = true;
            run.SetText(columnsText[i]);

            // headerCell.SetText(columnsText[i]);
        }

        Persoana[] persons = { profesorCoordonator, tutore, studentPracticant };

        // completare date
        for(int i=1;i<rows;i++)
        {
            var row = table.GetRow(i);

            row.GetCell(0).SetText(rowsText[i]);
        }
    }

    public static void seteazaLatimeCelula(XWPFTableCell cell, int width)
    {
        var ctTc = cell.GetCTTc();
        var tcPr = ctTc.tcPr ?? ctTc.AddNewTcPr();
        var tcW = tcPr.tcW ?? tcPr.AddNewTcW();
        tcW.w = width.ToString();
        tcW.type = NPOI.OpenXmlFormats.Wordprocessing.ST_TblWidth.dxa; // Unitatea de masura pentru latime este DXA
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
            string companie = fields[columns["companie"]];

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

            Portofoliu portofoliu = new Portofoliu(stagiuDePractica, studentPracticant, tutore, profesorCoordonator, companie);
            colectiePortofolii.Add(portofoliu);
        }

        return colectiePortofolii;
    }
}
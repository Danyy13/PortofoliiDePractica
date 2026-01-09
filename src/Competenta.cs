public class Competenta
{
    public string descriere { get; }
    public string modPregatire { get; }
    public string locDeMunca { get; }
    public string activitati { get; }
    public string observatii { get; }

    public Competenta(string descriere, string modPregatire, string locDeMunca, string activitati, string observatii = "")
    {
        this.descriere = descriere;
        this.modPregatire = modPregatire;
        this.locDeMunca = locDeMunca;
        this.activitati = activitati;
        this.observatii = observatii;
    }
}
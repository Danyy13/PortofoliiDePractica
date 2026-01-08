public class Competenta
{
    private string descriere { get; }
    private string modPregatire { get; }
    private string locDeMunca { get; }
    private string activitati { get; }
    private string observatii { get; }

    public Competenta(string descriere, string modPregatire, string locDeMunca, string activitati, string observatii = "")
    {
        this.descriere = descriere;
        this.modPregatire = modPregatire;
        this.locDeMunca = locDeMunca;
        this.activitati = activitati;
        this.observatii = observatii;
    }
}
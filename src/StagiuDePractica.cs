public class StagiuDePractica
{
    private int id;
    public int durata { get; }
    public string calendar { get; }
    public PerioadaStagiu perioadaStagiu { get; }
    public string adresa { get; }
    public string locatiiDeplasare { get; }
    public string listaConditii { get; }
    public string modalitatiComplementaritate { get; }
    public List<Competenta> listaCompetente { get; }
    public string modalitatiEvaluare { get; }

    public StagiuDePractica(int id, int durata, string calendar, PerioadaStagiu perioadaStagiu, string adresa, string locatiiDeplasare, string listaConditii, string modalitatiComplementaritate, List<Competenta> listaCompetente, string modalitatiEvaluare)
    {
        this.id = id;
        this.durata = durata;
        this.calendar = calendar;
        this.perioadaStagiu = perioadaStagiu;
        this.adresa = adresa;
        this.locatiiDeplasare = locatiiDeplasare;
        this.listaConditii = listaConditii;
        this.modalitatiComplementaritate = modalitatiComplementaritate;
        this.listaCompetente = listaCompetente;
        this.modalitatiEvaluare = modalitatiEvaluare;
    }
}
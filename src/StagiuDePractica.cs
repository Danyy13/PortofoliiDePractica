public class StagiuDePractica
{
    private int id;
    private int durata { get; }
    private string calendar { get; }
    private PerioadaStagiu perioadaStagiu { get; }
    private string adresa { get; }
    private string locatiiDeplasare { get; }
    private string listaConditii { get; }
    private string modalitatiComplementaritate { get; }
    private List<Competenta> listaCompetente { get; }
    private string modalitatiEvaluare { get; }

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
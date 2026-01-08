public class Persoana
{
    protected int id;
    protected string nume;
    protected string prenume;
    protected string functie;
    protected DateTime dataSemnare;

    public Persoana(int id, string nume, string prenume, string functie, DateTime dataSemnare)
    {
        this.id = id;
        this.nume = nume;
        this.prenume = prenume;
        this.functie = functie;
        this.dataSemnare = dataSemnare;
    }
}
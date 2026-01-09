public class Persoana
{
    protected int id;
    public string nume { get; protected set;}
    public string prenume { get; protected set;}
    public string functie { get; protected set;}
    public DateTime dataSemnare { get; protected set;}

    public Persoana(int id, string nume, string prenume, string functie, DateTime dataSemnare)
    {
        this.id = id;
        this.nume = nume;
        this.prenume = prenume;
        this.functie = functie;
        this.dataSemnare = dataSemnare;
    }
}
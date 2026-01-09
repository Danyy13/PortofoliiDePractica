public class Tutore : Persoana
{
    public string responsabilitati { get; }

    public Tutore(int id, string nume, string prenume, string functie, DateTime dataSemnare, string responsabilitati) : base(id, nume, prenume, functie, dataSemnare)
    {
        this.id = id;
        this.nume = nume;
        this.prenume = prenume;
        this.functie = functie;
        this.dataSemnare = dataSemnare;
        this.responsabilitati = responsabilitati;
    }
    
}
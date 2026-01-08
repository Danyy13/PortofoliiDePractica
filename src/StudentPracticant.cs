public class StudentPracticant : Persoana
{
    public StudentPracticant(int id, string nume, string prenume, string functie, DateTime dataSemnare) : base(id, nume, prenume, functie, dataSemnare)
    {
        this.id = id;
        this.nume = nume;
        this.prenume = prenume;
        this.functie = functie;
        this.dataSemnare = dataSemnare;
    }
}
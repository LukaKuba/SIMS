using System;

public enum Tip { Farmaceut, Upravnik, Lekar };
public class Korisnik
{
    public string JMBG;
    public string email;
    public string lozinka;
    public string ime
    {get; set;}
    public string prezime;
    public string mobilni;
    public Tip tipkorisnika;
    public bool blokiran;

    
    public Korisnik(string ime, string prezime, string JMBG, string email, string mobilni, string lozinka, Tip tipkorisnika, bool blokiran)
    {
        this.ime = ime;
        this.prezime = prezime;
        this.JMBG = JMBG;
        this.email = email;
        this.mobilni = mobilni;
        this.lozinka = lozinka;
        this.tipkorisnika = tipkorisnika;
        bool bokiran = blokiran;
    }

    
}




using System;


public class Lek
{
	public string sifra
	{ get; set; }
	public string ime
	{ get; set; }
	public string proizvodjac
	{ get; set; }
	public int kolicina
	{ get; set; }
	public bool prihvacen;
	public bool obrisan;
	public bool odbijen;
	public int cena
	{ get; set; }
	public Dictionary<string, string> sastojci;
	public int brojf
	{ get; set; }
	public int brojl
	{ get; set; }
	public string obrazlozenje
	{ get; set; }



	public Lek(string sifra, string ime, string proizvodjac, int kolicina, bool prihvacen, bool obrisan, int cena, Dictionary<string, string> sastojci, bool odbijen, int brojf, int brojl, string obrazlozenje)
    {
        this.sifra = sifra;
        this.ime = ime;
        this.proizvodjac = proizvodjac;
        this.kolicina = kolicina;
        this.prihvacen = prihvacen;
        this.obrisan = obrisan;
        this.cena = cena;
        this.sastojci = sastojci;
        this.odbijen = odbijen;
		this.brojf = brojf;
		this.brojl = brojl;
		this.obrazlozenje = obrazlozenje;
    }

}

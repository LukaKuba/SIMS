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
	public int cena
	{ get; set; }


	public Lek(string sifra, string ime, string proizvodjac, int kolicina, bool prihvacen, bool obrisan, int cena)
    {
		this.sifra = sifra;
		this.ime = ime;	
		this.proizvodjac = proizvodjac;
		this.kolicina = kolicina;
		this.prihvacen = prihvacen;
		this.obrisan = obrisan;
		this.cena = cena;

    }

}

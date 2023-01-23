using System;

public class Sastojak
{
	public string ime
	{get; set;}
	public string opis
	{get; set;}	
	public Sastojak(string ime, string opis)
    {
		this.ime = ime;	
		this.opis = opis;
    }
}

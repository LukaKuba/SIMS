using System;
using Microsoft.Office.Interop.Excel;



class Program
{
    static void Main(string[] args)
    {

        List<Sastojak> arrSastojci = new List<Sastojak>();

        Sastojak sastojak = new Sastojak("Aa", "ALALALALA");
        arrSastojci.Add(sastojak);
        sastojak = new Sastojak("Bb", "BLBLBLBLB");
        arrSastojci.Add(sastojak);
        sastojak = new Sastojak("Cc", "CLCLCLCLC");
        arrSastojci.Add(sastojak);

        List<Korisnik> arrUsers = new List<Korisnik>();
        Korisnik korisnik = new Korisnik(" "," "," "," "," "," " ,Tip.Lekar, false);

        List<Lek> arrLekovi = new List<Lek>();

        Lek lek = new Lek("0001", "viekvin", "hemofarm", 10, true, false, 1000);
        arrLekovi.Add(lek);
        lek = new Lek("0002", "grafalon", "hemofarm", 14, false, false, 800);
        arrLekovi.Add(lek);
        lek = new Lek("0003", "aldipete", "hemofarm", 12, false, false, 700);
        arrLekovi.Add(lek);
        lek = new Lek("0004", "menactra", "hemofarm", 21, true, false, 900);
        arrLekovi.Add(lek);

        Application excelApp = new Application();
        Workbook excelBook = excelApp.Workbooks.Open(@"C:\Users\user\Desktop\Luka faks\SIMS\Book2.xlsx");
        _Worksheet excelSheet = excelBook.Sheets[1];
        Microsoft.Office.Interop.Excel.Range excelRange = excelSheet.UsedRange;
        int rows = excelRange.Rows.Count;
        int cols = excelRange.Columns.Count;
        string ime;
        string prezime;
        string JMBG;
        string email1;
        string mobilni;
        string lozinka1;
        string pozicija;

        for (int i = 1; i <= rows; i++)
        {
 
                
                ime = Convert.ToString((excelRange.Cells[i, 1]).Text);
                prezime = Convert.ToString((excelRange.Cells[i, 2]).Text);
                JMBG = Convert.ToString((excelRange.Cells[i, 3]).Text);
                email1 = Convert.ToString((excelRange.Cells[i, 4]).Text);
                mobilni = Convert.ToString((excelRange.Cells[i, 5]).Text);
                lozinka1 = Convert.ToString((excelRange.Cells[i, 6]).Text);
                pozicija = Convert.ToString((excelRange.Cells[i, 7]).Text);
                if (pozicija == "Lekar")
                {
                    korisnik = new Korisnik(ime, prezime, JMBG, email1, mobilni, lozinka1, Tip.Lekar, false);
                    arrUsers.Add(korisnik);
                }
                if (pozicija == "Upravnik")
                {
                    korisnik = new Korisnik(ime, prezime, JMBG, email1, mobilni, lozinka1, Tip.Upravnik, false);
                    arrUsers.Add(korisnik);
                }
                if (pozicija == "Farmaceut")
                {
                    korisnik = new Korisnik(ime, prezime, JMBG, email1, mobilni, lozinka1, Tip.Farmaceut, false);
                    arrUsers.Add(korisnik);
                }
            
        }
        excelBook.Close();






        string email=" ";
        bool successfull = false;
        int s = 0;
        while (successfull == false)
        {

            
            Console.WriteLine("Unesite vas email:");
            email = Console.ReadLine();
            Console.WriteLine("Unesite vasu lozinku:");
            string lozinka = Console.ReadLine();


            foreach (Korisnik user in arrUsers)if(user.blokiran==false)
            {
                if (email == user.email && lozinka == user.lozinka)
                {
                    Console.WriteLine("You have successfully logged in !!!");

                    successfull = true;
                    break;
                }
            }

                if (successfull == false)
                {
                    s = s + 1;
                    if(s==3)
                    {
                    Environment.Exit(0);
                    }
                    Console.WriteLine("Your username or password is incorect, try again!");
                    
                }

        

        }
        int broj=23;
        while(broj!=0)
        {
            Prikazimeni(arrUsers, email);
            broj = Convert.ToInt32(Console.ReadLine());
            switch (broj)
            {
                case 1:
                    Prikazilekove(arrLekovi, arrUsers);
                    break;
                case 2:

                    Pretragalekova(arrLekovi);
                    break;
                case 3:

                    Prikazlekova(arrLekovi, email, arrUsers);
                    Registracija(arrUsers, email, rows, excelRange, excelApp);
                    break;

                case 4:

                    Prikazkorisnika(arrUsers, email);
                    break;
                case 0:

                    Environment.Exit(0);
                    break;
                  
            }
        } 
        

    }

    static void Prikazimeni(List<Korisnik> arrUsers , string email)
    {
        foreach (Korisnik user in arrUsers)
        {
            if (user.tipkorisnika == Tip.Farmaceut && user.email==email)
            {
                Console.WriteLine("Meni: \n 1. Prikaz lekova \n " +
                   "2. Pretraga lekova \n " +
                   "3. Prikaz lekova koji cekaju potvrdu \n " +
                   "4. Odobravanje i odbijanje lekova \n " +
                   "5. Prikaz Prihvacenih i odbijenih lekova \n " +
                   "0. Exit");
                
            }
            if (user.tipkorisnika == Tip.Lekar && user.email == email)
            {
                Console.WriteLine("Meni: \n 1. Prikaz lekova \n " +
                   "2. Pretraga lekova \n " +
                   "3. Prikaz lekova koji cekaju potvrdu \n " +
                   "4. Odobravanje i odbijanje lekova \n " +
                   "0. Exit");
                
            }
            if (user.tipkorisnika == Tip.Upravnik && user.email == email)
            {
                Console.WriteLine("Meni: \n 1. Prikaz lekova \n " +
                   "2. Pretraga lekova \n " +
                   "3. Registracija \n " +
                   "4. Prikaz svih korisnika \n " +
                   "5. Blokiranje korisnika \n " +
                   "6. Unos lekova \n " +
                   "7. Nabavka lekova \n " +
                   "8. Nabavka lekova kroz vreme \n " +
                   "0. Exit");
                
            }
 
        }
    }
    static void Prikazilekove(List<Lek> arrLekovi, List<Korisnik> arrUsers)
    {
        Console.WriteLine("Izaberite vrstu sortiranja: \n" +
            "1: po imenu \n" +
            "2: po ceni \n" +
            "3: po kolicini");
        int j = Convert.ToInt32(Console.ReadLine());
        if (j == 1)
        {

            var enum1 = from lek in arrLekovi
                        orderby lek.ime
                        select lek;

            foreach (Lek lek in enum1)
            {

                Console.WriteLine(lek.ime + " " + lek.sifra + " " + lek.proizvodjac + " " + lek.cena + " " + lek.kolicina);
            }


        }
        if (j == 2)
        {
            
            var enum1 = from lek in arrLekovi
                        orderby lek.cena
                        select lek;

            foreach (Lek lek in enum1)
            {
              
                Console.WriteLine(lek.ime + " " + lek.sifra + " " + lek.proizvodjac + " " + lek.cena + " " + lek.kolicina);
            }


        }
        if (j == 3)
        {

            var enum1 = from lek in arrLekovi
                        orderby lek.kolicina
                        select lek;

            foreach (Lek lek in enum1)
            {

                Console.WriteLine(lek.ime + " " + lek.sifra + " " + lek.proizvodjac + " " + lek.cena + " " + lek.kolicina);
            }


        } 

    }
    static void Pretragalekova(List<Lek> arrLekovi)
    {
        int j;
        do
        {
            Console.WriteLine("Izaberi nacin pretrage: \n " +
                "1. pretraga po sifri \n " +
                "2. pretraga po imenu \n " +
                "3. pretraga po proizvodjacu \n " +
                "4. pretraga po ceni \n " +
                "5. pretraga po kolicini \n " +
                "6. pretraga po sastojku \n " +
                "0. nazad");
            j = Convert.ToInt32(Console.ReadLine());
            if (j == 1)
            {
                Console.WriteLine("Unesite pretragu:");
                string model = Console.ReadLine();
                foreach (Lek lek in arrLekovi)
                {
                    if (lek.sifra == model)

                    {
                        Console.WriteLine(lek.ime + " " + lek.sifra + " " + lek.proizvodjac + " " + lek.cena + " " + lek.kolicina + " \n ");
                    }
                }
            }
            if (j == 2)
            {
                Console.WriteLine("Unesite pretragu:");
                string model = Console.ReadLine();
                foreach (Lek lek in arrLekovi)
                {
                    if (lek.ime == model || lek.ime.Contains(model))

                    {
                        Console.WriteLine(lek.ime + " " + lek.sifra + " " + lek.proizvodjac + " " + lek.cena + " " + lek.kolicina + " \n ");
                    }
                }
            }
            if (j == 3)
            {
                Console.WriteLine("Unesite pretragu:");
                string model = Console.ReadLine();
                foreach (Lek lek in arrLekovi)
                {
                    if (lek.proizvodjac == model || lek.proizvodjac.Contains(model))

                    {
                        Console.WriteLine(lek.ime + " " + lek.sifra + " " + lek.proizvodjac + " " + lek.cena + " " + lek.kolicina + " \n ");
                    }
                }

            }

            if (j == 4)
            {
                Console.WriteLine("Cena od:");
                int m1 = Convert.ToInt32(Console.ReadLine());
                Console.WriteLine("Cena do:");
                int m2 = Convert.ToInt32(Console.ReadLine());
                foreach (Lek lek in arrLekovi)
                {
                    if (lek.cena >= m1 && lek.cena <= m2)

                    {
                        Console.WriteLine(lek.ime + " " + lek.sifra + " " + lek.proizvodjac + " " + lek.cena + " " + lek.kolicina + " \n ");
                    }
                }
            }
            if (j == 5)
            {
                Console.WriteLine("Kolicina od:");
                int m1 = Convert.ToInt32(Console.ReadLine());
                Console.WriteLine("Kolicina do:");
                int m2 = Convert.ToInt32(Console.ReadLine());
                foreach (Lek lek in arrLekovi)
                {
                    if (lek.kolicina >= m1 && lek.kolicina <= m2)

                    {
                        Console.WriteLine(lek.ime + " " + lek.sifra + " " + lek.proizvodjac + " " + lek.cena + " " + lek.kolicina + " \n ");
                    }
                }
            }

        } while (j != 0);

    }
    static void Prikazlekova(List<Lek> arrLekovi, string email, List<Korisnik> arrUsers)
    {
        foreach (Korisnik user in arrUsers)
        {
            if ((user.tipkorisnika == Tip.Farmaceut || user.tipkorisnika == Tip.Lekar) && user.email == email)
            {
                foreach (Lek lek in arrLekovi)
                {
                    if(lek.prihvacen==false)
                    Console.WriteLine(lek.ime + " " + lek.sifra + " " + lek.proizvodjac + " " + lek.cena + " " + lek.kolicina);
                }
            }
        }
        Console.ReadLine();
    }
    static void Registracija(List<Korisnik> arrUsers, string email, int rows, Microsoft.Office.Interop.Excel.Range excelRange, Application excelApp)
    {
        foreach(Korisnik user in arrUsers)
        {
            if(user.email== email && user.tipkorisnika==Tip.Upravnik)
            {
                int pr1 = 0;
                while (pr1 == 0)
                {
                    
                    Console.WriteLine(" unesite ime:");
                    string ime = Console.ReadLine();
                    Console.WriteLine(" unesite rezime:");
                    string prezime = Console.ReadLine();
                    Console.WriteLine(" unesite JMBG:");
                    string JMBG = Console.ReadLine();
                    Console.WriteLine(" unesite email:");
                    string email1 = Console.ReadLine();
                    Console.WriteLine(" unesite lozinka:");
                    string lozinka = Console.ReadLine();
                    Console.WriteLine(" unesite mobilni:");
                    string mobilni = Console.ReadLine();
                    Console.WriteLine(" Izaberite tip korisnika \n " +
                        "1. Farmaceut \n " +
                        "2. Lekar");
                    int tip=Convert.ToInt32(Console.ReadLine());
                    
                    foreach (Korisnik user1 in arrUsers)
                    {
                        if(user1.email != email1 && user1.JMBG !=JMBG)
                        {
                            if (tip == 1)
                            {
                                pr1 = 1;
                                Korisnik korisnik = new Korisnik(ime, prezime, JMBG, email1, mobilni, lozinka, Tip.Farmaceut, false);
                                arrUsers.Add(korisnik);
                                /*int curen = rows + 1;
                                excelRange.Cells[curen, 1].Value = Convert.ToString(ime);
                                excelRange.Cells[curen, 2].Value = Convert.ToString(prezime);
                                excelRange.Cells[curen, 3].Value = Convert.ToString(JMBG);
                                excelRange.Cells[curen, 4].Value = Convert.ToString(email1);
                                excelRange.Cells[curen, 5].Value = Convert.ToString(mobilni);
                                excelRange.Cells[curen, 6].Value = Convert.ToString(lozinka);
                                excelRange.Cells[curen, 7].Value = ("Farmaceut");
                                excelApp.ActiveWorkbook.SaveAs(@"C:\Users\user\Desktop\Luka faks\SIMS\Book2.xlsx", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault);
                                excelApp.Quit();*/
                            }
                            if (tip == 2)
                            {
                                pr1 = 1;
                                Korisnik korisnik = new Korisnik(ime, prezime, JMBG, email1, mobilni, lozinka, Tip.Lekar, false);
                                arrUsers.Add(korisnik);
                                /*int curen = rows + 1;
                                excelRange.Cells[curen, 1].Value = Convert.ToString(ime);
                                excelRange.Cells[curen, 2].Value = Convert.ToString(prezime);
                                excelRange.Cells[curen, 3].Value = Convert.ToString(JMBG);
                                excelRange.Cells[curen, 4].Value = Convert.ToString(email1);
                                excelRange.Cells[curen, 5].Value = Convert.ToString(mobilni);
                                excelRange.Cells[curen, 6].Value = Convert.ToString(lozinka);
                                excelRange.Cells[curen, 7].Value = ("Lekar");
                                excelApp.ActiveWorkbook.SaveAs(@"C:\Users\user\Desktop\Luka faks\SIMS\Book2.xlsx", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault);
                                excelApp.Quit();*/

                            }
                            rows = rows + 1;
                            Console.WriteLine("Uspesno ste dodali korisnika: ");
                        }
                        else
                        {
                            Console.WriteLine("korisnik vec postoji u sistemu, pokusajte ponovo ");
                        }
                    }
                }
                
               

            }
        }
    }
    static void Prikazkorisnika(List<Korisnik> arrUsers,string email)
    {
        foreach (Korisnik user1 in arrUsers)
        {
            if (user1.email == email && user1.tipkorisnika == Tip.Upravnik)
            {
                foreach (Korisnik user in arrUsers)
                {
                    Console.WriteLine(user.ime + " " + user.prezime + " " + user.JMBG + " " + user.email + " " + user.mobilni + " " + user.tipkorisnika);
                }
            }
        }
    }
    static void Blokiranjekorisnika(List<Korisnik> arrusers)
    {
        Console.WriteLine("Unesite JMBG korisnika:");
        string JMBG=Console.ReadLine();
        foreach (Korisnik user in arrusers)
        {
            if (user.JMBG == JMBG);
            {
                user.blokiran = true;
            }
        }

    }
}

        


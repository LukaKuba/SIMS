using System;
using Microsoft.Office.Interop.Excel;



class Program
{
    static void Main(string[] args)
    {

        List<Sastojak> arrSastojci = new List<Sastojak>();
        List<Korisnik> arrUsers = new List<Korisnik>();
        List<Lek> arrLekovi = new List<Lek>();

        Sastojak sastojak = new Sastojak("", "");

        Dictionary<string, string> TempDic = new Dictionary<string, string>();

        Lek lek = new Lek("", "", "", 0, true, false, 0, TempDic, false, 0, 0, "");

        Application excelApp2 = new Application();
        Workbook excelBook2 = excelApp2.Workbooks.Open(@"C:\Users\user\Desktop\Luka\SIMS\Sastojci.xlsx");
        _Worksheet excelSheet2 = excelBook2.Sheets[1];
        Microsoft.Office.Interop.Excel.Range excelRange2 = excelSheet2.UsedRange;
        int rows2 = excelRange2.Rows.Count;
        int cols2 = excelRange2.Columns.Count;

        string imee;
        string opis;

        for (int i = 1; i <= rows2; i++)
        {
            imee = Convert.ToString((excelRange2.Cells[i, 1]).Text);
            opis = Convert.ToString((excelRange2.Cells[i, 2]).Text);
            sastojak = new Sastojak(imee, opis);
            arrSastojci.Add(sastojak);
        }
        excelApp2.Quit();
        System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp2);

        Application excelApp1 = new Application();
        Workbook excelBook1 = excelApp1.Workbooks.Open(@"C:\Users\user\Desktop\Luka\SIMS\Lekovi.xlsx");
        _Worksheet excelSheet1 = excelBook1.Sheets[1];
        Microsoft.Office.Interop.Excel.Range excelRange1 = excelSheet1.UsedRange;
        int rows1 = excelRange1.Rows.Count;
        int cols1 = excelRange1.Columns.Count;

        string sifra;
        string ime1;
        string proizvodjac;
        string kolicina1;
        int kolicina;
        string cena1;
        int cena;
        string obrisan;
        string prihvacen;
        bool obrisan1;
        bool prihvacen1;
        string sastojak_ime;
        string odbijen;
        bool odbijen1;
        int brojl;
        int brojf;
        string obrazlozenje;
        string brojl1;
        string brojf1;




        for (int i = 1; i <= rows1; i++)
        {

            Dictionary<string, string> TempDic1 = new Dictionary<string, string>();



            ime1 = Convert.ToString((excelRange1.Cells[i, 2]).Text);
            proizvodjac = Convert.ToString((excelRange1.Cells[i, 3]).Text);
            sifra = Convert.ToString((excelRange1.Cells[i, 1]).Text);
            kolicina1 = Convert.ToString((excelRange1.Cells[i, 4]).Text);
            kolicina = Convert.ToInt32(kolicina1);
            cena1 = Convert.ToString((excelRange1.Cells[i, 7]).Text);
            cena = Convert.ToInt32(cena1);
            obrisan = Convert.ToString((excelRange1.Cells[i, 6]).Text);
            prihvacen = Convert.ToString((excelRange1.Cells[i, 5]).Text);
            odbijen = Convert.ToString((excelRange1.Cells[i, 9]).Text);
            sastojak_ime = Convert.ToString((excelRange1.Cells[i, 8]).Text);
            brojl1 = Convert.ToString((excelRange1.Cells[i, 10]).Text);
            brojl = Convert.ToInt32(brojl1);
            brojf1 = Convert.ToString((excelRange1.Cells[i, 11]).Text);
            brojf = Convert.ToInt32(brojf1);
            obrazlozenje = Convert.ToString((excelRange1.Cells[i, 12]).Text);

            if (obrisan == "TRUE")
            {
                obrisan1 = true;
            }
            else
            {
                obrisan1 = false;
            }
            if (odbijen == "TRUE")
            {
                odbijen1 = true;
            }
            else
            {
                odbijen1 = false;
            }
            if (prihvacen == "TRUE")
            {
                prihvacen1 = true;
            }
            else
            {
                prihvacen1 = false;
            }

            IList<string> names = sastojak_ime.Split(',').ToList<string>();
            for (int j = 0; j < names.Count; j++)
            {
                names[j] = names[j].Replace(" ", String.Empty);

                for (int g = 0; g < arrSastojci.Count; g++)
                {
                    if (names[j] == arrSastojci[g].ime)
                    {
                        TempDic1.Add(arrSastojci[g].ime, arrSastojci[g].opis);
                    }

                }


            }
            lek = new Lek(sifra, ime1, proizvodjac, kolicina, prihvacen1, obrisan1, cena, TempDic1, odbijen1, brojl, brojf, obrazlozenje);
            arrLekovi.Add(lek);



        }
        excelApp1.Quit();
        System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp1);

        Korisnik korisnik = new Korisnik(" ", " ", " ", " ", " ", " ", Tip.Upravnik, false);

        Application excelApp = new Application();
        Workbook excelBook = excelApp.Workbooks.Open(@"C:\Users\user\Desktop\Luka\SIMS\Book2.xlsx");
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
        string blokiran1;
        bool blokiran;

        for (int i = 1; i <= rows; i++)
        {


            ime = Convert.ToString((excelRange.Cells[i, 1]).Text);
            prezime = Convert.ToString((excelRange.Cells[i, 2]).Text);
            JMBG = Convert.ToString((excelRange.Cells[i, 3]).Text);
            email1 = Convert.ToString((excelRange.Cells[i, 4]).Text);
            mobilni = Convert.ToString((excelRange.Cells[i, 5]).Text);
            lozinka1 = Convert.ToString((excelRange.Cells[i, 6]).Text);
            pozicija = Convert.ToString((excelRange.Cells[i, 7]).Text);
            blokiran1 = Convert.ToString((excelRange.Cells[i, 7]).Text);
            if (blokiran1 == "TRUE")
            {
                blokiran = true;
            }
            else
            {
                blokiran = false;
            }
            if (pozicija == "Lekar")
            {
                korisnik = new Korisnik(ime, prezime, JMBG, email1, mobilni, lozinka1, Tip.Lekar, blokiran);
                arrUsers.Add(korisnik);
            }
            if (pozicija == "Upravnik")
            {
                korisnik = new Korisnik(ime, prezime, JMBG, email1, mobilni, lozinka1, Tip.Upravnik, blokiran);
                arrUsers.Add(korisnik);
            }
            if (pozicija == "Farmaceut")
            {
                korisnik = new Korisnik(ime, prezime, JMBG, email1, mobilni, lozinka1, Tip.Farmaceut, blokiran);
                arrUsers.Add(korisnik);
            }

        }
        excelApp.Quit();
        System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);


        string email = " ";
        bool successfull = false;
        int s = 0;
        while (successfull == false)
        {


            Console.WriteLine("Unesite vas email:");
            email = Console.ReadLine();
            Console.WriteLine("Unesite vasu lozinku:");
            string lozinka = Console.ReadLine();


            foreach (Korisnik user in arrUsers) if (user.blokiran == false)
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
                if (s == 3)
                {
                    Environment.Exit(0);
                }
                Console.WriteLine("Your username or password is incorect, try again!");

            }



        }
        int broj = 23;
        while (broj != 0)
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
                    foreach (Korisnik user in arrUsers)
                    {
                        if ((user.tipkorisnika == Tip.Farmaceut || user.tipkorisnika == Tip.Lekar) && user.email == email)
                        {
                            Cekajupotvrdu(arrLekovi);
                        }
                    }
                    foreach (Korisnik user in arrUsers)
                    {
                        if (user.tipkorisnika == Tip.Upravnik && email == user.email)
                        {
                            Registracija(arrUsers, email);
                        }
                        break;
                    }
                    break;

                case 4:
                    foreach (Korisnik user in arrUsers)
                    {
                        if (user.tipkorisnika == Tip.Upravnik && email == user.email)
                        {
                            Prikazkorisnika(arrUsers, email);
                        }
                        if (user.tipkorisnika == Tip.Lekar && user.email == email)
                        {

                            Prihvatanjeiodbijanjel(arrUsers, arrLekovi, email);
                        }
                        if (user.tipkorisnika == Tip.Farmaceut && user.email == email)
                        {

                            Prihvatanjeiodbijanjef(arrUsers, arrLekovi, email);
                        }
                    }
                    break;
                case 5:
                    foreach (Korisnik user in arrUsers)
                    {
                        if (user.tipkorisnika == Tip.Upravnik && email == user.email)
                        {
                            Blokiranjekorisnika(arrUsers);
                        }
                        if (user.tipkorisnika == Tip.Farmaceut && user.email == email)
                        {

                            Prikazpio(arrUsers, arrLekovi);
                        }
                    }
                    break;
                case 6:
                    foreach (Korisnik user in arrUsers)
                    {
                        if (user.tipkorisnika == Tip.Upravnik && email == user.email)
                        {
                            Unoslekova(arrLekovi, arrSastojci, arrUsers);
                        }
                    }
                    break;
                case 7:
                    foreach (Korisnik user in arrUsers)
                    {
                        if (user.tipkorisnika == Tip.Upravnik && email == user.email)
                        {
                            Nabavkalekova(arrUsers, arrLekovi);
                        }
                    }
                    break;
                case 8:
                    foreach (Korisnik user in arrUsers)
                    {
                        if (user.tipkorisnika == Tip.Upravnik && email == user.email)
                        {

                        }
                    }
                    break;
                case 0:

                    Environment.Exit(0);
                    break;

            }
        }


    }

    static void Prikazimeni(List<Korisnik> arrUsers, string email)
    {
        foreach (Korisnik user in arrUsers)
        {
            if (user.tipkorisnika == Tip.Farmaceut && user.email == email)
            {
                Console.WriteLine("Meni: \n 1. Prikaz lekova \n " +
                   "2. Pretraga lekova \n " +
                   "3. Prikaz lekova koji cekaju potvrdu \n " +
                   "4. Odobravanje i odbijanje lekova \n " +
                   "5. Prikaz prihvacenih i odbijenih lekova \n " +
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
                if (lek.prihvacen == true)
                {
                    Console.WriteLine(lek.ime + " " + lek.sifra + " " + lek.proizvodjac + " " + lek.cena + " " + lek.kolicina);
                }
            }


        }
        if (j == 2)
        {

            var enum1 = from lek in arrLekovi
                        orderby lek.cena
                        select lek;

            foreach (Lek lek in enum1)
            {
                if (lek.prihvacen == true)
                {
                    Console.WriteLine(lek.ime + " " + lek.sifra + " " + lek.proizvodjac + " " + lek.cena + " " + lek.kolicina);
                }
            }


        }
        if (j == 3)
        {

            var enum1 = from lek in arrLekovi
                        orderby lek.kolicina
                        select lek;

            foreach (Lek lek in enum1)
            {
                if (lek.prihvacen == true)
                {
                    Console.WriteLine(lek.ime + " " + lek.sifra + " " + lek.proizvodjac + " " + lek.cena + " " + lek.kolicina);
                }
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

                        foreach (KeyValuePair<string, string> kvp in lek.sastojci)
                        {

                            Console.WriteLine("Sastojak = {0}, Opis = {1}", kvp.Key, kvp.Value + "\n");
                        }
                    }
                }
            }
            if (j == 6)
            {
                Console.WriteLine("Unesite saastojke");
                string sastojci1 = Console.ReadLine();
                sastojci1 = sastojci1.Replace(" ", String.Empty);
                IList<string> lista = sastojci1.Split('&').ToList<string>();
                foreach (Lek lek in arrLekovi)
                {
                    int broj = lista.Count;
                    int counter = 0;
                    foreach (KeyValuePair<string, string> kvp in lek.sastojci)
                    {
                        for (int i = 0; i < lista.Count; i++)
                        {
                            string s = Convert.ToString(kvp.Key);
                            if (lista[i].Contains("|"))
                            {
                                IList<string> lista2 = lista[i].Split('|').ToList<string>();

                                for (int j1 = 0; j1 < lista2.Count; j1++)
                                {
                                    if (s == lista2[j1])
                                    {
                                        broj --;
                                        counter ++;

                                    }
                                }
                            }
                            else
                            {
                                if (lista[i] == s)
                                {
                                    broj ++;
                                }
                            }
                        }


                    }
                    if (counter > 1)
                    {
                        broj = broj + counter - 1;
                    }
                    if (broj == 0)
                    {
                        Console.WriteLine(lek.ime + " " + lek.sifra + " " + lek.proizvodjac + " " + lek.cena + " " + lek.kolicina + " \n ");

                        foreach (KeyValuePair<string, string> kvp1 in lek.sastojci)
                        {

                            Console.WriteLine("Sastojak = {0}, Opis = {1}", kvp1.Key, kvp1.Value + "\n");
                        }

                    }


                }
            }

        } while (j != 0);

    }
    static void Registracija(List<Korisnik> arrUsers, string email)
    {
        Application excelApp = new Application();
        Workbook excelBook = excelApp.Workbooks.Open(@"C:\Users\user\Desktop\Luka\SIMS\Book2.xlsx");
        _Worksheet excelSheet = excelBook.Sheets[1];
        Microsoft.Office.Interop.Excel.Range excelRange = excelSheet.UsedRange;
        int rows = excelRange.Rows.Count;
        int cols = excelRange.Columns.Count;

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
            Console.WriteLine(" unesite lozinku:");
            string lozinka = Console.ReadLine();
            Console.WriteLine(" unesite mobilni:");
            string mobilni = Console.ReadLine();
            Console.WriteLine(" Izaberite tip korisnika \n " +
                "1. Farmaceut \n " +
                "2. Lekar");
            int tip = Convert.ToInt32(Console.ReadLine());

            foreach (Korisnik user1 in arrUsers)
            {
                if (user1.email != email1 && user1.JMBG != JMBG)
                {
                    if (tip == 1)
                    {
                        pr1 = 1;
                        Korisnik korisnik = new Korisnik(ime, prezime, JMBG, email1, mobilni, lozinka, Tip.Farmaceut, false);
                        arrUsers.Add(korisnik);
                        int curen = rows + 1;
                        excelRange.Cells[curen, 1] = Convert.ToString(ime);
                        excelRange.Cells[curen, 2] = Convert.ToString(prezime);
                        excelRange.Cells[curen, 3] = Convert.ToString(JMBG);
                        excelRange.Cells[curen, 4] = Convert.ToString(email1);
                        excelRange.Cells[curen, 5] = Convert.ToString(mobilni);
                        excelRange.Cells[curen, 6] = Convert.ToString(lozinka);
                        excelRange.Cells[curen, 7] = ("Farmaceut");
                        excelRange.Cells[curen, 8] = ("FALSE");
                        excelApp.ActiveWorkbook.Save();
                        excelApp.Quit();
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                    }
                    if (tip == 2)
                    {
                        pr1 = 1;

                        Korisnik korisnik = new Korisnik(ime, prezime, JMBG, email1, mobilni, lozinka, Tip.Lekar, false);
                        arrUsers.Add(korisnik);
                        int curen = rows + 1;
                        excelRange.Cells[curen, 1] = Convert.ToString(ime);
                        excelRange.Cells[curen, 2] = Convert.ToString(prezime);
                        excelRange.Cells[curen, 3] = Convert.ToString(JMBG);
                        excelRange.Cells[curen, 4] = Convert.ToString(email1);
                        excelRange.Cells[curen, 5] = Convert.ToString(mobilni);
                        excelRange.Cells[curen, 6] = Convert.ToString(lozinka);
                        excelRange.Cells[curen, 7] = ("Lekar");
                        excelRange.Cells[curen, 8] = ("FALSE");
                        excelApp.ActiveWorkbook.Save();
                        excelApp.Quit();
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);

                    }
                    rows = rows + 1;
                    Console.WriteLine("Uspesno ste dodali korisnika ");
                    break;
                }
                else
                {
                    Console.WriteLine("korisnik vec postoji u sistemu, pokusajte ponovo ");
                }
            }
        }



    }


    static void Prikazkorisnika(List<Korisnik> arrUsers, string email)
    {
        Console.WriteLine("Izaberite vrstu sortiranja: \n" +
    "1: po imenu \n" +
    "2: po prezimenu ");
        int j = Convert.ToInt32(Console.ReadLine());
        Console.WriteLine("Izaberite tip korisnika: \n" +
    "1: Lekari \n" +
    "2: Farmaceuti \n" +
    "3: Svi korisnici");
        int k = Convert.ToInt32(Console.ReadLine());
        if (j == 1)
        {

            var enum1 = from user in arrUsers
                        orderby user.ime
                        select user;

            if (k == 1)
            {
                foreach (Korisnik user in arrUsers)
                {
                    if (user.tipkorisnika == Tip.Lekar)
                    {
                        Console.WriteLine(user.ime + " " + user.prezime + " " + user.JMBG + " " + user.email + " " + user.mobilni + " " + user.tipkorisnika);
                    }
                }
            }
            if (k == 2)
            {
                foreach (Korisnik user in arrUsers)
                {
                    if (user.tipkorisnika == Tip.Farmaceut)
                    {
                        Console.WriteLine(user.ime + " " + user.prezime + " " + user.JMBG + " " + user.email + " " + user.mobilni + " " + user.tipkorisnika);
                    }
                }
            }
            if (k == 3)
            {
                foreach (Korisnik user in arrUsers)
                {
                    Console.WriteLine(user.ime + " " + user.prezime + " " + user.JMBG + " " + user.email + " " + user.mobilni + " " + user.tipkorisnika);
                }
            }

        }
        if (j == 2)
        {

            var enum1 = from user in arrUsers
                        orderby user.prezime
                        select user;

            if (k == 1)
            {
                foreach (Korisnik user in arrUsers)
                {
                    if (user.tipkorisnika == Tip.Lekar)
                    {
                        Console.WriteLine(user.ime + " " + user.prezime + " " + user.JMBG + " " + user.email + " " + user.mobilni + " " + user.tipkorisnika);
                    }
                }
            }
            if (k == 2)
            {
                foreach (Korisnik user in arrUsers)
                {
                    if (user.tipkorisnika == Tip.Farmaceut)
                    {
                        Console.WriteLine(user.ime + " " + user.prezime + " " + user.JMBG + " " + user.email + " " + user.mobilni + " " + user.tipkorisnika);
                    }
                }
            }
            if (k == 3)
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
        string JMBG = Console.ReadLine();
        int j = 0;
        Console.ReadLine();
        Application excelApp = new Application();
        Workbook excelBook = excelApp.Workbooks.Open(@"C:\Users\user\Desktop\Luka\SIMS\Book2.xlsx");
        _Worksheet excelSheet = excelBook.Sheets[1];
        Microsoft.Office.Interop.Excel.Range excelRange = excelSheet.UsedRange;
        int rows = excelRange.Rows.Count;
        int cols = excelRange.Columns.Count;
        foreach (Korisnik user in arrusers)
        {
            if (user.JMBG == JMBG && user.blokiran == false)
            {
                Console.WriteLine("Korisnik nije blokiran, da li zeline da ga blokirate? \n 1. Da \n 2. Ne");
                int i = Convert.ToInt32(Console.ReadLine());
                if (i == 1)
                {
                    user.blokiran = true;
                }
            }
            else if (user.JMBG == JMBG && user.blokiran == true)
            {
                Console.WriteLine("Korisnik je blokiran, da li zeline da ga odblokirate? \n 1. Da \n 2. Ne");
                int i = Convert.ToInt32(Console.ReadLine());
                if (i == 2)
                {
                    user.blokiran = false;
                }
            }
            j = j + 1;

            excelRange.Cells[j, 8] = Convert.ToString(user.blokiran);
        }
        excelApp.ActiveWorkbook.Save();
        excelApp.Quit();
        System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
    }
    static void Nabavkalekova(List<Korisnik> arrUsers, List<Lek> arrLekovi)
    {
        Console.WriteLine("Unesite sifru leka: ");
        string a = Console.ReadLine();
        Console.WriteLine("Unesite kolicinu: ");
        string sad= Console.ReadLine();
        int b = Convert.ToInt32(sad);
        int nova = 0;
        foreach (Lek lek in arrLekovi)
        {
            if (lek.sifra == a)
            {
                lek.kolicina = lek.kolicina + b;
                nova = lek.kolicina;

            }

        }

        Application excelApp = new Application();
        Workbook excelBook = excelApp.Workbooks.Open(@"C:\Users\user\Desktop\Luka\SIMS\Lekovi.xlsx");
        _Worksheet excelSheet = excelBook.Sheets[1];
        Microsoft.Office.Interop.Excel.Range excelRange = excelSheet.UsedRange;
        int rows = excelRange.Rows.Count;
        int cols = excelRange.Columns.Count;
        for (int i = 1; i <= rows; i++)
        {
           
            if (a.Equals(excelRange.Cells[i, 1]))
            {
                excelRange.Cells[i, 4] = Convert.ToString(nova);
            }
        }

        excelApp.ActiveWorkbook.Save();
        excelApp.Quit();
        System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
    }
    static void Unoslekova(List<Lek> arrLekovi, List<Sastojak> arrSastojci, List<Korisnik> arrUsers)
    {
        Console.WriteLine(" unesite sifru");
        string sifra = Console.ReadLine();
        Console.WriteLine(" unesite ime:");
        string ime = Console.ReadLine();
        Console.WriteLine(" unesite proizvodjaca:");
        string proizvodjac = Console.ReadLine();
        Console.WriteLine(" unesite cenu:");
        int cena = Convert.ToInt32(Console.ReadLine());
        Console.WriteLine(" unesite sastojke:");
        string sastojak_ime = Console.ReadLine();
        Dictionary<string, string> TempDic = new Dictionary<string, string>();
        IList<string> names = sastojak_ime.Split(',').ToList<string>();
        for (int j = 0; j < names.Count; j++)
        {
            names[j] = names[j].Replace(" ", String.Empty);

            for (int g = 0; g < arrSastojci.Count; g++)
            {
                if (names[j] == arrSastojci[g].ime)
                {
                    TempDic.Add(arrSastojci[g].ime, arrSastojci[g].opis);
                }

            }


        }
        Lek lek = new Lek(sifra, ime, proizvodjac, 0, false, false, cena, TempDic, false, 0, 0, "");
        Application excelApp = new Application();
        Workbook excelBook = excelApp.Workbooks.Open(@"C:\Users\user\Desktop\Luka\SIMS\Lekovi.xlsx");
        _Worksheet excelSheet = excelBook.Sheets[1];
        Microsoft.Office.Interop.Excel.Range excelRange = excelSheet.UsedRange;
        int rows = excelRange.Rows.Count;
        int cols = excelRange.Columns.Count;
        int curen = rows + 1;
        excelRange.Cells[curen, 1] = Convert.ToString(sifra);
        excelRange.Cells[curen, 2] = Convert.ToString(ime);
        excelRange.Cells[curen, 3] = Convert.ToString(proizvodjac);
        excelRange.Cells[curen, 4] = Convert.ToString("0");
        excelRange.Cells[curen, 5] = ("FALSE");
        excelRange.Cells[curen, 6] = ("FALSE");
        excelRange.Cells[curen, 9] = ("FALSE");
        excelRange.Cells[curen, 7] = Convert.ToString(cena);
        excelRange.Cells[curen, 8] = Convert.ToString(sastojak_ime);
        excelRange.Cells[curen, 10] = ("0");
        excelRange.Cells[curen, 11] = ("0");


        excelApp.ActiveWorkbook.Save();
        excelApp.Quit();
        System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);

    }
    static void Prihvatanjeiodbijanjef(List<Korisnik> arrUsers, List<Lek> arrLekovi, string email)
    {
        Application excelApp = new Application();
        Workbook excelBook = excelApp.Workbooks.Open(@"C:\Users\user\Desktop\Luka\SIMS\Lekovi.xlsx");
        _Worksheet excelSheet = excelBook.Sheets[1];
        Microsoft.Office.Interop.Excel.Range excelRange = excelSheet.UsedRange;
        int rows = excelRange.Rows.Count;
        int cols = excelRange.Columns.Count;
        int i = 0;
        foreach (Lek lek in arrLekovi)
        {
            i++;
            if (lek.prihvacen == false && lek.odbijen == false)
            {
                Console.WriteLine(lek.ime + " " + lek.sifra + " " + lek.proizvodjac + " " + lek.cena + " " + lek.kolicina + " \n ");

                foreach (KeyValuePair<string, string> kvp in lek.sastojci)
                {

                    Console.WriteLine("Sastojak = {0}, Opis = {1}", kvp.Key, kvp.Value + "\n");
                }
                Console.WriteLine("1. Prihvati lek \n " +
                "2. Odbij lek \n " +
                "3. Preskoci");
                int br = Convert.ToInt32(Console.ReadLine());
                if (br == 1)
                {
                    lek.brojf = lek.brojf + 1;
                    excelRange.Cells[i, 10] = Convert.ToString(lek.brojf);
                    if (lek.brojf >= 3 && lek.brojl >= 1)
                    {
                        lek.prihvacen = true;
                        excelRange.Cells[i, 5] = ("TRUE");

                    }
                }
                if (br == 2)
                {
                    lek.odbijen = true;
                    Console.WriteLine("Unesite obrazlozenje");
                    string neki = Console.ReadLine();
                    string imeprez = "";


                    excelRange.Cells[i, 9] = ("TRUE");
                    foreach (Korisnik user in arrUsers)
                    {
                        if (email == user.email)
                        {
                            imeprez = ("Farmaceut" + " " + user.ime + " " + user.prezime);
                           
                        }
                    }
                    lek.obrazlozenje = (neki + " " + imeprez);
                    excelRange.Cells[i, 12] = Convert.ToString(lek.obrazlozenje);
                }
            }
        }
        excelApp.ActiveWorkbook.Save();
        excelApp.Quit();
        System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
    }
    static void Prihvatanjeiodbijanjel(List<Korisnik> arrUsers, List<Lek> arrLekovi, string email)
    {
        Application excelApp = new Application();
        Workbook excelBook = excelApp.Workbooks.Open(@"C:\Users\user\Desktop\Luka\SIMS\Lekovi.xlsx");
        _Worksheet excelSheet = excelBook.Sheets[1];
        Microsoft.Office.Interop.Excel.Range excelRange = excelSheet.UsedRange;
        int rows = excelRange.Rows.Count;
        int cols = excelRange.Columns.Count;
        int i = 0;
        foreach (Lek lek in arrLekovi)
        {
            i++;
            if (lek.prihvacen == false && lek.odbijen == false)
            {
                Console.WriteLine(lek.ime + " " + lek.sifra + " " + lek.proizvodjac + " " + lek.cena + " " + lek.kolicina + " \n ");

                foreach (KeyValuePair<string, string> kvp in lek.sastojci)
                {

                    Console.WriteLine("Sastojak = {0}, Opis = {1}", kvp.Key, kvp.Value + "\n");
                }
                Console.WriteLine("1. Prihvati lek \n " +
                "2. Odbij lek \n " +
                "3. Preskoci");
                int br = Convert.ToInt32(Console.ReadLine());
                if (br == 1)
                {
                    lek.brojl = lek.brojl + 1;
                    excelRange.Cells[i, 11] = Convert.ToString(lek.brojl);
                    if (lek.brojf >= 3 && lek.brojl >= 1)
                    {
                        lek.prihvacen = true;
                        excelRange.Cells[i, 5] = ("TRUE");

                    }
                }
                if (br == 2)
                {
                    lek.odbijen = true;
                    Console.WriteLine("Unesite obrazlozenje: ");
                    string neki = Console.ReadLine();
                    string imeprez= "";


                    excelRange.Cells[i, 9] = ("TRUE");
                    foreach (Korisnik user in arrUsers)
                    {
                        if (email == user.email)
                        {
                            imeprez = ("Lekar" + " " + user.ime + " " + user.prezime);
                           
                        }
                    }
                    lek.obrazlozenje = (neki + " " + imeprez);
                    excelRange.Cells[i, 12] = Convert.ToString(lek.obrazlozenje);
                }
            }
        }
        excelApp.ActiveWorkbook.Save();
        excelApp.Quit();
        System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
    }
    static void Cekajupotvrdu(List<Lek> arrLekovi)
    {
        foreach (Lek lek in arrLekovi)
        {
            if (lek.prihvacen == false && lek.odbijen == false)
            {
                Console.WriteLine(lek.ime + " " + lek.sifra + " " + lek.proizvodjac + " " + lek.cena + " \n ");
            }
        }
    }
    static void Prikazpio(List<Korisnik> arrUsers, List<Lek> arrLekovi)
    {
        Console.WriteLine("1. Prihvaceni \n" + "2. Odbijeni \n" + "3. Prihvaceni i odbijeni");
        int k = Convert.ToInt32(Console.ReadLine());
        foreach (Lek lek in arrLekovi)
        {
            if(k==1 && lek.prihvacen==true)
            {
                Console.WriteLine(lek.ime + " " + lek.sifra + " " + lek.proizvodjac + " " + lek.cena + " \n ");
            }
            if (k == 2 && lek.odbijen == true)
            {
                Console.WriteLine(lek.ime + " " + lek.sifra + " " + lek.proizvodjac + " " + lek.cena + " \n " + lek.obrazlozenje);
            }
            if (k == 3 && (lek.prihvacen == true||lek.odbijen==true))
            {
                Console.WriteLine(lek.ime + " " + lek.sifra + " " + lek.proizvodjac + " " + lek.cena + " \n " + lek.obrazlozenje);
            }
        }

    }
}



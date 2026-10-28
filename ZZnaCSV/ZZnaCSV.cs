using Hydra;
using System.Drawing;
using System.Windows.Forms;
using System.Diagnostics;
using ZZnaCSV;
using System.Collections.Generic;
using static ZZnaCSV.ZS;
using System.Text;
using System;
using System.Data.SqlClient;
using System.Data;
using cdn_api;
using static ZZnaCSV.Twr;

[assembly: CallbackAssemblyDescription(Name: "ONE-005 ZSnaFS",
                                           Description: "Tworzenie dokumentu FS na podstawie dokumentu ZS",
                                           Author: "OneSupport",
                                           Version: "1.0",
                                           SystemVersion: "2021.0.2",
                                           Created: "18-11-2021")]


namespace ZZnaCSV
{
    [SubscribeProcedure(Procedures.ZaNLista, "ZS na FS")]
    public class ZSnaFS_Class : Callback
    {

        private int XLWersjaAPI = 20202;
        private int XLTrybWsadowy = 1;
        private int XLWinieta = -1;
        private int XLUtworzWlasnaSesje = 1;
        private int sesja = -1;
        private string XLSQLServer = "192.168.1.202";
        private string XLSQLDatabase = "Tuflowers-Mob";//"ERPXL_PoloMarket"
        private string XLUser = "ADMIN"; //
        private string XLPass = "NE!w0rk12!";//"Dupa2003!";
        private ClaWindow RapButton;
        private ClaWindow zakladkaTabWewnetrzne;
        private ClaWindow filtrBtn;
        private List<ListaZS> ZSy = new List<ListaZS>();
        private List<ListaTwr> towary = new List<ListaTwr>();
        int listID;
        public override void Init()
        {

            AddSubscription(true, 0, Events.JustAfterWindowOpening, new TakeEventDelegate(OnOpenWindow));
            AddSubscription(false, 0, Events.ResizeWindow, new TakeEventDelegate(Resize));
        }



        private bool OnOpenWindow(Procedures ProcId, int ControlId, Events Event)
        {

            ClaWindow Parent = GetWindow();
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Rectangle WndBnd = GetWindow().Bounds;
            zakladkaTabWewnetrzne = GetWindow().AllChildren["?TabOdb"];
            RapButton = zakladkaTabWewnetrzne.Children.Add(ControlTypes.button);
            RapButton.Visible = true;
            //RapButton.TextRaw = "FS";
            RapButton.ToolTipRaw = "Przekształcenie indeksów na dedykowane";
            RapButton.IconRaw = "C:\\OneSupport\\36063.ico";

            filtrBtn = GetWindow().AllChildren["?AnalizaProduktu"];                            //pobiera współrzędne przycisku zapisz
            RapButton.Bounds = new Rectangle(filtrBtn.Bounds.X + 18, filtrBtn.Bounds.Y, 15, 15);
            listID = GetWindow().AllChildren["?ListaZam"].Id;      //tu sprawdzic

            RapButton.OnAfterMouseDown += new TakeEventDelegate(BtnOnMouseDown);

            //pobrane z pola w zakładce Raporty
            return true;
        }

        public bool Resize(Procedures ProcId, int ControlId, Events Event)
        {
            filtrBtn = GetWindow().AllChildren["?AnalizaProduktu"];                            //pobiera współrzędne przycisku zapisz
            RapButton.Bounds = new Rectangle(filtrBtn.Bounds.X + 18, filtrBtn.Bounds.Y, 15, 15);

            return true;
        }

        public bool BtnOnMouseDown(Procedures ProcId, int ControlId, Events Event)
        {

            try
            {

                ADODB._Recordset recordset = Runtime.WindowController.GetQueueMarked((int)ProcId, listID, GetCallbackThread());
                // if(recordset!=null)
                // MessageBox.Show(Convert.ToString(listID), "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                ZbieranieZS(recordset);

                if (Runtime.ConfigurationDictionary.Baza.ToString().Substring(0, 22) == "ERPXL_Tuflowers_Mobile")
                    XLSQLDatabase = "Tuflowers-Mob";
                else if (Runtime.ConfigurationDictionary.Baza.ToString().Substring(0, 24) == "ERPXL_Tuflowers_20220331")
                    XLSQLDatabase = "Tuflowers-TEST-NOWY";
                else if (Runtime.ConfigurationDictionary.Baza.ToString().Substring(0, 24) == "ERPXL_Tuflowers_20220908")
                    XLSQLDatabase = "Tuflowers_20220908";
                else if (Runtime.ConfigurationDictionary.Baza.ToString().Substring(0, 15) == "ERPXL_Tuflowers")
                    XLSQLDatabase = "Tuflowers";

               // MessageBox.Show("Baza SQL: "+ Runtime.ConfigurationDictionary.Baza.ToString().Substring(0, 24)+"\nBaza XL: "+XLSQLDatabase, "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                XLLoginInfo_20202 nowyUzytkownik = new XLLoginInfo_20202();
                nowyUzytkownik.Wersja = XLWersjaAPI;
                nowyUzytkownik.Winieta = XLWinieta;
                nowyUzytkownik.TrybWsadowy = XLTrybWsadowy;
                nowyUzytkownik.UtworzWlasnaSesje = XLUtworzWlasnaSesje;
                nowyUzytkownik.SerwerKlucza = XLSQLServer;
                nowyUzytkownik.Baza = XLSQLDatabase;
                nowyUzytkownik.OpeIdent = XLUser;
                nowyUzytkownik.OpeHaslo = XLPass;
                nowyUzytkownik.PortHTTPSerweraKlucza = 1;
                nowyUzytkownik.RodzajSerweraKlucza = 0;
                nowyUzytkownik.SerwerKlucza = "TF-SQL::5000185869";
                nowyUzytkownik.ProgramID = "ONE-005 ZSnaFS";

                int blad = cdn_api.cdn_api.XLLogin(nowyUzytkownik, ref sesja);

                if (blad != 0)
                {
                    MessageBox.Show("Błąd logowania: " + blad, "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

                proceduraZS();

                /* foreach(var item in listaTowars)
                 {
                     MessageBox.Show("GIDTyp "+Convert.ToString(item.GIDTyp), "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                     MessageBox.Show("GIDFirma " + Convert.ToString(item.GIDFirma), "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                     MessageBox.Show("GIDNumer " + Convert.ToString(item.GIDNumer), "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                     MessageBox.Show("GIDLp " + Convert.ToString(item.GIDLp), "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                 }*/
                cdn_api.cdn_api.XLLogout(sesja);

            }
            catch { }
            ZSy.Clear();

            return true;
        }

        bool ZbieranieZS(ADODB._Recordset recordset)
        {

            try
            {
                StringBuilder dokumenty = new StringBuilder();
                if (recordset != null)
                {
                    int count = recordset.RecordCount;

                    if (count != 0)
                    {
                        recordset.MoveFirst();

                        while (recordset.EOF == false)
                        {
                            ADODB.Fields fields = recordset.Fields;

                            ZSy.Add(new ListaZS()
                            {
                                GIDTyp = Convert.ToInt16(fields[0].Value),
                                GIDFirma = Convert.ToInt32(fields[1].Value),
                                GIDNumer = Convert.ToInt32(fields[2].Value),
                                GIDLp = Convert.ToInt16(fields[3].Value),
                            });
                            recordset.MoveNext();
                        }
                    }
                    else
                    {
                        MessageBox.Show("Nie zaznaczono dokumentów ZS. Operacja anulowana.", "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        return true;
                    }
                }
                else
                {
                    MessageBox.Show("Nie zaznaczono dokumentów ZS. Operacja anulowana.", "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return true;
                }
                return false;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Błąd pobierania zaznaczonych ZS --> " + ex.Message, "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return true;
            }
        }

        void proceduraZS()
        {
            string connectionString = "";
            if (Runtime.ConfigurationDictionary.Baza.ToString().Substring(0, 22) == "ERPXL_Tuflowers_Mobile")
                connectionString = @"Data Source = TF-SQL; Initial Catalog = ERPXL_Tuflowers_Mobile; User ID = onesupport; Password=Ancr3d0$$db";
            else if (Runtime.ConfigurationDictionary.Baza.ToString().Substring(0, 24) == "ERPXL_Tuflowers_20220331")
                connectionString = @"Data Source = TF-SQL; Initial Catalog = ERPXL_Tuflowers_20220331; User ID = onesupport; Password=Ancr3d0$$db";
            else if (Runtime.ConfigurationDictionary.Baza.ToString().Substring(0, 24) == "ERPXL_Tuflowers_20220908")
                connectionString = @"Data Source = TF-SQL; Initial Catalog = ERPXL_Tuflowers_20220908; User ID = onesupport; Password=Ancr3d0$$db";
            else if (Runtime.ConfigurationDictionary.Baza.ToString().Substring(0, 15) == "ERPXL_Tuflowers")
                connectionString = @"Data Source = TF-SQL; Initial Catalog = ERPXL_Tuflowers; User ID = onesupport; Password=Ancr3d0$$db";
            //connectionString = @"Data Source = TF-SQL; Initial Catalog = ERPXL_Tuflowers_Mobile;persist security info=True; Integrated Security=SSPI;";
            if (connectionString != "")
            {
                string kom = "";
                foreach (var item in ZSy)
                {
                    string komunikat = "";
                    // dodaje kolumny
                    procedura(connectionString, item.GIDTyp, item.GIDNumer, ref komunikat);
                    if (komunikat.Contains(" - OK | "))
                        modyfikacja(item.GIDTyp, item.GIDNumer, item.GIDFirma, item.GIDLp, connectionString);
                    towary.Clear();
                    kom += komunikat+"\n";
                   
                }
                MessageBox.Show(kom, "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        void procedura(string connectionString, int GIDTyp, int GIDNumer, ref string komunikat)
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(connectionString))
                {

                    conn.Open();

                    // 1.  create a command object identifying the stored procedure
                    SqlCommand cmd = new SqlCommand("cdn.ONE_PrzeksztalcIndeksyNaZS_Z", conn);


                    // 2. set the command object so it knows to execute a stored procedure
                    cmd.CommandType = CommandType.StoredProcedure;

                    // 3. add parameter to command, which will be passed to the stored procedure
                    cmd.Parameters.Add("@komunikat", SqlDbType.VarChar, 8000);
                    cmd.Parameters["@komunikat"].Direction = ParameterDirection.Output;
                    cmd.Parameters.Add(new SqlParameter("@gidtyp", GIDTyp));
                    cmd.Parameters.Add(new SqlParameter("@gidnumer", GIDNumer));
                    //cmd.Parameters.Add(new SqlParameter("@komunikat", komunikat));


                    // execute the command
                    using (SqlDataReader rdr = cmd.ExecuteReader())
                    {
                        // iterate through results, printing each to console
                        while (rdr.Read())
                        {

                            //Console.WriteLine("Product: {0,-35} Total: {1,2}", rdr["ProductName"], rdr["Total"]);
                            //MessageBox.Show("W procedurze: "+Convert.ToString(rdr["komunikat"]), "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                    komunikat = Convert.ToString(cmd.Parameters["@komunikat"].Value);
                    
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Złe wywołanie procedury: " + ex.Message, "Komunikat",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        void modyfikacja(int GIDTyp, int GIDNumer, int GIDFirma, short GIDLp, string connectionString)
        {
            string LpQuery = "SELECT [ZaE_GIDLp] FROM [CDN].[ZamElem] WHERE ZaE_GIDTyp="+GIDTyp+" and ZaE_GIDFirma="+GIDFirma+" and ZaE_GIDNumer="+GIDNumer+" ORDER BY [ZaE_GIDLp] DESC";
            List<string> lp = SelectList(LpQuery, "CDN.ZamElem", connectionString);

            XLOtwarcieDokumentuZamInfo_20202 zamowienie = new XLOtwarcieDokumentuZamInfo_20202();
            zamowienie.Wersja = XLWersjaAPI;
            zamowienie.ZaNFirma = GIDFirma;
            zamowienie.ZaNLp = GIDLp;
            zamowienie.ZaNNumer = GIDNumer;
            zamowienie.ZaNTyp = GIDTyp;
            zamowienie.Tryb = 2;
            int dokumentID = 0;
            int blad = cdn_api.cdn_api.XLOtworzDokumentZam(sesja, ref dokumentID, zamowienie);

            if (blad != 0)
            {
                MessageBox.Show("Błąd otwierania zamówienia: " + blad, "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

            foreach (string pozycja in lp)
            {
                /*string twrKodQuery = "SELECT TOP (1) [ZaE_TwrNumer] FROM [CDN].[ZamElem] WHERE ZaE_GIDTyp =" + GIDTyp + " and ZaE_GIDFirma = " + GIDFirma + " and ZaE_GIDNumer = " + GIDNumer + " and ZaE_GIDLp = " + pozycja;
                string twrKod = ConnectToSQL(connectionString, twrKodQuery);
                string IloscQuery = "SELECT TOP (1) [ZaE_Ilosc] FROM [CDN].[ZamElem] WHERE ZaE_GIDTyp =" + GIDTyp + " and ZaE_GIDFirma = " + GIDFirma + " and ZaE_GIDNumer = " + GIDNumer + " and ZaE_GIDLp = " + pozycja;
                string ilosc = ConnectToSQL(connectionString, IloscQuery);

                towary.Add(new ListaTwr()
                {
                    GIDTyp = 16,
                    GIDFirma = GIDFirma,
                    GIDNumer = Convert.ToInt32(twrKod),
                    GIDLp = 0,
                    Ilosc = ilosc,
                });*/

                //towary.Add()

                XLUsunZamElemInfo_20202 elem = new XLUsunZamElemInfo_20202();
                elem.Pozycja = Convert.ToInt32(pozycja);
                elem.Wersja = XLWersjaAPI;
                int usuniecieBlad=cdn_api.cdn_api.XLUsunPozycjeZam(dokumentID, elem);
                if (usuniecieBlad != 0)
                {
                    MessageBox.Show("Błąd usuwania pozycji zamówienia: " + usuniecieBlad, "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }


            }

            string twrQuery = "SELECT [TwrKod] ,[Ilosc] FROM [CDN].[ONE_TabPrzeksztalcIndeksyNaZS_Z] WHERE GIDTyp="+GIDTyp+" and GIDNumer="+GIDNumer;
            towary = SelectListTwr(twrQuery, "CDN.ONE_TabPrzeksztalcIndeksyNaZS_Z", connectionString);

            foreach (ListaTwr towar in towary)
            {
                XLDokumentZamElemInfo_20202 element = new XLDokumentZamElemInfo_20202();
                element.Ilosc = towar.Ilosc;
                /*element.TwrFirma = towar.GIDFirma;
                element.TwrLp = towar.GIDLp;
                element.TwrNumer = towar.GIDNumer;
                element.TwrTyp = towar.GIDTyp;*/
                element.Towar = towar.Kod;
                element.Wersja = XLWersjaAPI;
                int dodawanieBlad=cdn_api.cdn_api.XLDodajPozycjeZam(dokumentID, element);
                if (dodawanieBlad != 0)
                {
                    MessageBox.Show("Błąd dodawania pozycji zamówienia: " + dodawanieBlad, "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

                /*XLAtrybutInfo_20202 atrybut5 = new XLAtrybutInfo_20202();
                atrybut5.Wersja = XLWersjaAPI;
                atrybut5.GIDFirma = towar.GIDFirma;
                atrybut5.GIDLp = towar.GIDLp;
                atrybut5.GIDNumer = towar.GIDNumer;
                atrybut5.GIDSubLp = 0;
                atrybut5.GIDTyp = towar.GIDTyp;
                atrybut5.Wartosc = "FS";
                atrybut5.Klasa = "PA/FS";
                int blad_atr = cdn_api.cdn_api.XLDodajAtrybut(sesja, atrybut5);

                if (blad_atr != 0)
                {
                    MessageBox.Show("Błąd atrybutu: " + blad_atr, "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }*/
            }  

            XLZamkniecieDokumentuZamInfo_20202 dokZam = new XLZamkniecieDokumentuZamInfo_20202();
            dokZam.Wersja = XLWersjaAPI;
            //dokZam.TrybZamkniecia = 0;
            int zamkniecieBlad=cdn_api.cdn_api.XLZamknijDokumentZam(dokumentID,dokZam);
            if (zamkniecieBlad != 0)
            {
                MessageBox.Show("Błąd zamykania pracy z dokumentem: " + zamkniecieBlad, "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            

        }

        public string ConnectToSQL(string connection_string, string query)
        {
            string str = "";
            try
            {
                using (SqlConnection connection = new SqlConnection(connection_string))
                {
                    SqlDataAdapter adapter = new SqlDataAdapter();
                    adapter.SelectCommand = new SqlCommand(query, connection);
                    DataSet dataSet = new DataSet();
                    adapter.Fill(dataSet, "Query");
                    if (dataSet.Tables.Count == 0)
                    {
                        return "";
                    }
                    foreach (DataRow row in (InternalDataCollectionBase)dataSet.Tables[0].Rows)
                    {
                        str = row[0].ToString();
                    }
                }
            }
            catch (Exception ex)
            {
                str = str + "Bład SQL:" + ex.ToString() + "\n";
            }
            return str;
        }

        public List<string> SelectList(string query, string srcTableName, string connectionString)
        {
            List<string> str = new List<string>();
            string sttr = "";
            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    SqlDataAdapter adapter = new SqlDataAdapter();
                    adapter.SelectCommand = new SqlCommand(query, connection);
                    DataSet dataSet = new DataSet();
                    adapter.Fill(dataSet, srcTableName);
                    foreach (DataRow row in (InternalDataCollectionBase)dataSet.Tables[0].Rows)
                        str.Add(row[0].ToString());
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(sttr = sttr + "Bład SQL:" + ex.ToString() + "\n");
            }
            return str;
        }

        private List<ListaTwr> SelectListTwr(string query, string srcTableName, string connectionString)
        {
            List<ListaTwr> twr = new List<ListaTwr>();
            string sttr = "";
            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    SqlDataAdapter adapter = new SqlDataAdapter();
                    adapter.SelectCommand = new SqlCommand(query, connection);
                    DataSet dataSet = new DataSet();
                    adapter.Fill(dataSet, srcTableName);
                    foreach (DataRow row in (InternalDataCollectionBase)dataSet.Tables[0].Rows)
                    {
                        //twr.Add(row[0].ToString());
                        twr.Add(new ListaTwr()
                        {
                            
                            Kod = row[0].ToString(),
                            Ilosc = row[1].ToString()
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(sttr = sttr + "Bład SQL:" + ex.ToString() + "\n");
            }
            return twr;
        }
        public override void Cleanup()
        {

        }

    }
}

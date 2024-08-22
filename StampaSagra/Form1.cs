using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing.Printing;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Web;
using System.Drawing;
using System.Linq;

/* Programma per gestione stampa di PLS:
 * dati i parametri in ingresso, genera uno scontrino e lo stampa.
 * Richiede parametri di avvio (args[])
* 2 possibilità:
* 1) la ricevo 3 parametri (compreso il primo che è prendente predefinito)
    *      [0] inutile
    *      [1] ID_cliente
    *      [2] stampante_cassa
    *  in questo caso, la stampa va in cucina ed è la copia per il cliente
    * 2) ricevo 4 parametri
    *      [0] inutile
    *      [1] ID_cliente
    *      [2] stampante_cucina
    *      [4] stampate_bere
    *   in questo caso, la stampa avviene dopo l'inserimento della posizione
*/





namespace StampaSagra
{
    public partial class Form1 : Form
    {
        static int page_width = 300;
        static int logo_x = 5;
        static int logo_y = 30;
        static int logo_dim_x = 125;
        static int logo_dim_y = 100;
        static int barcode_x = 140;
        static int barcode_y = 40;
        static int barcode_dim_x = 157;
        static int barcode_dim_y = 70;
        static int left_margin = 5;
        static int first_column = 165;
        static int second_column = 225;
        static int max_description_lenght = 140;
        static int totale_x = 150;
        static int totale_x_numero = 210;
        static int copia_x = 5;
        static int copia_y = 100;
        static int copia_dim_x = 300;
        static int copia_dim_y = 260;
        static int footer_dim_x = 270;
        static int footer_dim_y = 110;
        static int footer_x = 10;



        public Form1()
        {
            InitializeComponent();
        }

        //caricamento del form
        private void Form1_Load(object sender, EventArgs e)
        {
            //connessione al db, la string adi connessione + ben nascosata nelle risorse
#if DEBUG
            string sql_connec_string = StampaSagra.Properties.Resources.ConnectionStringRemote;
#else
            string sql_connec_string = StampaSagra.Properties.Resources.ConnectionStringLocal;
#endif
            SqlConnection sql_connection = connectToSql(sql_connec_string);
            //Legge gli args
            string[] args = Environment.GetCommandLineArgs();
            int ID_cliente = 0;
            int ordineInterno = 0;
            String stampante = "";
            String categorie = "";
            /*OLD*/
            // 2 possibilità
            //Args[0] path dell'eseguibile
            //[1] id_cliente
            //[2] stampante_cassa
            //Args[0] path dell'eseguibile
            //[1] id_cliente
            //[2] stampante_cucina
            //[3] stampante_bere
            /*END OLD*/

            /*NEW
             * 2 possibilità:
             * [0] path dell'eseguibile (non serve, la genera lui!)
             * [1] id cliente
             * [2]      0 per stampare la copia cliente (con sponsor e con copia cliente esposto); 
             *          1 per stampare come copia interna (senza sponsor e con "copia" per la cucina);
             * [3] categoria:
             *          immettere la categoria separata da virgola
             *          1,2,3,6,9
             * [4] stampante:
             *          al solito, il nome della stampante.
             * 
             * */
            
            // Nel caso passo solo una stampante, è il caso della stampa in cassa
            if (args.Length == 5)
            {
                
                try
                {
                    ID_cliente = Int32.Parse(args[1]);
                    ordineInterno = Int32.Parse(args[2]);
                    categorie = expandCetegorie(args[3]);
                    stampante = args[4];
#if DEBUG
                    //per debug
                    stampante = "Microsoft Print to PDF";
#endif
                    // se l'ordine è interno, una procedura di stampa
                    if (ordineInterno == 1)
                    {
                        stampaInterno(getDatiOrdine(ID_cliente, sql_connection, categorie), stampante, ID_cliente, getInfoCliente(ID_cliente, sql_connection), categorie);
                    }
                    else if (ordineInterno == 0)
                    {
                        // se ometto le categorie, le xonsidera tutte
                        stampaEsterno(getDatiOrdine(ID_cliente, sql_connection), stampante, ID_cliente, getInfoCliente(ID_cliente, sql_connection));
                    }

                }
                catch (Exception ex)
                {
                    System.Windows.Forms.Application.Exit();
                    MessageBox.Show(ex.Message, "Errore Nel parsing dei parametri (caso cassa)", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }


            }
            else
            {
                MessageBox.Show("Parametri non validi", "Errore parametri", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            System.Windows.Forms.Application.Exit();
        }

    
        private string expandCetegorie(String raw_string)
        {
            String result = "(";
            // parsa le categorie decentemente
            int i = 0;
            foreach (char s in raw_string)
            {
                //finché non trova il separatore, continua
                if (s == ',')
                {
                    result += ",";
                }
                else
                {
                    result += raw_string[i];
                }
                i++;
            }
            result += ")";
            return result;
        }

            void stampaInterno(List<List<String>> row_query_data, String stampante,  int ID_cliente,  List<String> dati_cliente, String categorie)
        {
            // Per gli ordini interni, senza footer e balle varie, la copia va in cucina.
            try
            {
                //Formatto decentementi i dati
                String nome_cliente = dati_cliente[0];
                int coperti = int.Parse(dati_cliente[1]);
                String posizione = dati_cliente[2];
                String oraPagato = dati_cliente[3];
                String oraIngressoCucina = dati_cliente[4];

                // importo il logo
                Bitmap logo_sagra = new Bitmap(StampaSagra.Properties.Resources.ComitatoFrazionale_PonteNelleAlpi, logo_dim_x, logo_dim_y);
                

                //Setto la lunghezza della pagina in funzione della lunghezza dell'ordine
                //int page_width = 314; //centesimi di inc
                // dinamico
                int page_height = 320 + 15 * row_query_data.Count();

                //Creo il documento
                PrintDocument pd = new PrintDocument();

                //Creo la dimensioni della carta
                PaperSize ps = new PaperSize("Scontrino", page_width, page_height);
                pd.DefaultPageSettings.PaperSize = ps;

                // aggiungo l'evento di stampa
                pd.PrintPage += new PrintPageEventHandler(pd_PrintPage);
                void pd_PrintPage(object sender, PrintPageEventArgs e)
                {
                    Graphics g = e.Graphics;
                    // disegno il logo
                    g.DrawImage(logo_sagra, logo_x, logo_y);
                    

                    //creo il codice a barre (mantenendo le proporzioni)
                    // il flag true/false serve a scrivere o meno l'id cliente sotto il codice
                    Image imgBarcode = BarcodeLib.Barcode.DoEncode(BarcodeLib.TYPE.CODE128A, ID_cliente.ToString(), true, Color.Black, Color.White, barcode_dim_x, barcode_dim_y);
                    g.DrawImage(imgBarcode, barcode_x, barcode_y);

                    //variabile di spazio, fin dove sono arrivato a scriver
                    int space = 155;

                    //creo il penel
                    SolidBrush sb = new SolidBrush(Color.Black);
                
                    // se c'è qualcosa da stampare
                    if (row_query_data.Count() != 0)
                        {
                        intestazione (g, new Font("Lucida Console", 8, FontStyle.Regular), sb, space);
                        space += 20;

                        // stampo la lista e totale

                        space = stampaArticoli(g, new Font("Lucida Console", 8, FontStyle.Regular), sb, space, row_query_data, e);
                        space += 20;
                    }

                    // stamo le informazioni del cliente
                    space = stampaInformazioniCliente(g, new Font("Lucida Console", 11, FontStyle.Bold), sb, space, nome_cliente, posizione, coperti, oraPagato, oraIngressoCucina);
                    space += 5;

                    // stampo il footer
                    stampaFooterAccoglienza(g, new Font("Lucida Console", 5, FontStyle.Italic), sb, space, categorie);
                   
                    g.Dispose();

                    pd.PrintController = new StandardPrintController();

                    pd.DefaultPageSettings.Margins.Left = 0;
                    pd.DefaultPageSettings.Margins.Right = 0;
                    pd.DefaultPageSettings.Margins.Top = 0;
                    pd.DefaultPageSettings.Margins.Bottom = 0;

                }
                print_execution(pd, stampante);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Errore nella fase di stampa", MessageBoxButtons.OK, MessageBoxIcon.Error);
                System.Windows.Forms.Application.Exit();
            }

        }

        void stampaEsterno(List<List<String>> row_query_data, String stampante, int ID_cliente, List<String> dati_cliente)
        {
            //metodo che viene chiamato quando deve stampare la cassa
            try
            {
                //Formatto decentementi i dati
                String nome_cliente = dati_cliente[0];
                int coperti = int.Parse(dati_cliente[1]);
                String posizione = dati_cliente[2];
                String oraPagato = dati_cliente[3];
                bool asporto = false;

                if (posizione == "ASPORTO")
                {
                    asporto = true;
                }



                // importo il logo
                Bitmap logo_sagra = new Bitmap(StampaSagra.Properties.Resources.ComitatoFrazionale_PonteNelleAlpi, logo_dim_x, logo_dim_y); 
                Bitmap copia = new Bitmap(StampaSagra.Properties.Resources.copia, copia_dim_x, copia_dim_y);
                Bitmap footer = new Bitmap(StampaSagra.Properties.Resources.logoarstudio, footer_dim_x, footer_dim_y);
                //Setto la lunghezza della pagina in funzione della lunghezza dell'ordine
                int page_height = 430 + 14 * row_query_data.Count();
                //Creo il documento
                PrintDocument pd = new PrintDocument();
                //Creo la dimensioni della carta
                PaperSize ps = new PaperSize("Scontrino", page_width, page_height);
                pd.DefaultPageSettings.PaperSize = ps;
                // aggiungo l'evento di stampa
                pd.PrintPage += new PrintPageEventHandler(pd_PrintPage);
                void pd_PrintPage(object sender, PrintPageEventArgs e)
                {
                    //Stampo il logo
                    Graphics g = e.Graphics;
                    // disegno il logo
                    g.DrawImage(logo_sagra, logo_x, logo_y);
                    g.DrawImage(copia, copia_x, copia_y);
                    //creo il codice a barre (mantenendo le proporzioni)
                    // il flag true/false serve a scrivere o meno l'id cliente sotto il codice
                    Image imgBarcode = BarcodeLib.Barcode.DoEncode(BarcodeLib.TYPE.CODE128A, ID_cliente.ToString(), true, Color.Black, Color.White, barcode_dim_x, barcode_dim_y);
                    g.DrawImage(imgBarcode, barcode_x, barcode_y);
                    int space = 155;
                    // credo i font
                    SolidBrush sb = new SolidBrush(Color.Black);
                    // minimo di intenstazione
                    intestazione(g, new Font("Lucida Console", 8, FontStyle.Regular), sb, space);
                    space += 20;
                    space = stampaArticoli(g, new Font("Lucida Console", 8, FontStyle.Regular), sb, space, row_query_data, e);
                    space += 20;

                    space = stampaInformazioniCliente(g, new Font("Lucida Console", 8, FontStyle.Regular), sb, space, nome_cliente, posizione, coperti, oraPagato);


                    space = stampaFooterCassa(g, new Font("Lucida Console", 5, FontStyle.Italic), sb, space, containsElementiCongelati(row_query_data), asporto);
                    g.DrawImage(footer, footer_x, space);

                    g.Dispose();

                    pd.PrintController = new StandardPrintController();

                    pd.DefaultPageSettings.Margins.Left = 0;
                    pd.DefaultPageSettings.Margins.Right = 0;
                    pd.DefaultPageSettings.Margins.Top = 0;
                    pd.DefaultPageSettings.Margins.Bottom = 0;


                }
                print_execution(pd, stampante);

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Errore nella fase di stampa", MessageBoxButtons.OK, MessageBoxIcon.Error);
                System.Windows.Forms.Application.Exit();
            }

        }

        private void stampaFooterAccoglienza(Graphics g, Font f, SolidBrush sb, int space, string categorie)
        {
            // stampa se la copia è del bar cucina o della cucina (na frocioneria insomma).
                space += 12;
                g.DrawString("Copia interna, categorie: "+ categorie, f, sb, left_margin, space);

        }

        private int stampaInformazioniCliente(Graphics g, Font f, SolidBrush sb, int space, String nome_cliente, String posizione, int coperti, String oraPagato, String oraIngressoCucina)
        {
            // stampta le informazioni del cliente
            // setto l'offset in funzione del carattere
            int offset = f.Height + 3;
            g.DrawString("Cliente: " + FirstCharToUpper(nome_cliente), f, sb, left_margin, space);
            //Se la posizione non è assegnata
            if (posizione != "0")
            {
                space += offset;
                g.DrawString("Posizione: " + posizione, f, sb, left_margin, space);
            }
            space += offset;
            g.DrawString("Coperti: " + coperti.ToString(), f, sb, left_margin, space);
            space += offset;
            g.DrawString("h pag:   " + oraPagato, f, sb, left_margin, space);
            //space += offset;
            //g.DrawString("h accog: " + oraIngressoCucina, f, sb, left_margin, space);
            return space;
        }

        private int stampaInformazioniCliente(Graphics g, Font f, SolidBrush sb, int space, String nome_cliente, String posizione, int coperti, String oraPagato)
        {
            // stampa le informazioni del cliente
            g.DrawString("Cliente: " + FirstCharToUpper(nome_cliente), f, sb, left_margin, space);
            //Se la posizione non è assegnata
            if (posizione != "0")
            {
                space += 10;
                g.DrawString("Posizione: " + posizione, f, sb, left_margin, space);
            }
            space += 10;
            g.DrawString("Coperti: " + coperti.ToString(), f, sb, left_margin, space);
            space += 10;
            g.DrawString("Ora pagamento: " + oraPagato, f, sb, left_margin, space);
            return space;
        }

        private int stampaArticoli(Graphics g, Font f, SolidBrush sb, int space, List<List<String>> row_query_data, PrintPageEventArgs e)
        {
            // stampa la tabella con gli articoli
            double totale = 0;
            foreach (List<String> righe in row_query_data)
            {
                SizeF lunghezza_descrizione = e.Graphics.MeasureString(righe[0], f);
                //creo un box, così va a capo in automatico la descrizizone che è il più critico
                // Stamp il nome della pietanza
                g.DrawString(righe[0], f, sb, new Rectangle(left_margin, space, max_description_lenght, (((int)Math.Ceiling(lunghezza_descrizione.Width)) / max_description_lenght + 1)*11));
                // Stampa la quantità
                g.DrawString(righe[1], f, sb, first_column, space);
                // parsing strani per formattare a dovere
                // Stampa il prezzo
                g.DrawString((Double.Parse(righe[2])).ToString("C", System.Globalization.CultureInfo.CreateSpecificCulture("it-IT")), f, sb, second_column, space);
                totale += Double.Parse(righe[2]);
                //Se è più lungo di tot allora devo saltare una riga
                if (lunghezza_descrizione.Width > 130)
                {

                    int righe_da_aggiungere = ((int)Math.Ceiling(lunghezza_descrizione.Width)) / max_description_lenght + 1;
                    space += 11 * righe_da_aggiungere;
                }
                else
                {
                    space += 11;
                }
            }
            space += 5;
            g.DrawString("Totale", new Font("Lucida Console", 10, FontStyle.Regular), sb, totale_x, space);
            g.DrawString(totale.ToString("C", System.Globalization.CultureInfo.CreateSpecificCulture("it-IT")), new Font("Lucida Console", 10, FontStyle.Regular), sb, totale_x_numero, space);
            return space;
        }

        void intestazione(Graphics g, Font f, SolidBrush sb, int space)
        {
            // intestazione degli ordini
            g.DrawString("Articolo", f, sb, left_margin, space);
            g.DrawString("QT", f, sb, first_column, space);
            g.DrawString("Prezzo", f, sb, second_column, space);
        }

        private bool containsElementiCongelati (List<List<String>> row_query_data)
        {
            // Mi dice se ci sono prodotti surgelati all'interno dell'ordine
            foreach (List<String> row in row_query_data)
            {
                // se la descrizione contiene "*"
                if (row[0].Contains("*"))
                {
                    return true;
                }
            }
            return false;
        }

        private int stampaFooterCassa(Graphics g, Font font, SolidBrush sb, int space, bool elementi_congelati, bool asporto)
        {
            //Stampa le ultime due cagate in cursico per la cassa
            if (elementi_congelati)
            {
                space += 12;
                g.DrawString("* Prodotto surgelato --- scontrino non fiscale", font, sb, left_margin, space);
            }
            else
            {
                space += 12;
                g.DrawString("Scontrino non fiscale", font, sb, left_margin, space);
            }
            space += 12;
            //stampa l'accoglienza
            /*if (!asporto)
            {
                g.DrawString("Magna che se sfreda.", new Font("Lucida Console", 12, FontStyle.Bold), sb, left_margin + 20, space);
                space += 20;
            }*/
            g.DrawString("Gestionale Polpet la Sagra. Nicola Pison, Enrico Pierobon.", font, sb, left_margin, space);
            space += 20;

            return space;
        }

 

        private string FirstCharToUpper(string input)
        {
            // Alza la prima lettera del carattere (per far bella figura con i clienti, sempre)
            if (String.IsNullOrEmpty(input))
                throw new ArgumentException("ARGH!");
            return input.First().ToString().ToUpper() + input.Substring(1);
        }

        public void killMyself()
        {
            System.Windows.Forms.Application.Exit();
        }


        List<List<String>> getDatiOrdine(int ID_cliente, SqlConnection conn)
        {
            // Ritorna l'ordine effettuao dal cliente, nel mangiare
            // richiede: 
            // ID_cliente ovvero l'id del cliente in questione
            // SqlConnection ovvero la connessione in sql al db
            // Ritorna:
            // descrizione quantità ed prezzo
            // lista di liste
            List<List<String>> row_query_data = new List<List<string>>();
            try
            {
                /* preparo la query */
                // FIXME: Probabile baco, formattazione della categoria
                SqlCommand query_command = new SqlCommand(
                  @"
                    SELECT Articoli.descrizione, count(Articoli.ID_articolo) as 'Quantità', Sum(Articoli.prezzo) as 'Prezzo'
                        FROM dbo.Cliente INNER JOIN dbo.Ordine ON Cliente.ID_cliente = Ordine.ID_cliente 
                            INNER JOIN dbo.Articoli ON Articoli.ID_articolo = Ordine.ID_articolo 
                            INNER JOIN dbo.DettagliCliente ON Cliente.ID_cliente = DettagliCliente.ID_cliente
                        WHERE Ordine.ID_cliente = @param1 
                        Group by Articoli.descrizione, Articoli.prezzo, Articoli.ID_categoria
                        Order by Articoli.ID_categoria, Articoli.descrizione",
                  conn);
                // gioco con i parametri (dovrbbe esser più sicuro)
                SqlParameter ID_cliente_param = new SqlParameter("@param1", SqlDbType.Int, 4);
                ID_cliente_param.Value = ID_cliente;
                query_command.Parameters.Add(ID_cliente_param);
                // mi connetto
                conn.Open();
                SqlDataReader reader = query_command.ExecuteReader();
                while (reader.Read())
                {
                    //lista temporanea per estrapolare le righe
                    List<String> riga = new List<string>();
                    for (int i = 0; i < reader.VisibleFieldCount; i++)
                    {
                        //giro sulle righe
                        String temp_string = reader.GetSqlValue(i).ToString();
                        //decodifico l'html (ho caratteri stronzi), così li sistemo
                        temp_string = HttpUtility.HtmlDecode(temp_string);
                        // scazzo per il carattere €
                        if (temp_string.Contains("â‚¬"))
                        {
                            temp_string = temp_string.Replace("â‚¬", "€");
                        }
                        riga.Add(temp_string);
                    }
                    // aggiungo le righe alla lista principale
                    row_query_data.Add(riga);
                }
                conn.Close();

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "Errore nella query di lettura", MessageBoxButtons.OK, MessageBoxIcon.Error);
                System.Windows.Forms.Application.Exit();
            }
            return row_query_data;
        }

        List<List<String>> getDatiOrdine(int ID_cliente, SqlConnection conn, String categorie)
        {
            // Ritorna l'ordine effettuao dal cliente, nel mangiare
            // richiede: 
            // ID_cliente ovvero l'id del cliente in questione
            // SqlConnection ovvero la connessione in sql al db
            // Ritorna:
            // descrizione quantità ed prezzo
            // lista di liste
            List<List<String>> row_query_data = new List<List<string>>();
            try
            {
                /* preparo la query */
                // FIXME: Probabile baco, formattazione della categoria
                SqlCommand query_command = new SqlCommand(
                  @"
                    SELECT Articoli.descrizione, count(Articoli.ID_articolo) as 'Quantità', Sum(Articoli.prezzo) as 'Prezzo'
                        FROM dbo.Cliente INNER JOIN dbo.Ordine ON Cliente.ID_cliente = Ordine.ID_cliente 
                            INNER JOIN dbo.Articoli ON Articoli.ID_articolo = Ordine.ID_articolo 
                            INNER JOIN dbo.DettagliCliente ON Cliente.ID_cliente = DettagliCliente.ID_cliente
                        WHERE Ordine.ID_cliente = @param1 AND Articoli.ID_categoria in "+ categorie + @"
                        Group by Articoli.descrizione, Articoli.prezzo, Articoli.ID_categoria
                        Order by Articoli.ID_categoria, Articoli.descrizione",
                  conn);
                // gioco con i parametri (dovrbbe esser più sicuro)
                SqlParameter ID_cliente_param = new SqlParameter("@param1", SqlDbType.Int, 4);
                ID_cliente_param.Value = ID_cliente;
                query_command.Parameters.Add(ID_cliente_param);
                // mi connetto
                conn.Open();
                //per debug
                SqlDataReader reader = query_command.ExecuteReader();
                while (reader.Read())
                {
                    //lista temporanea per estrapolare le righe
                    List<String> riga = new List<string>();
                    for (int i = 0; i < reader.VisibleFieldCount; i++)
                    {
                        //giro sulle righe
                        String temp_string = reader.GetSqlValue(i).ToString();
                        //decodifico l'html (ho caratteri stronzi), così li sistemo
                        temp_string = HttpUtility.HtmlDecode(temp_string);
                        // scazzo per il carattere €
                        if (temp_string.Contains("â‚¬"))
                        {
                            temp_string = temp_string.Replace("â‚¬", "€");
                        }
                        riga.Add(temp_string);
                    }
                    // aggiungo le righe alla lista principale
                    row_query_data.Add(riga);
                }
                conn.Close();

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "Errore nella query di lettura", MessageBoxButtons.OK, MessageBoxIcon.Error);
                System.Windows.Forms.Application.Exit();
            }
            return row_query_data;
        }

        List<String> getInfoCliente(int ID_cliente, SqlConnection conn)
        {
            // Ritorna nome, coperti, porizione, ora_pagato ed ora ingesso cucina dato l'id del cliente
            List<String> dati_cliente = new List<string>();
            try
            {
                /* preparo la query */
                SqlCommand query_command = new SqlCommand(
                   @"SELECT Cliente.nome, Cliente.coperti, dbo.DettagliCliente.posizione, DettagliCliente.OraPagato, DettagliCliente.OraIngressoCucina
                        FROM dbo.Cliente INNER JOIN dbo.Ordine ON Cliente.ID_cliente = Ordine.ID_cliente 
                            INNER JOIN dbo.DettagliCliente ON Cliente.ID_cliente = DettagliCliente.ID_cliente
						WHERE Ordine.ID_cliente = @param1 
                        Group by Cliente.nome, Cliente.coperti, dbo.DettagliCliente.posizione, DettagliCliente.OraPagato, DettagliCliente.OraIngressoCucina
                                ", conn);
                // gioco con i parametri (dovrbbe esser più sicuro)
                SqlParameter ID_cliente_param = new SqlParameter("@param1", SqlDbType.Int, 4);
                ID_cliente_param.Value = ID_cliente;
                query_command.Parameters.Add(ID_cliente_param);
                // mi connetto
                conn.Open();
                SqlDataReader reader = query_command.ExecuteReader();
                while (reader.Read())
                {
                    for (int i = 0; i < reader.VisibleFieldCount; i++)
                    {
                        //giro sulle righe
                        String temp_string = reader.GetSqlValue(i).ToString();
                        //decodifico l'html (ho caratteri stronzi), così li sistemo
                        temp_string = HttpUtility.HtmlDecode(temp_string);
                        dati_cliente.Add(temp_string);
                    }
                }
                conn.Close();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "Errore nella query di lettura", MessageBoxButtons.OK, MessageBoxIcon.Error);
                System.Windows.Forms.Application.Exit();
            }
            if (dati_cliente.Count != 0)
            {
                if (dati_cliente[4] == "1900-01-01 00:00:00.000")
                {
                    dati_cliente[4] = "0";
                }
                return dati_cliente;
            }
            else
            {
                MessageBox.Show("Errore: ID cliente non trovato", "Errore dati cliente", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.Close();
                return dati_cliente;
            }
        }

        SqlConnection connectToSql(string sql_connection_string)
        {
            //Connessione al DB data la stringa
            try
            {
                return new SqlConnection(sql_connection_string);
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "Errore connessione al Database", MessageBoxButtons.OK, MessageBoxIcon.Error);
                System.Windows.Forms.Application.Exit();
                return null;
            }
        }

        void print_execution(PrintDocument pd, String stampante)
        {
#if DEBUG
            //Esegue l'azione di stampa (utile per debug)
            // Per debug
            PrintPreviewDialog printPreviewDialog1 = new PrintPreviewDialog();
            pd.PrinterSettings.PrinterName = stampante;
            printPreviewDialog1.Document = pd;
            printPreviewDialog1.ShowDialog();
#else

            // Per relase
            pd.PrinterSettings.PrinterName = stampante;
            pd.Print();
#endif
        }


    }
}

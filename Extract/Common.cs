using System;
using System.Data.OleDb;
using System.Windows.Forms;
using System.Collections;
using System.Collections.Generic;

namespace Extract
{

    public class Common
    {

        #region "Divers"

        public static void LienOuvre(String LeLien)
        {
            System.Diagnostics.Process.Start(LeLien);
        }

        public static void DocOuvre(string LeDoc)
        {
            System.Diagnostics.Process.Start(LeDoc);
        }

        public static DateTime DateMin(DateTime a, DateTime b)
        {
            if (a < b) { return a; } else { return b; }

        }

        public static DateTime DateMax(DateTime a, DateTime b)
        {
            if (a > b) { return a; } else { return b; }

        }

        public static object Nz(object O, object ValeurNull)
        {
            return (O == null) ? ValeurNull : O;
        }


        public static string Num2Sql(string T) {
            try
            {
                double B;
                double.TryParse(T.Replace(",", "."), out B);
                return B.ToString().Replace(",", ".");
            }
            catch
            {
                return "0";
            }
        }

        public static string Num2Txt(double B) {
            return B.ToString().Replace(".", ",");
        }

        public static decimal Txt2Num(string T)
        {
            try
            {
                decimal D;
                decimal.TryParse(T.Replace(",", "."), out D);
                return D;
            }
            catch
            {
                return 0;
            }
        }

        //TODO :
        //public static decimal Sql2Num(string T)
        //{
        //    decimal a;
        //    a = Nz(T,0)
        //    return a;
        //}

        public static string Txt2Sql(string s)
        {
            return s.Replace("'", "''");
        }

        public static string Date2Sql(DateTime D)
        {
            if (D == null) { return "NULL"; } else { return "'" + D.ToString("yyyy-MM-dd") + "'"; }
        }


        public static string Date2XL(DateTime D)
        {
            string s;
            if (D == null) { return ""; } else { return D.ToString("dd/MM/yyyy"); }
        }

        public static DateTime FindeMois(DateTime D)
        {
            return D.AddDays(-D.Day + 1).AddMonths(1).AddDays(-1);
        }

        //    Function repNom(s As String) As String
        //        If s.Substring(s.Length - 1, 1) <> "\" Then s &= "\"
        //        Return s
        //    End Function

        #endregion

        #region "SQL"

        public static void ConnexionFerme(ref OleDbConnection LaConnect) {
            try { if (LaConnect.State == System.Data.ConnectionState.Open) { LaConnect.Close(); } }
            catch {
                throw new Exception("Erreur fermeture connexion");
            }
        }

        public static bool ConnexionInit(string strCon, ref OleDbConnection LaConnect)
        {
            try
            {
                ConnexionFerme(ref LaConnect);
                LaConnect.ConnectionString = strCon;
                return true;
            }
            catch (Exception Ex) {
                MessageBox.Show(Ex.Message);
                return false;
                //throw new Exception("Erreur d'initialisation de connexion");
            }
        }

        public static bool ConnexionTest(string strCon)
        {
            try
            {
                using (OleDbConnection consql = new OleDbConnection(strCon))
                {
                    if (consql.State == System.Data.ConnectionState.Open) { consql.Close(); }
                    consql.Open();
                    consql.Close();
                    return true;
                }
            }
            catch { return false; }
        }

        public static bool SqlDo(string s, ref OleDbConnection LaConnect)
        {
            System.Data.OleDb.OleDbCommand lareq = new OleDbCommand();
            try
            {
                if (LaConnect.State != System.Data.ConnectionState.Open) { LaConnect.Open(); }
                lareq.CommandText = s;
                lareq.Connection = LaConnect;
                lareq.CommandType = System.Data.CommandType.Text;
                lareq.ExecuteNonQuery();
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return false;
            }
        }

        public static OleDbDataReader SqlLit(string s, ref OleDbConnection LaConnect)
        {
            System.Data.OleDb.OleDbCommand lareq = new OleDbCommand();
            try
            {
                if (LaConnect.State != System.Data.ConnectionState.Open) { LaConnect.Open(); }
                lareq.CommandText = s;
                lareq.Connection = LaConnect;
                lareq.CommandType = System.Data.CommandType.Text;
                return lareq.ExecuteReader();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return null;
            }

        }
        #endregion

        #region "Liste Combo"

        public class ListItem {
            public object Val;
            public string Txt;

            public ListItem(object NewValue, string NewText) {
                Val = NewValue;
                Txt = NewText;
            }

            public override string ToString()   {  return Txt;  }
        }


        public static void ListFill(string SSql, ListBox aList, ref OleDbConnection ConSql)
        {
            OleDbDataReader LeRs;
            ArrayList LaListe = new ArrayList();

            aList.DataSource = null;
            aList.Items.Clear();
            LeRs = SqlLit(SSql, ref ConSql);
            while (LeRs.Read()) {
                LaListe.Add(new ListItem(LeRs.GetInt32(0), LeRs.GetString(1)));
            }
            LeRs.Close();
            aList.DisplayMember = "Txt";
            aList.ValueMember = "Val";
            aList.DataSource = LaListe;
        }

        public static void ComboFill(string SSql, ComboBox aList, ref OleDbConnection ConSql)
        {
            OleDbDataReader LeRs;
            ArrayList LaListe = new ArrayList();

            aList.DataSource = null;
            aList.Items.Clear();
            LeRs = SqlLit(SSql, ref ConSql);
            while (LeRs.Read())
            {
                LaListe.Add(new ListItem(LeRs.GetInt32(0), LeRs.GetString(1)));
            }
            LeRs.Close();
            aList.DisplayMember = "Txt";
            aList.ValueMember = "Val";
            aList.DataSource = LaListe;
        }

        public static void ComboSetVal(string LaValeur, ComboBox aCombo)
        {
            foreach (ListItem l in (aCombo.Items))
            {
                if (l.Val.ToString() == LaValeur) { aCombo.SelectedItem = l; }
            }
        }

        public static string ComboGetVal(ComboBox aCombo)
        {
            return ((ListItem)(aCombo.SelectedItem)).Val.ToString();
        }


        public static void ListSetVal(string LaValeur, ListBox aList)
        {
            foreach (ListItem l in (aList.Items))
            {
                if (l.Val.ToString() == LaValeur) { aList.SelectedItem = l; }
            }
        }

        public static string ListGetVal(ListBox aList)
        {
            return ((ListItem)(aList.SelectedItem)).Val.ToString();
        }


        #endregion


        #region "Formulaire"

        public class SQLchamp
        {
            public string champNom;
            public string champVal;
            public string champTyp;

            public SQLchamp(string LeChamp, string LaValeur, string LeType)
            {
                champNom = LeChamp;
                champVal = LaValeur;
                champTyp = LeType;
            }

        }

        public static void FormVide(Control b)
        {
            foreach (Control c in (b.Controls)) {

                if ((string)c.Tag != "") {

                    try
                    {
                        if (c.GetType().ToString().Contains("TextBox")) { c.Text = ""; }

                        if (c.GetType().ToString().Contains("ComboBox")) {
                            ((ComboBox)c).SelectedIndex = -1;
                        }
                        
                        if (c.GetType().ToString().Contains("Label")) { c.Text = ""; }

                        if (c.GetType().ToString().Contains("GroupBox")) { FormVide(c); }
                    } catch (Exception ex) {
                        MessageBox.Show(ex.Message);
                    }
                }

            }
        }


        public static void FormRempli(Control b, string sSql, ref OleDbConnection consql)
        {
            OleDbDataReader LeRs;
            LeRs = SqlLit(sSql, ref consql);
            while (LeRs.Read())
            {
                foreach (Control c in (b.Controls))
                {
                    if ((string)(Nz(c.Tag, "")) != "")
                    {
                        try
                        {
                            string LaValeur = "";
                            if (LeRs.GetFieldType(LeRs.GetOrdinal(c.Tag.ToString().Split(',')[0])).Name.Contains("Int32")) { LaValeur = LeRs.GetInt32(LeRs.GetOrdinal(c.Tag.ToString().Split(',')[0])).ToString(); }
                            if (LeRs.GetFieldType(LeRs.GetOrdinal(c.Tag.ToString().Split(',')[0])).Name.Contains("String")) { LaValeur = LeRs.GetString(LeRs.GetOrdinal(c.Tag.ToString().Split(',')[0])); }


                            if (c.GetType().ToString().Contains("TextBox")) { c.Text = LaValeur; }

                            if (c.GetType().ToString().Contains("Label")) { c.Text = LaValeur; }

                            if (c.GetType().ToString().Contains("CheckBox")) { ((CheckBox)c).Checked = (bool)Nz(LeRs.GetBoolean(LeRs.GetOrdinal(c.Tag.ToString().Split(',')[0])), false); }

                            if (c.GetType().ToString().Contains("ComboBox")) {
                                //LeRs.GetFieldType(3).Name
                                ComboSetVal(LaValeur, (ComboBox)c);
                            }

                        } catch (Exception ex)
                        {
                            MessageBox.Show(c.Name + "-" + ex.Message);
                        }

                    }

                    //TODO: date
                    //                    If TypeOf(c) Is Windows.Forms.DateTimePicker Then
                    //                       If Nz(lers(c.tag), "") = "" Then
                    //                           c.value = CDate("31/12/2100")
                    //                            If c.showcheckbox Then c.checked = False
                    //                        Else
                    //                            c.value = lers(c.tag)
                    //                            c.enabled = True
                    //                            If c.showcheckbox Then c.checked = True
                    //                        End If
                    //                    End If



                    //            End If
                }
            }
            LeRs.Close();
            //    lers.Close()

        }

        public static List<SQLchamp> FormRecupereChamp(Control B)
        {
            List<SQLchamp> LesChamps = new List<SQLchamp>();

            foreach (Control c in (B.Controls))
            {
                if ((string)(Nz(c.Tag, "")) != "")
                {
                    try {
                        string[] leTag = c.Tag.ToString().Split(',');

                        if (c.GetType().ToString().Contains("TextBox") || c.GetType().ToString().Contains("Label") ) {
                            if (leTag.Length > 1) {
                                if (leTag[1].ToUpper().Contains("N")) {
                                    LesChamps.Add(new SQLchamp(leTag[0], "'" + Num2Sql(c.Text) + "'", leTag[1]));
                                } else
                                {
                                    LesChamps.Add(new SQLchamp(leTag[0], "'" + Txt2Sql(c.Text) + "'", leTag[1]));
                                }
                            }
                        }

                        if (c.GetType().ToString().Contains("ComboBox")) {LesChamps.Add(new SQLchamp(leTag[0], "'" + ((ListItem)(((ComboBox)(c)).SelectedItem)).Val + "'", "t"));}

                        if (c.GetType().ToString().Contains("CheckBox")) {LesChamps.Add(new SQLchamp(leTag[0], ((CheckBox)c).Checked ? "0" : "1", "t"));}

                        if (c.GetType().ToString().Contains("GroupBox")) {LesChamps.AddRange(FormRecupereChamp(c));}

                        //            If TypeOf(c) Is Windows.Forms.DateTimePicker Then
                        //               If c.enabled = False Or c.checked = False Then

                        //                   lesChamps.Add(New SQLchamp(c.tag, "NULL", "d"))
                        //                Else
                        //                    lesChamps.Add(New SQLchamp(c.tag, SqlDate(c), "d"))
                        //                End If
                        //            End If

                    }
                    catch {

                    }
                }
            }

            return LesChamps;

        }

        //Function FormVerif(f As Windows.Forms.Control, erp As Windows.Forms.ErrorProvider) As Boolean
        //    Dim b As Boolean = True
        //    For Each c In f.Controls
        //        Try
        //            If TypeOf(c) Is Windows.Forms.GroupBox Then
        //               If Not FormVerif(c, erp) Then b = False

        //           End If

        //           If c.tag<> "" Then
        //               If c.tag.ToString.Contains(",") Then
        //                   If c.tag.ToString.Split(",")(1).Contains("o") Then
        //                       If TypeOf(c) Is Windows.Forms.TextBox Then
        //                          If c.text = "" Then
        //                              erp.SetError(c, "Obligatoire")
        //                                b = False
        //                            End If
        //                        End If
        //                        If TypeOf(c) Is Windows.Forms.ComboBox Then
        //                           If c.tag.ToString.Contains(",t") Then
        //                               If c.text = "" Then
        //                                   erp.SetError(c, "Obligatoire")
        //                                    b = False
        //                                End If
        //                            Else
        //                                If c.selectedindex< 0 Then
        //                                    erp.SetError(c, "Obligatoire")
        //                                    b = False
        //                                End If
        //                            End If

        //                        End If
        //                    End If
        //                End If
        //            End If
        //        Catch ex As Exception
        //            MsgBox(ex.Message)
        //            Throw New Exception(ex.Message)
        //        End Try
        //    Next
        //    Return b
        //End Function

        public static int FormEnreg(Control LeControl, string LaTable, ref OleDbConnection LaConnect)
        {
            string LidChamp = "";
            int LidVal = 0;
            string SSqlAjoutChp = "";
            string SSqlAjoutVal = "";
            string SSqlModif = "";
            string SSql = "";

            //recupere les champs
            List<Common.SQLchamp> LesChamps = new List<SQLchamp>(Common.FormRecupereChamp(LeControl));

            //créer les chaines sql en ajout et modif en meme temps et mémorise l'ID
            foreach (Common.SQLchamp C in (LesChamps))
            {
                if (C.champTyp == "k")  {
                    LidChamp = C.champNom;
                    if (C.champVal != "") { LidVal = (int)Txt2Num(C.champVal.Replace("'","")); }
                } else
                {
                    if (C.champVal != "")
                    {
                        SSqlAjoutChp += C.champNom + ",";
                        SSqlAjoutVal += C.champVal + ",";

                    }
                    SSqlModif += C.champNom + "=" + C.champVal + ",";
                }
            }

            //supprime la derniere virgule de chaque chaine SQL
            SSqlModif = SSqlModif.Remove(SSqlModif.Length - 1, 1);
            SSqlAjoutChp = SSqlAjoutChp.Remove(SSqlAjoutChp.Length - 1, 1);
            SSqlAjoutVal = SSqlAjoutVal.Remove(SSqlAjoutVal.Length - 1, 1);

            //execute la requete
            if (LidVal != 0) {
                SSql = " update " + LaTable + " set " + SSqlModif + " where " + LidChamp + "=" + LidVal;
            } else {
                SSql = "insert into " + LaTable + " (" + SSqlAjoutChp + ") values (" + SSqlAjoutVal + ")";
            }
            //MessageBox.Show(SSql);
            SqlDo(SSql, ref LaConnect);

            //  recupere l'id du nouvel enreg
            if (LidVal == 0)
            {
                OleDbDataReader LeRs;
                SSql = "select max(" + LidChamp + ") M from " + LaTable;
                LeRs = Common.SqlLit(SSql, ref LaConnect);
                while (LeRs.Read())
                {
                    LidVal = LeRs.GetInt32(0);
                }
                LeRs.Close();
            }
            return LidVal;
        }

        #endregion


    }
}

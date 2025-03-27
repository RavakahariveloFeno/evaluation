using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace evaluation.uc
{
    public partial class Evaluation : UserControl
    {
        service.DataService _dataService;
        service.EvaluationService _evaluationService;
        Commission realisation = new Commission();
        Commission ro = new Commission();
        Commission resultatPondere = new Commission();

        public Evaluation()
        {
            InitializeComponent();
            _dataService = new service.DataService();
            _evaluationService = new service.EvaluationService();

            getAllTrigram();
            getAllEvaluation();
            getAllIndicateur();
        }

        private void getAllEvaluation()
        {
            DataTable dt = _evaluationService.getAll("indicateur,Information_sur_le_coeficient");
            //DataTable dt = _evaluationService.getAll("indicateur,poid,objectif,Information_sur_le_coeficient,coef");

            if (dt != null && dt.Rows.Count > 0)
            {
                dt.Columns.Add("realisation", typeof(string));
                dt.Columns.Add("R/O", typeof(string));
                dt.Columns.Add("Resultat_pondéré", typeof(string));

                dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;

                dataGridView1.Columns.Clear();
                dataGridView1.DataSource = dt;

                // Vérifier que les lignes existent avant d'accéder aux cellules
                if (dataGridView1.Rows.Count > 1)
                {
                    displayValueCommission(dataGridView1);
                }
            }
            else
            {
                MessageBox.Show("Aucune donnée disponible.");
            }

            dataGridView2.Columns.Clear();
            dataGridView2.Columns.Add("Col1", "Formule");
            dataGridView2.Columns.Add("Col2", "Valeur");

            decimal ab = resultatPondere.contribution + resultatPondere.dmt;
            decimal cdef = resultatPondere.qualite + resultatPondere.quizz + resultatPondere.t2bPersonalisation + resultatPondere.t2bSolution;
            decimal g = 0;

            dataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            dataGridView2.Rows.Add("PV1 = Somme résultat [(A) + (B)]", ab);
            dataGridView2.Rows.Add("PV2 = Somme résultat [(C ) + (D) + ( E) + (F)]", cdef);
            dataGridView2.Rows.Add("PV3 = résultat (G)",g);
            dataGridView2.Rows.Add("% d'atteinte des objectifs",ab+cdef+g);
            dataGridView2.Rows.Add("Base numéraire",0);
            dataGridView2.Rows.Add("PVV final",0);
            dataGridView2.Rows.Add("Montant commission en Ar arrondi",0);
            dataGridView2.Rows.Add("Observations","");
        }

        private void getAllTrigram()
        {
            // Récupérer le DataTable depuis la méthode getAll avec la colonne "Trigramme"
            DataTable dt = _dataService.getAll("Trigramme");

            // Vérifier si le DataTable n'est pas null et contient des données
            if (dt != null && dt.Rows.Count > 0)
            {
                // Ajouter une ligne vide au début du DataTable
                DataRow emptyRow = dt.NewRow();
                dt.Rows.InsertAt(emptyRow, 0);  // Insérer la ligne vide en première position

                // Définir la source de données pour le ComboBox
                comboBox1.DataSource = dt;

                // Spécifier quelle colonne utiliser pour l'affichage
                comboBox1.DisplayMember = "Trigramme";  // Nom de la colonne dans le DataTable
                comboBox1.ValueMember = "Trigramme";    // La valeur associée à chaque élément
            }
        }

        private decimal formuleA(decimal realisation,decimal coef,decimal poid,decimal ro) {;
            if (realisation >= coef)
            {
                return poid * ro;
            }
            return 0;
        }

        private decimal formuleB(decimal poid, decimal ro)
        {
            return poid * ro;
        }

        private void displayValueCommission(DataGridView dataGridView)
        {
            dataGridView.Rows[0].Cells["realisation"].Value = realisation.taux;
            dataGridView.Rows[1].Cells["realisation"].Value = realisation.contribution;
            dataGridView.Rows[2].Cells["realisation"].Value = realisation.dmt;
            dataGridView.Rows[3].Cells["realisation"].Value = realisation.qualite;
            dataGridView.Rows[4].Cells["realisation"].Value = realisation.quizz;
            dataGridView.Rows[5].Cells["realisation"].Value = realisation.t2bPersonalisation;
            dataGridView.Rows[6].Cells["realisation"].Value = realisation.t2bSolution;

            dataGridView.Rows[1].Cells["R/O"].Value = ro.contribution;
            dataGridView.Rows[2].Cells["R/O"].Value = ro.dmt;
            dataGridView.Rows[3].Cells["R/O"].Value = ro.qualite;
            dataGridView.Rows[4].Cells["R/O"].Value = ro.quizz;
            dataGridView.Rows[5].Cells["R/O"].Value = ro.t2bPersonalisation;
            dataGridView.Rows[6].Cells["R/O"].Value = ro.t2bSolution;

            dataGridView.Rows[1].Cells["Resultat_pondéré"].Value = resultatPondere.contribution;
            dataGridView.Rows[2].Cells["Resultat_pondéré"].Value = resultatPondere.dmt;
            dataGridView.Rows[3].Cells["Resultat_pondéré"].Value = resultatPondere.qualite;
            dataGridView.Rows[4].Cells["Resultat_pondéré"].Value = resultatPondere.quizz;
            dataGridView.Rows[5].Cells["Resultat_pondéré"].Value = resultatPondere.t2bPersonalisation;
            dataGridView.Rows[6].Cells["Resultat_pondéré"].Value = resultatPondere.t2bSolution;
        }

        private void getTrigrammeData(string trigramme)
        {
            // Récupérer la ligne correspondant au trigramme
            DataRow row = _dataService.getByTrigramme(trigramme);

            if (row != null)
            {
                Trigram.trigramme = row["Trigramme"].ToString();
                Trigram.nom = row["Nom et prénoms"].ToString();
                Trigram.typeContrat = getTypeDeContrat(trigramme);
                Trigram.date = row["Date de prise d appel"].ToString();
                Trigram.mois = countMonth(Trigram.date).ToString();
                Trigram.superviseur = row["Superviseur"].ToString();

                string taux = row["Taux_d_absentéisme"].ToString().Split('%')[0].Replace('.', ',');
                string qualite = row["Qualité"].ToString().Split('%')[0].Replace('.', ',');
                string quizz = row["Quizz"].ToString().Split('%')[0].Replace('.', ',');
                string t2bPersonalisation = row["T2B_Personnalisation"].ToString().Split('%')[0].Replace('.', ',');
                string t2bSolution = row["T2B_Solution"].ToString().Split('%')[0].Replace('.', ',');
                string contribution = row["Contribution_individuelle"].ToString();
                string dmt = row["DMT"].ToString();
                try
                {
                    realisation.taux = decimal.Parse(taux);
                    realisation.contribution = decimal.Parse(contribution);
                    realisation.dmt = decimal.Parse(dmt);
                    realisation.qualite = decimal.Parse(qualite);
                    realisation.quizz = decimal.Parse(quizz);
                    realisation.t2bPersonalisation = decimal.Parse(t2bPersonalisation);
                    realisation.t2bSolution = decimal.Parse(t2bSolution);

                    ro.contribution = realisation.contribution / tContributionObjectif.Value;
                    ro.dmt = realisation.dmt / tDmtObjectif.Value;
                    ro.qualite = realisation.qualite / tQualiteObjectif.Value;
                    ro.quizz = realisation.quizz / tQuizObjectif.Value;
                    ro.t2bPersonalisation = realisation.t2bPersonalisation / tT2bPersObjectif.Value;
                    ro.t2bSolution = realisation.t2bSolution / tT2bSolutionObjectif.Value;

                    resultatPondere.contribution = formuleA(realisation.contribution, tContributionCoef.Value, tContributionPoid.Value, ro.contribution);
                    resultatPondere.dmt = formuleB(tDmtPoid.Value, ro.dmt);
                    resultatPondere.qualite = formuleA(realisation.qualite, tQualiteCoef.Value, tQualitePoid.Value, ro.qualite);
                    resultatPondere.quizz = formuleA(realisation.quizz, tQuizCoef.Value, tQuizPoid.Value, ro.quizz);
                    resultatPondere.t2bPersonalisation = formuleA(realisation.t2bPersonalisation, tT2bPersCoef.Value, tT2bPersPoid.Value, ro.t2bPersonalisation);
                    resultatPondere.t2bSolution = formuleA(realisation.t2bSolution, tT2bSolutionCoef.Value, tT2bSolutionPoid.Value, ro.t2bSolution);
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message);
                    throw;
                }

                lprenom.Text = Trigram.nom;
                ltypeContrat.Text = Trigram.typeContrat;
                ldate.Text = Trigram.date;
                lmois.Text = Trigram.mois;
                lsuperviseur.Text = Trigram.superviseur;
            }
            else {
                lprenom.Text = "";
                ltypeContrat.Text = "";
                ldate.Text = "";
                lmois.Text = "";
                lsuperviseur.Text = "";
            }
        }

        private string getTypeDeContrat(string trigramme){
            string[] tab = trigramme.Split('_');
            if (tab.Length > 1 && tab[1].ToLower()=="tmp") {
                return "Temporaire";
            }
            return "CDI";
        }

        private int countMonth(string date)
        {
            try
            {
                int currentMonth = DateTime.Now.Month;
                int currentYear = DateTime.Now.Year;

                int x = currentYear + currentMonth;

                string[] tab = date.Split('/', ' ');
                int month = Convert.ToInt32(tab[1]);
                int year = Convert.ToInt32(tab[2]);

                int y = year + month;

                return x-y;
            }
            catch (Exception)
            {
                return 0;
                //throw;
            }
        }

        private void updateData() {
            string trigramSelected = comboBox1.Text;
            getTrigrammeData(trigramSelected);
            getAllEvaluation();
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            updateData();
        }

        private void button1_Click(object sender, EventArgs e)
        {

        }

        private void getByIndicateur(string id)
        {
            // Récupérer la ligne correspondant au trigramme
            DataRow row = _evaluationService.getByIndicateur(id);

            if (row != null)
            {
                try
                {
                    string poid = row["poid"].ToString();
                    string objectif = row["objectif"].ToString();
                    string coef = row["coef"].ToString();

                    Parametre.poid = Convert.ToInt32(poid);
                    Parametre.objectif = Convert.ToInt32(objectif);
                    Parametre.coef = Convert.ToInt32(coef);

                    if (id=="contribution")
                    {
                        tContributionPoid.Value = Parametre.poid;
                        tContributionObjectif.Value = Parametre.objectif;
                        tContributionCoef.Value = Parametre.coef;
                    }
                    else if (id=="dmt")
                    {
                        tDmtPoid.Value = Parametre.poid;
                        tDmtObjectif.Value = Parametre.objectif;
                        tDmtCoef.Value = Parametre.coef;
                    }
                    else if (id == "qualite")
                    {
                        tQualitePoid.Value = Parametre.poid;
                        tQualiteObjectif.Value = Parametre.objectif;
                        tQualiteCoef.Value = Parametre.coef;
                    }
                    else if (id == "quiz")
                    {
                        tQuizPoid.Value = Parametre.poid;
                        tQuizObjectif.Value = Parametre.objectif;
                        tQuizCoef.Value = Parametre.coef;
                    }
                    else if (id == "T2B_personalisation")
                    {
                        tT2bPersPoid.Value = Parametre.poid;
                        tT2bPersObjectif.Value = Parametre.objectif;
                        tT2bPersCoef.Value = Parametre.coef;
                    }
                    else if (id == "T2B_solution")
                    {
                        tT2bSolutionPoid.Value = Parametre.poid;
                        tT2bSolutionObjectif.Value = Parametre.objectif;
                        tT2bSolutionCoef.Value = Parametre.coef;
                    }
                    else if (id == "note")
                    {
                        tNotePoid.Value = Parametre.poid;
                        tNoteObjectif.Value = Parametre.objectif;
                        tNoteCoef.Value = Parametre.coef;
                    }
                    else if (id == "taux")
                    {
                        tTauxPoid.Value = Parametre.poid;
                        tTauxObjectif.Value = Parametre.objectif;
                        tTauxCoef.Value = Parametre.coef;
                    }
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message);
                    throw;
                }
            }
        }

        private void getAllIndicateur()
        {
            getByIndicateur("contribution");
            getByIndicateur("dmt");
            getByIndicateur("qualite");
            getByIndicateur("quiz");
            getByIndicateur("T2B_personalisation");
            getByIndicateur("T2B_solution");
            getByIndicateur("note");
            getByIndicateur("taux");
        }

        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dataGridView2_CellContentClick_1(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void tContributionPoid_ValueChanged(object sender, EventArgs e)
        {
            updateData();
        }

        private void tContributionPoid_KeyPress(object sender, KeyPressEventArgs e)
        {
            updateData();
        }
    }

    public static class Trigram
    {
        public static string trigramme { get; set; }
        public static string nom { get; set; }
        public static string typeContrat { get; set; }
        public static string date { get; set; }
        public static string mois { get; set; }
        public static string superviseur { get; set; }
        public static string taux { get; set; }
    }

    public static class Parametre
    {
        public static int poid = 0;
        public static int objectif = 0;
        public static int coef = 0;
    }

    public class Commission
    {
        public decimal taux =0;
        public decimal contribution = 0;
        public decimal dmt = 0;
        public decimal qualite = 0;
        public decimal quizz = 0;
        public decimal t2bPersonalisation = 0;
        public decimal t2bSolution = 0;
        public decimal notemanageriale = 0;
    }
}

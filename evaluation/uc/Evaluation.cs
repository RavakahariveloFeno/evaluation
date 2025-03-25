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
        float tauxRealisation = -1;

        Trigram trigram;
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
            DataTable dt = _evaluationService.getAll("indicateur,poid,objectif,realisation,RO,Information_sur_le_coeficient,coef,Résultat_pondéré");

            if (dt != null && dt.Rows.Count > 0)
            {
                dataGridView1.Columns.Clear();
                dataGridView1.DataSource = dt;

                // Vérifier que les lignes existent avant d'accéder aux cellules
                if (dataGridView1.Rows.Count > 1)
                {
                    dataGridView1.Rows[0].Cells["realisation"].Value = tauxRealisation;
                }
            }
            else
            {
                MessageBox.Show("Aucune donnée disponible.");
            }
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


        private void getTrigrammeData(string trigramme)
        {
            // Récupérer la ligne correspondant au trigramme
            DataRow row = _dataService.getByTrigramme(trigramme);

            if (row != null)
            {
                trigram = new Trigram();

                trigram.trigramme = row["Trigramme"].ToString();
                trigram.nom = row["Nom et prénoms"].ToString();
                trigram.typeContrat = getTypeDeContrat(trigramme);
                trigram.date = row["Date de prise d appel"].ToString();
                trigram.mois = countMonth(trigram.date).ToString();
                trigram.superviseur = row["Superviseur"].ToString();

                string taux = row["Taux_d_absentéisme"].ToString().Split('%')[0].Replace('.',',');
                try
                {
                    tauxRealisation = float.Parse(taux);
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message);
                    throw;
                }

                lprenom.Text = trigram.nom;
                ltypeContrat.Text = trigram.typeContrat;
                ldate.Text = trigram.date;
                lmois.Text = trigram.mois;
                lsuperviseur.Text = trigram.superviseur;
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

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            string trigramSelected = comboBox1.Text;
            getTrigrammeData(trigramSelected);

            getAllEvaluation();
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
                    Parametre parametre = new Parametre();
                    string poid = row["poid"].ToString();
                    string objectif = row["objectif"].ToString();
                    string coef = row["coef"].ToString();

                    parametre.poid = Convert.ToInt32(poid);
                    parametre.objectif = Convert.ToInt32(objectif);
                    parametre.coef = Convert.ToInt32(coef);

                    if (id=="contribution")
                    {
                        tContributionPoid.Value = parametre.poid;
                        tContributionObjectif.Value = parametre.objectif;
                        tContributionCoef.Value = parametre.coef;
                    }
                    else if (id=="dmt")
                    {
                        tDmtPoid.Value = parametre.poid;
                        tDmtObjectif.Value = parametre.objectif;
                        tDmtCoef.Value = parametre.coef;
                    }
                    else if (id == "qualite")
                    {
                        tQualitePoid.Value = parametre.poid;
                        tQualiteObjectif.Value = parametre.objectif;
                        tQualiteCoef.Value = parametre.coef;
                    }
                    else if (id == "quiz")
                    {
                        tQuizPoid.Value = parametre.poid;
                        tQuizObjectif.Value = parametre.objectif;
                        tQuizCoef.Value = parametre.coef;
                    }
                    else if (id == "T2B_personalisation")
                    {
                        tT2bPersPoid.Value = parametre.poid;
                        tT2bPersObjectif.Value = parametre.objectif;
                        tT2bPersCoef.Value = parametre.coef;
                    }
                    else if (id == "T2B_solution")
                    {
                        tT2bSolutionPoid.Value = parametre.poid;
                        tT2bSolutionObjectif.Value = parametre.objectif;
                        tT2bSolutionCoef.Value = parametre.coef;
                    }
                    else if (id == "note")
                    {
                        tNotePoid.Value = parametre.poid;
                        tNoteObjectif.Value = parametre.objectif;
                        tNoteCoef.Value = parametre.coef;
                    }
                    else if (id == "taux")
                    {
                        tTauxPoid.Value = parametre.poid;
                        tTauxObjectif.Value = parametre.objectif;
                        tTauxCoef.Value = parametre.coef;
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
    }

    public class Trigram {
        public string trigramme { get; set; }
        public string nom { get; set; }
        public string typeContrat { get; set; }
        public string date { get; set; }
        public string mois { get; set; }
        public string superviseur { get; set; }
        public string taux { get; set; }
    }

    public class Parametre
    {
        public int poid { get; set; }
        public int objectif { get; set; }
        public int coef { get; set; }
    }
}

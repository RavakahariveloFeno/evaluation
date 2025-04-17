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
        string[] Mois = new string[] 
        {
            "janvier", "février", "mars", "avril", "mai", "juin",
            "juillet", "août", "septembre", "octobre", "novembre", "décembre"
        };

        decimal tContributionPoid, tContributionObjectif, tContributionCoef, tDmtPoid, tDmtObjectif, tDmtCoef, tQualitePoid, tQualiteObjectif, tQualiteCoef, tQuizPoid, tQuizObjectif, tQuizCoef, tT2bPersPoid, tT2bPersObjectif, tT2bPersCoef, tT2bSolutionPoid, tT2bSolutionObjectif, tT2bSolutionCoef, tNotePoid;
        service.DataService _dataService;
        service.EvaluationService _evaluationService;
        service.CsvService _csvService;
        decimal baseNumeriale = 0, montant = 0;

        type.Commission realisation = new type.Commission();
        type.Commission ro = new type.Commission();
        type.Commission resultatPondere = new type.Commission();

        public Evaluation()
        {
            InitializeComponent();
            dataGridView1.RowPrePaint += dataGridView1_RowPrePaint;
            dataGridView1.DataBindingComplete += dataGridView1_DataBindingComplete;

            dataGridView1.DefaultCellStyle.Font = new Font("Arial", 11);

            _dataService = new service.DataService();
            _evaluationService = new service.EvaluationService();
            _csvService = new service.CsvService();


            getParametre();
            getAllTrigram();
            getAllIdAgent();
            getAllEvaluation();
            getAllIndicateur();
        }

        private void dataGridView1_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            DataGridView dgv = sender as DataGridView;
            dgv.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None;
            dgv.RowTemplate.Height = 80;

            foreach (DataGridViewRow row in dgv.Rows)
            {
                row.Height = 80;
            }
        }


        private void dataGridView1_RowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)
        {
            DataGridView dgv = sender as DataGridView;

            if (e.RowIndex < 0 || e.RowIndex >= dgv.Rows.Count)
                return;

            // Appliquer une couleur différente selon la ligne avec des codes RGB
            switch (e.RowIndex)
            {
                case 0:
                    dgv.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.FromArgb(145 ,100 ,205);
                    break;
                case 1:
                    dgv.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.FromArgb(255 ,121 ,0);
                    break;
                case 2:
                    dgv.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.FromArgb(255 ,121 ,0);
                    break;
                case 3:
                    dgv.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.FromArgb(75 ,180 ,230);
                    break;
                case 4:
                    dgv.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.FromArgb(75 ,180 ,230);
                    break;
                case 5:
                    dgv.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.FromArgb(255 ,180 ,230);
                    break;
                case 6:
                    dgv.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.FromArgb(255 ,180 ,230);
                    break;
                case 7:
                    dgv.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.FromArgb(255 ,210 ,0);
                    break;
                default:
                    dgv.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.FromArgb(255, 255, 255); // White
                    break;
            }
        }



        private void getParametre(){
            decimal[] csv = _csvService.ReadValues();

            if (csv != null && csv.Length > 1) // ✅ Vérification pour éviter l'erreur
            {
                baseNumeriale = csv[0];
                montant = csv[1];
            }
            else
            {
                baseNumeriale = 0; // Valeur par défaut
                montant = 0;       // Valeur par défaut
            }
        }

        private void getAllEvaluation()
        {
            DataTable dt = _evaluationService.getAll("Indicateurs,Information_sur_le_coeficient");

            if (dt != null && dt.Rows.Count > 0)
            {
                string contributionCoef = dt.Rows[1]["Information_sur_le_coeficient"].ToString().Replace("{x}", tContributionCoef.ToString());
                string qualiteCoef = dt.Rows[3]["Information_sur_le_coeficient"].ToString().Replace("{x}", tQualiteCoef.ToString());
                string quizzCoef = dt.Rows[4]["Information_sur_le_coeficient"].ToString().Replace("{x}", tQuizCoef.ToString());
                string personalitionCoef = dt.Rows[5]["Information_sur_le_coeficient"].ToString().Replace("{x}", tT2bPersCoef.ToString());
                string solutionCoef = dt.Rows[6]["Information_sur_le_coeficient"].ToString().Replace("{x}", tT2bSolutionCoef.ToString());
                // Modification des valeurs de chaque ligne individuellement
                dt.Rows[1]["Information_sur_le_coeficient"] = contributionCoef;
                dt.Rows[3]["Information_sur_le_coeficient"] = qualiteCoef;
                dt.Rows[4]["Information_sur_le_coeficient"] = quizzCoef;
                dt.Rows[5]["Information_sur_le_coeficient"] = personalitionCoef;
                dt.Rows[6]["Information_sur_le_coeficient"] = solutionCoef;

                // Ajouter les nouvelles colonnes
                dt.Columns.Add("Poids", typeof(string));
                dt.Columns.Add("Objectif", typeof(string));
                dt.Columns.Add("Coefficient", typeof(string));
                dt.Columns.Add("Réalisation", typeof(string));
                dt.Columns.Add("R/O", typeof(string));
                dt.Columns.Add("Résultat_pondéré", typeof(string));

                // Configuration du DataGridView
                dataGridView1.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
                dataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
                dataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
                dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

                dataGridView1.Columns.Clear();
                dataGridView1.DataSource = dt;
                dataGridView1.AllowUserToAddRows = false;

                // Ajuster la largeur des colonnes
                //dataGridView1.Columns["indicateur"].Width = 200;
                //dataGridView1.Columns["Information_sur_le_coeficient"].Width = 250;

                dataGridView1.DefaultCellStyle.Padding = new Padding(3);

                // Définir les polices
                dataGridView1.DefaultCellStyle.Font = new Font("Arial", 11, FontStyle.Regular);
                dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font("Arial", 11, FontStyle.Bold);

                // Vérifier l'existence des lignes avant d'accéder aux cellules
                if (dataGridView1.Rows.Count > 1)
                {
                    displayValueCommission(dataGridView1);
                }
            }
            else
            {
                MessageBox.Show("Aucune donnée disponible.");
            }

            // Calcul et affichage des valeurs
            decimal ab = Math.Round(resultatPondere.contribution + resultatPondere.dmt, 2);
            decimal cdef = Math.Round(resultatPondere.qualite + resultatPondere.quizz + resultatPondere.t2bPersonalisation + resultatPondere.t2bSolution, 2);
            decimal g = realisation.notemanageriale;

            lpv1.Text = ab.ToString() + "%";
            lpv2.Text = cdef.ToString() + "%";
            lpv3.Text = g.ToString() + "%";
            latteinte.Text = Math.Round((ab + cdef + g), 2).ToString() + "%";
            lbase.Text = baseNumeriale.ToString("#,##0") + " Ar";
            lpvvfinal.Text = Math.Round((ab/100 + cdef/100 + g/100) * baseNumeriale, 2).ToString("#,##0") + " Ar";
            lmontant.Text = montant.ToString("#,##0") + " Ar";

            // Mise en forme conditionnelle pour Résultat_pondéré > Poids
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (row.Index == 0) continue; // Ignorer la première ligne si nécessaire

                try
                {
                    string poidsStr = row.Cells["Poids"].Value != null ? row.Cells["Poids"].Value.ToString().Replace("%", "") : "0";
                    string resultatStr = row.Cells["Résultat_pondéré"].Value != null ? row.Cells["Résultat_pondéré"].Value.ToString().Replace("%", "") : "0";

                    decimal poids = 0;
                    decimal resultat = 0;

                    if (Decimal.TryParse(poidsStr, out poids) && Decimal.TryParse(resultatStr, out resultat))
                    {
                        if (resultat > poids)
                        {
                            row.Cells["Résultat_pondéré"].Style.ForeColor = Color.Red;
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Erreur lors de la mise en forme conditionnelle : " + ex.Message);
                }
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

        private void getAllIdAgent()
        {
            // Récupérer le DataTable depuis la méthode getAll avec la colonne "ID_agent"
            DataTable dt = _dataService.getAll("ID_agent");

            // Vérifier si le DataTable n'est pas null et contient des données
            if (dt != null && dt.Rows.Count > 0)
            {
                // Supprimer les doublons sur la colonne "ID_agent"
                DataView view = new DataView(dt);
                DataTable distinctTable = view.ToTable(true, "ID_agent"); // 'true' = distinct

                // Ajouter une ligne vide au début du DataTable
                DataRow emptyRow = distinctTable.NewRow();
                distinctTable.Rows.InsertAt(emptyRow, 0);  // Insérer la ligne vide en première position

                // Définir la source de données pour le ComboBox
                comboBox2.DataSource = distinctTable;
                comboBox2.DisplayMember = "ID_agent";
                comboBox2.ValueMember = "ID_agent";
            }
        }

        private decimal formuleA(decimal realisation,decimal coef,decimal poid,decimal ro) {;
            if (realisation >= coef)
            {
                decimal x = (poid * ro)/100;
                return Math.Round(x,2);
            }
            return 0;
        }

        private decimal formuleB(decimal poid, decimal ro)
        {
            return Math.Round((poid * ro) / 100, 2);
        }

        private void displayValueCommission(DataGridView dataGridView)
        {
            dataGridView.Rows[1].Cells["Poids"].Value = tContributionPoid + "%";
            dataGridView.Rows[2].Cells["Poids"].Value = tDmtPoid + "%";
            dataGridView.Rows[3].Cells["Poids"].Value = tQualitePoid + "%";
            dataGridView.Rows[4].Cells["Poids"].Value = tQuizPoid + "%";
            dataGridView.Rows[5].Cells["Poids"].Value = tT2bPersPoid + "%";
            dataGridView.Rows[6].Cells["Poids"].Value = tT2bSolutionPoid + "%";
            dataGridView.Rows[7].Cells["Poids"].Value = tNotePoid + "%";

            dataGridView.Rows[1].Cells["Objectif"].Value = tContributionObjectif + "%";
            dataGridView.Rows[2].Cells["Objectif"].Value = tDmtObjectif;
            dataGridView.Rows[3].Cells["Objectif"].Value = tQualiteObjectif + "%";
            dataGridView.Rows[4].Cells["Objectif"].Value = tQuizObjectif;
            dataGridView.Rows[5].Cells["Objectif"].Value = tT2bPersObjectif + "%";
            dataGridView.Rows[6].Cells["Objectif"].Value = tT2bSolutionObjectif + "%";

            dataGridView.Rows[1].Cells["Coefficient"].Value = tContributionCoef + "%";
            dataGridView.Rows[2].Cells["Coefficient"].Value = tDmtCoef + "%";
            dataGridView.Rows[3].Cells["Coefficient"].Value = tQualiteCoef + "%";
            dataGridView.Rows[4].Cells["Coefficient"].Value = tQuizCoef;
            dataGridView.Rows[5].Cells["Coefficient"].Value = tT2bPersCoef + "%";
            dataGridView.Rows[6].Cells["Coefficient"].Value = tT2bSolutionCoef + "%";

            dataGridView.Rows[0].Cells["Réalisation"].Value = realisation.taux + "%";
            dataGridView.Rows[1].Cells["Réalisation"].Value = realisation.contribution + "%";
            dataGridView.Rows[2].Cells["Réalisation"].Value = realisation.dmt;
            dataGridView.Rows[3].Cells["Réalisation"].Value = realisation.qualite + "%";
            dataGridView.Rows[4].Cells["Réalisation"].Value = Math.Round(realisation.quizz, 2);
            dataGridView.Rows[5].Cells["Réalisation"].Value = realisation.t2bPersonalisation + "%";
            dataGridView.Rows[6].Cells["Réalisation"].Value = realisation.t2bSolution + "%";
            dataGridView.Rows[7].Cells["Réalisation"].Value = realisation.notemanageriale + "%";

            dataGridView.Rows[1].Cells["R/O"].Value = ro.contribution + "%";
            dataGridView.Rows[2].Cells["R/O"].Value = ro.dmt + "%";
            dataGridView.Rows[3].Cells["R/O"].Value = ro.qualite + "%";
            dataGridView.Rows[4].Cells["R/O"].Value = ro.quizz + "%";
            dataGridView.Rows[5].Cells["R/O"].Value = ro.t2bPersonalisation + "%";
            dataGridView.Rows[6].Cells["R/O"].Value = ro.t2bSolution + "%";

            dataGridView.Rows[1].Cells["Résultat_pondéré"].Value = resultatPondere.contribution + "%";
            dataGridView.Rows[2].Cells["Résultat_pondéré"].Value = resultatPondere.dmt + "%";
            dataGridView.Rows[3].Cells["Résultat_pondéré"].Value = resultatPondere.qualite + "%";
            dataGridView.Rows[4].Cells["Résultat_pondéré"].Value = resultatPondere.quizz + "%";
            dataGridView.Rows[5].Cells["Résultat_pondéré"].Value = resultatPondere.t2bPersonalisation + "%";
            dataGridView.Rows[6].Cells["Résultat_pondéré"].Value = resultatPondere.t2bSolution + "%";
            dataGridView.Rows[7].Cells["Résultat_pondéré"].Value = realisation.notemanageriale + "%";
        }

        private void getTrigrammeData(string trigramme)
        {
            // Récupérer la ligne correspondant au trigramme
            DataRow row = _dataService.getByTrigramme(trigramme);

            if (row != null)
            {
                type.Trigram.trigramme = row["Trigramme"].ToString();
                type.Trigram.nom = row["Nom et prénoms"].ToString();
                type.Trigram.typeContrat = getTypeDeContrat(trigramme);
                type.Trigram.date = row["Date de prise d appel"].ToString();
                //type.Trigram.mois = countMonth(type.Trigram.date).ToString();
                type.Trigram.mois = row["Ancieneté en mois"].ToString();
                type.Trigram.superviseur = row["Superviseur"].ToString();
                type.Trigram.observation = row["Observations"].ToString();
                type.Trigram.idAgent = row["Id_agent"].ToString();
                richTextBox1.Text = type.Trigram.observation;

                string taux = row["Taux_d_absentéisme"].ToString().Split('%')[0].Replace('.', ',');
                string qualite = row["Qualité"].ToString().Split('%')[0].Replace('.', ',');
                string quizz = row["Quizz"].ToString().Split('%')[0].Replace('.', ',');
                string t2bPersonalisation = row["T2B_Personnalisation"].ToString().Split('%')[0].Replace('.', ',');
                string t2bSolution = row["T2B_Solution"].ToString().Split('%')[0].Replace('.', ',');
                string contribution = row["R/O en CI"].ToString().Split('%')[0].Replace('.', ',');
                string dmt = row["DMT"].ToString();
                string notemanageriale = row["Appréciation_managériale"].ToString().ToString().Split('%')[0].Replace('.', ',');

                try
                {
                    realisation.taux = taux=="" ? 0 : Math.Round(decimal.Parse(taux) * 100);
                    realisation.contribution = contribution == "" ? 0 : Math.Round(decimal.Parse(contribution) * 100);
                    realisation.dmt = dmt == "" ? 0 : decimal.Parse(dmt);
                    realisation.qualite = qualite == "" ? 0 : Math.Round(decimal.Parse(qualite) * 100);
                    realisation.quizz = quizz == "" ? 0 : Math.Round((decimal.Parse(quizz) * 20 / 100) * 100);
                    realisation.t2bPersonalisation = t2bPersonalisation == "" ? 0 : Math.Round(decimal.Parse(t2bPersonalisation) * 100);
                    realisation.t2bSolution = t2bSolution == "" ? 0 : Math.Round(decimal.Parse(t2bSolution) * 100);
                    realisation.notemanageriale = notemanageriale == "" ? 0 : Math.Round(decimal.Parse(notemanageriale) * 100);

                    ro.contribution = Math.Round((realisation.contribution / tContributionObjectif)*100,2);

                    _csvService = new service.CsvService("bareme.csv");
                    string roDmt = _csvService.GetPourcentageBarem(realisation.dmt).Split('%')[0].Replace('.', ',');
                    ro.dmt = roDmt == "" ? 0 : decimal.Parse(roDmt);

                    ro.qualite = Math.Round((realisation.qualite / tQualiteObjectif) * 100, 2);
                    ro.quizz = Math.Round(((realisation.quizz * 100) / tQuizObjectif), 2);
                    ro.t2bPersonalisation = Math.Round((realisation.t2bPersonalisation / tT2bPersObjectif) * 100, 2);
                    ro.t2bSolution = Math.Round((realisation.t2bSolution / tT2bSolutionObjectif) * 100, 2);

                    resultatPondere.contribution = formuleA(realisation.contribution, tContributionCoef, tContributionPoid, ro.contribution);
                    resultatPondere.dmt = formuleB(tDmtPoid, ro.dmt);
                    resultatPondere.qualite = formuleA(realisation.qualite, tQualiteCoef, tQualitePoid, ro.qualite);
                    resultatPondere.quizz = formuleA(realisation.quizz, tQuizCoef, tQuizPoid, ro.quizz);
                    resultatPondere.t2bPersonalisation = formuleA(realisation.t2bPersonalisation, tT2bPersCoef, tT2bPersPoid, ro.t2bPersonalisation);
                    resultatPondere.t2bSolution = formuleA(realisation.t2bSolution, tT2bSolutionCoef, tT2bSolutionPoid, ro.t2bSolution);
                }
                catch (Exception e)
                {
                    realisation.taux = 0;
                    realisation.contribution = 0;
                    realisation.dmt = 0;
                    realisation.qualite = 0;
                    realisation.quizz = 0;
                    realisation.t2bPersonalisation = 0;
                    realisation.t2bSolution = 0;

                    ro.contribution = 0;
                    ro.dmt = 0;
                    ro.qualite = 0;
                    ro.quizz = 0;
                    ro.t2bPersonalisation = 0;
                    ro.t2bSolution = 0;

                    resultatPondere.contribution = 0;
                    resultatPondere.dmt = 0;
                    resultatPondere.qualite = 0;
                    resultatPondere.quizz = 0;
                    resultatPondere.t2bPersonalisation = 0;
                    resultatPondere.t2bSolution = 0;
                }

                lprenom.Text = type.Trigram.nom;
                lIdAgent.Text = type.Trigram.idAgent;
                ltypeContrat.Text = type.Trigram.typeContrat;
                ldate.Text = type.Trigram.date;
                lmois.Text = type.Trigram.mois;
                lsuperviseur.Text = type.Trigram.superviseur;
                ltrigramme.Text = comboBox1.Text;
            }
            else {
                lprenom.Text = "";
                lIdAgent.Text = "";
                ltypeContrat.Text = "";
                ldate.Text = "";
                lmois.Text = "";
                lsuperviseur.Text = "";
                ltrigramme.Text = "";
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
                    string poid = row["Poids"].ToString();
                    string objectif = row["Objectif"].ToString();
                    string coef = row["Coefficient"].ToString();

                    type.Parametre.poid = Convert.ToInt32(poid);
                    type.Parametre.objectif = Convert.ToInt32(objectif);
                    type.Parametre.coef = Convert.ToInt32(coef);

                    if (id == "contribution")
                    {
                        tContributionPoid = type.Parametre.poid;
                        tContributionObjectif = type.Parametre.objectif;
                        tContributionCoef = type.Parametre.coef;
                    }
                    else if (id == "dmt")
                    {
                        tDmtPoid = type.Parametre.poid;
                        tDmtObjectif = type.Parametre.objectif;
                        tDmtCoef = type.Parametre.coef;
                    }
                    else if (id == "qualite")
                    {
                        tQualitePoid = type.Parametre.poid;
                        tQualiteObjectif = type.Parametre.objectif;
                        tQualiteCoef = type.Parametre.coef;
                    }
                    else if (id == "quiz")
                    {
                        tQuizPoid = type.Parametre.poid;
                        tQuizObjectif = type.Parametre.objectif;
                        tQuizCoef = type.Parametre.coef;
                    }
                    else if (id == "T2B_personalisation")
                    {
                        tT2bPersPoid = type.Parametre.poid;
                        tT2bPersObjectif = type.Parametre.objectif;
                        tT2bPersCoef = type.Parametre.coef;
                    }
                    else if (id == "T2B_solution")
                    {
                        tT2bSolutionPoid = type.Parametre.poid;
                        tT2bSolutionObjectif = type.Parametre.objectif;
                        tT2bSolutionCoef = type.Parametre.coef;
                    }
                    else if (id == "note")
                    {
                        tNotePoid = type.Parametre.poid;
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

        private void panel4_Paint(object sender, PaintEventArgs e)
        {

        }

        private void label38_Click(object sender, EventArgs e)
        {

        }

        private void panel3_Paint(object sender, PaintEventArgs e)
        {

        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            string dateTrigram = Mois[dateTimePicker1.Value.Month] + " " + dateTimePicker1.Value.Year;
            window.ExportType exportType = new window.ExportType(
                tContributionPoid, 
                tContributionObjectif, 
                tContributionCoef, 
                tDmtPoid, 
                tDmtObjectif, 
                tDmtCoef, 
                tQualitePoid, 
                tQualiteObjectif, 
                tQualiteCoef, 
                tQuizPoid, 
                tQuizObjectif, 
                tQuizCoef, 
                tT2bPersPoid, 
                tT2bPersObjectif, 
                tT2bPersCoef, 
                tT2bSolutionPoid, 
                tT2bSolutionObjectif, 
                tT2bSolutionCoef, 
                tNotePoid,
                montant,
                baseNumeriale,
                dateTrigram,
                comboBox2.Text
                );
            exportType.ShowDialog();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            updateData();
        }

        private void tableLayoutPanel1_Paint(object sender, PaintEventArgs e)
        {

        }
    }
}

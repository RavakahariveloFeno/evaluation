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
    public partial class Settings : UserControl
    {
        service.EvaluationService _evaluationService;
        service.CsvService _csvService;
        decimal baseNumeriale = 0, montant = 0;
        public Settings()
        {
            InitializeComponent();
            dataGridView1.DefaultCellStyle.Font = new Font("Arial", 11);
            dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font("Arial", 11, FontStyle.Bold);
            dataGridView1.ReadOnly = false; // Permet la modification des cellules
            dataGridView1.AllowUserToAddRows = false; // Désactive l'ajout de nouvelles lignes manuelles

            _evaluationService = new service.EvaluationService();
            _csvService = new service.CsvService();

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

            tbase.Value = baseNumeriale;
            tmontant.Value = montant;
            getAllIndicateur();
            getAllBarem();
        }

        private void getAllBarem() {
            _csvService = new service.CsvService("bareme.csv");
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dataGridView1.DataSource = _csvService.ReadBarem();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void tNotePoid_ValueChanged(object sender, EventArgs e)
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
                        tContributionPoid.Value = type.Parametre.poid;
                        tContributionObjectif.Value = type.Parametre.objectif;
                        tContributionCoef.Value = type.Parametre.coef;
                    }
                    else if (id == "dmt")
                    {
                        tDmtPoid.Value = type.Parametre.poid;
                        tDmtObjectif.Value = type.Parametre.objectif;
                        tDmtCoef.Value = type.Parametre.coef;
                    }
                    else if (id == "qualite")
                    {
                        tQualitePoid.Value = type.Parametre.poid;
                        tQualiteObjectif.Value = type.Parametre.objectif;
                        tQualiteCoef.Value = type.Parametre.coef;
                    }
                    else if (id == "quiz")
                    {
                        tQuizPoid.Value = type.Parametre.poid;
                        tQuizObjectif.Value = type.Parametre.objectif;
                        tQuizCoef.Value = type.Parametre.coef;
                    }
                    else if (id == "T2B_personalisation")
                    {
                        tT2bPersPoid.Value = type.Parametre.poid;
                        tT2bPersObjectif.Value = type.Parametre.objectif;
                        tT2bPersCoef.Value = type.Parametre.coef;
                    }
                    else if (id == "T2B_solution")
                    {
                        tT2bSolutionPoid.Value = type.Parametre.poid;
                        tT2bSolutionObjectif.Value = type.Parametre.objectif;
                        tT2bSolutionCoef.Value = type.Parametre.coef;
                    }
                    else if (id == "note")
                    {
                        tNotePoid.Value = type.Parametre.poid;
                        //tNoteObjectif.Value = Parametre.objectif;
                        //tNoteCoef.Value = Parametre.coef;
                    }
                    //else if (id == "taux")
                    //{
                    //    tTauxPoid.Value = Parametre.poid;
                    //    tTauxObjectif.Value = Parametre.objectif;
                    //    tTauxCoef.Value = Parametre.coef;
                    //}
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

        private void saveBareme()
        {
            _csvService = new service.CsvService("bareme.csv");

            List<type.Bareme> updatedData = new List<type.Bareme>();

            // Parcourir chaque ligne du DataGridView
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (!row.IsNewRow) // Évite la ligne vide
                {
                    updatedData.Add(new type.Bareme
                    {
                        Pourcentage = row.Cells[0].Value.ToString(),
                        Min = row.Cells[1].Value.ToString(),
                        Max = row.Cells[2].Value.ToString()
                    });
                }
            }

            _csvService.SaveToCsv(updatedData);
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            bool success = _evaluationService.updateIndicateur("contribution", (int)tContributionPoid.Value, (int)tContributionObjectif.Value, (int)tContributionCoef.Value);
            success &= _evaluationService.updateIndicateur("dmt", (int)tDmtPoid.Value, (int)tDmtObjectif.Value, (int)tDmtCoef.Value);
            success &= _evaluationService.updateIndicateur("qualite", (int)tQualitePoid.Value, (int)tQualiteObjectif.Value, (int)tQualiteCoef.Value);
            success &= _evaluationService.updateIndicateur("quiz", (int)tQuizPoid.Value, (int)tQuizObjectif.Value, (int)tQuizCoef.Value);
            success &= _evaluationService.updateIndicateur("T2B_personalisation", (int)tT2bPersPoid.Value, (int)tT2bPersObjectif.Value, (int)tT2bPersCoef.Value);
            success &= _evaluationService.updateIndicateur("T2B_solution", (int)tT2bSolutionPoid.Value, (int)tT2bSolutionObjectif.Value, (int)tT2bSolutionCoef.Value);
            success &= _evaluationService.updateIndicateur("note", (int)tNotePoid.Value, 0, 0); // Pas de objectif et coef pour "note"

            _csvService = new service.CsvService();
            _csvService.UpdateValues(tbase.Value, tmontant.Value);

            saveBareme();

            if (success)
            {
                MessageBox.Show("Mise à jour réussie !");
            }
            else
            {
                MessageBox.Show("Erreur lors de la mise à jour.");
            }
        }
        
        private void button1_Click(object sender, EventArgs e)
        {
            getAllIndicateur();
        }

        private void label25_Click(object sender, EventArgs e)
        {

        }

    }
}

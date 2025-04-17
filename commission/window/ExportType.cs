using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using System.IO;
using iTextSharp.text;
using iTextSharp.text.pdf;

using Excel = Microsoft.Office.Interop.Excel;

namespace evaluation.window
{
    public partial class ExportType : Form
    {
        service.CsvService _csvService;
        service.DataService _dataService;
        type.Commission realisation = new type.Commission();
        type.Commission ro = new type.Commission();
        type.Commission resultatPondere = new type.Commission();
        decimal tContributionPoid, tContributionObjectif, tContributionCoef, tDmtPoid, tDmtObjectif, tDmtCoef, tQualitePoid, tQualiteObjectif, tQualiteCoef, tQuizPoid, tQuizObjectif, tQuizCoef, tT2bPersPoid, tT2bPersObjectif, tT2bPersCoef, tT2bSolutionPoid, tT2bSolutionObjectif, tT2bSolutionCoef, tNotePoid, tmontant, tbaseNumeriale;
        decimal ab, cdef, g, atteinte, baseNumeriale, pvvFinal, montant;
        string dateTrigram, idAgentSelected;
        public ExportType(
            decimal tContributionPoid,
            decimal tContributionObjectif,
            decimal tContributionCoef,
            decimal tDmtPoid,
            decimal tDmtObjectif,
            decimal tDmtCoef,
            decimal tQualitePoid,
            decimal tQualiteObjectif,
            decimal tQualiteCoef,
            decimal tQuizPoid,
            decimal tQuizObjectif,
            decimal tQuizCoef,
            decimal tT2bPersPoid,
            decimal tT2bPersObjectif,
            decimal tT2bPersCoef,
            decimal tT2bSolutionPoid,
            decimal tT2bSolutionObjectif,
            decimal tT2bSolutionCoef,
            decimal tNotePoid,
            decimal tmontant,
            decimal tbaseNumeriale,
            string dateTrigram,
            string idAgentSelected
            )
        {
            InitializeComponent();
            _dataService = new service.DataService();

            this.tContributionPoid = tContributionPoid;
            this.tContributionObjectif = tContributionObjectif;
            this.tContributionCoef = tContributionCoef;
            this.tDmtPoid = tDmtPoid;
            this.tDmtObjectif = tDmtObjectif;
            this.tDmtCoef = tDmtCoef;
            this.tQualitePoid = tQualitePoid;
            this.tQualiteObjectif = tQualiteObjectif;
            this.tQualiteCoef = tQualiteCoef;
            this.tQuizPoid = tQuizPoid;
            this.tQuizObjectif = tQuizObjectif;
            this.tQuizCoef = tQuizCoef;
            this.tT2bPersPoid = tT2bPersPoid;
            this.tT2bPersObjectif = tT2bPersObjectif;
            this.tT2bPersCoef = tT2bPersCoef;
            this.tT2bSolutionPoid = tT2bSolutionPoid;
            this.tT2bSolutionObjectif = tT2bSolutionObjectif;
            this.tT2bSolutionCoef = tT2bSolutionCoef;
            this.tNotePoid = tNotePoid;
            this.tmontant=tmontant;
            this.tbaseNumeriale = tbaseNumeriale;
            this.dateTrigram = dateTrigram;
            this.idAgentSelected = idAgentSelected;
        }

        private void ExportToExcel(DataGridView dgv)
        {
            if (dgv.Rows.Count > 0)
            {
                using (SaveFileDialog saveFileDialog = new SaveFileDialog())
                {
                    saveFileDialog.Filter = "Fichier Excel (*.xlsx)|*.xlsx";
                    saveFileDialog.Title = "Enregistrer sous";
                    saveFileDialog.FileName = "MatricePVV " + dateTrigram + ".xlsx";

                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        Excel.Application excelApp = new Excel.Application();
                        Excel.Workbook workbook = excelApp.Workbooks.Add();
                        Excel.Worksheet worksheet = workbook.Sheets[1];

                        // Ajouter les en-têtes de colonnes avec fond coloré
                        for (int i = 0; i < dgv.Columns.Count; i++)
                        {
                            worksheet.Cells[1, i + 1] = dgv.Columns[i].HeaderText;
                            Excel.Range headerCell = worksheet.Cells[1, i + 1];
                            headerCell.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow); // Couleur de fond
                            headerCell.Font.Bold = true; // Texte en gras pour les en-têtes
                        }

                        // Ajouter les données des cellules
                        for (int i = 0; i < dgv.Rows.Count; i++)
                        {
                            for (int j = 0; j < dgv.Columns.Count; j++)
                            {
                                worksheet.Cells[i + 2, j + 1] = dgv.Rows[i].Cells[j].Value != null ? dgv.Rows[i].Cells[j].Value.ToString() : "";
                            }
                        }

                        // Ajustement automatique de la largeur des colonnes
                        worksheet.Columns.AutoFit();

                        // Sauvegarde du fichier à l'emplacement choisi
                        workbook.SaveAs(saveFileDialog.FileName);

                        // Libération des ressources
                        workbook.Close();
                        excelApp.Quit();

                        MessageBox.Show("Exportation réussie !", "Succès", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        this.Close();
                    }
                }
            }
            else
            {
                MessageBox.Show("Aucune donnée à exporter.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }



        private void button1_Click(object sender, EventArgs e)
        {
            DataTable dt = _dataService.getAll();

            if (dt.Rows.Count == 0)
            {
                MessageBox.Show("Aucune donnée trouvée pour l'exportation.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            int lignes = dt.Rows.Count;
            int cpt = 1;
            foreach (DataRow row in dt.Rows)
            {
                lcount.Visible = true;
                lcount.Text = cpt + " / " + lignes;
                cpt++;

                string trigramme = row["Trigramme"].ToString();

                _dataService = new service.DataService();
                getTrigrammeData(trigramme);
                bool success = _dataService.updatePv(trigramme, ab, cdef, g, montant, type.Trigram.observation);
            }

            DataTable dtUpdated = _dataService.getAll();
            dataGridView1.DataSource = dtUpdated;
            ExportToExcel(dataGridView1);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            ExportToPdf();
        }

        private bool isTrigramExport(){
            if (
                realisation.taux==0 &&
                realisation.contribution==0 &&
                realisation.dmt==0 &&
                realisation.qualite==0 &&
                realisation.quizz==0 &&
                realisation.t2bPersonalisation==0 &&
                realisation.t2bSolution==0
                )
            {
                return false;
            }
            return true;
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
                type.Trigram.observation = row["Observations"].ToString();

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
                    realisation.taux = taux == "" ? 0 : Math.Round(decimal.Parse(taux) * 100);
                    realisation.contribution = contribution == "" ? 0 : Math.Round(decimal.Parse(contribution) * 100);
                    realisation.dmt = dmt == "" ? 0 : decimal.Parse(dmt);
                    realisation.qualite = qualite == "" ? 0 : Math.Round(decimal.Parse(qualite) * 100);
                    realisation.quizz = quizz == "" ? 0 : Math.Round((decimal.Parse(quizz) * 20 / 100) * 100);
                    realisation.t2bPersonalisation = t2bPersonalisation == "" ? 0 : Math.Round(decimal.Parse(t2bPersonalisation) * 100);
                    realisation.t2bSolution = t2bSolution == "" ? 0 : Math.Round(decimal.Parse(t2bSolution) * 100);
                    realisation.notemanageriale = notemanageriale == "" ? 0 : Math.Round(decimal.Parse(notemanageriale) * 100);

                    ro.contribution = Math.Round((realisation.contribution / tContributionObjectif) * 100, 2);

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

                    this.ab = Math.Round(resultatPondere.contribution + resultatPondere.dmt, 2);
                    this.cdef = Math.Round(resultatPondere.qualite + resultatPondere.quizz + resultatPondere.t2bPersonalisation + resultatPondere.t2bSolution, 2);
                    this.g = realisation.notemanageriale;
                    this.atteinte = Math.Round((ab + cdef + g), 2);
                    this.baseNumeriale = this.tbaseNumeriale;
                    this.pvvFinal = Math.Round((ab / 100 + cdef / 100 + g / 100) * baseNumeriale, 2);
                    this.montant = this.tmontant;
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

            }
        }

        //private void getTrigrammeData(string trigramme)
        //{
        //    // Récupérer la ligne correspondant au trigramme
        //    DataRow row = _dataService.getByTrigramme(trigramme);

        //    if (row != null)
        //    {
        //        type.Trigram.trigramme = row["Trigramme"].ToString();
        //        type.Trigram.nom = row["Nom et prénoms"].ToString();
        //        type.Trigram.typeContrat = getTypeDeContrat(trigramme);
        //        type.Trigram.date = row["Date de prise d appel"].ToString();
        //        //type.Trigram.mois = countMonth(type.Trigram.date).ToString();
        //        type.Trigram.mois = row["Ancieneté en mois"].ToString();
        //        type.Trigram.superviseur = row["Superviseur"].ToString();
        //        type.Trigram.observation = row["Observations"].ToString();
        //        type.Trigram.idAgent = row["Id_agent"].ToString();

        //        string taux = row["Taux_d_absentéisme"].ToString().Split('%')[0].Replace('.', ',');
        //        string qualite = row["Qualité"].ToString().Split('%')[0].Replace('.', ',');
        //        string quizz = row["Quizz"].ToString().Split('%')[0].Replace('.', ',');
        //        string t2bPersonalisation = row["T2B_Personnalisation"].ToString().Split('%')[0].Replace('.', ',');
        //        string t2bSolution = row["T2B_Solution"].ToString().Split('%')[0].Replace('.', ',');
        //        string contribution = row["R/O en CI"].ToString().Split('%')[0].Replace('.', ',');
        //        string dmt = row["DMT"].ToString();
        //        string notemanageriale = row["Appréciation_managériale"].ToString().ToString().Split('%')[0].Replace('.', ',');
                
        //        try
        //        {
        //           realisation.taux = taux=="" ? 0 : decimal.Parse(taux);
        //            realisation.contribution = contribution == "" ? 0 : decimal.Parse(contribution);
        //            realisation.dmt = dmt == "" ? 0 : decimal.Parse(dmt);
        //            realisation.qualite = qualite == "" ? 0 : decimal.Parse(qualite);
        //            realisation.quizz = quizz == "" ? 0 : decimal.Parse(quizz);
        //            realisation.t2bPersonalisation = t2bPersonalisation == "" ? 0 : decimal.Parse(t2bPersonalisation);
        //            realisation.t2bSolution = t2bSolution == "" ? 0 : decimal.Parse(t2bSolution);
        //            realisation.notemanageriale = notemanageriale == "" ? 0 : decimal.Parse(notemanageriale);

        //            ro.contribution = Math.Round(realisation.contribution / tContributionObjectif, 2);

        //            _csvService = new service.CsvService("bareme.csv");
        //            string roDmt = _csvService.GetPourcentageBarem(realisation.dmt).Split('%')[0].Replace('.', ',');
        //            ro.dmt = roDmt == "" ? 0 : decimal.Parse(roDmt);

        //            ro.qualite = Math.Round(realisation.qualite / tQualiteObjectif, 2);
        //            ro.quizz = Math.Round(realisation.quizz / tQuizObjectif, 2);
        //            ro.t2bPersonalisation = Math.Round(realisation.t2bPersonalisation / tT2bPersObjectif, 2);
        //            ro.t2bSolution = Math.Round(realisation.t2bSolution / tT2bSolutionObjectif, 2);

        //            resultatPondere.contribution = formuleA(realisation.contribution, tContributionCoef, tContributionPoid, ro.contribution);
        //            resultatPondere.dmt = formuleB(tDmtPoid, ro.dmt);
        //            resultatPondere.qualite = formuleA(realisation.qualite, tQualiteCoef, tQualitePoid, ro.qualite);
        //            resultatPondere.quizz = formuleA(realisation.quizz, tQuizCoef, tQuizPoid, ro.quizz);
        //            resultatPondere.t2bPersonalisation = formuleA(realisation.t2bPersonalisation, tT2bPersCoef, tT2bPersPoid, ro.t2bPersonalisation);
        //            resultatPondere.t2bSolution = formuleA(realisation.t2bSolution, tT2bSolutionCoef, tT2bSolutionPoid, ro.t2bSolution);

        //            this.ab = Math.Round(resultatPondere.contribution + resultatPondere.dmt, 2);
        //            this.cdef = Math.Round(resultatPondere.qualite + resultatPondere.quizz + resultatPondere.t2bPersonalisation + resultatPondere.t2bSolution, 2);
        //            this.g = realisation.notemanageriale;
        //            this.atteinte = Math.Round((ab + cdef + g), 2);
        //            this.baseNumeriale = this.tbaseNumeriale;
        //            this.pvvFinal = Math.Round((ab + cdef + g) * baseNumeriale, 2);
        //            this.montant = this.tmontant;
        //        }
        //        catch (Exception e)
        //        {
        //            MessageBox.Show(e.Message);
        //            realisation.taux = 0;
        //            realisation.contribution = 0;
        //            realisation.dmt = 0;
        //            realisation.qualite = 0;
        //            realisation.quizz = 0;
        //            realisation.t2bPersonalisation = 0;
        //            realisation.t2bSolution = 0;

        //            ro.contribution = 0;
        //            ro.dmt = 0;
        //            ro.qualite = 0;
        //            ro.quizz = 0;
        //            ro.t2bPersonalisation = 0;
        //            ro.t2bSolution = 0;

        //            resultatPondere.contribution = 0;
        //            resultatPondere.dmt = 0;
        //            resultatPondere.qualite = 0;
        //            resultatPondere.quizz = 0;
        //            resultatPondere.t2bPersonalisation = 0;
        //            resultatPondere.t2bSolution = 0;
        //        }
        //    }
        //}

        private string getTypeDeContrat(string trigramme)
        {
            string[] tab = trigramme.Split('_');
            if (tab.Length > 1 && tab[1].ToLower() == "tmp")
            {
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

                return x - y;
            }
            catch (Exception)
            {
                return 0;
                //throw;
            }
        }

        private decimal formuleA(decimal realisation, decimal coef, decimal poid, decimal ro)
        {
            if (realisation >= coef)
            {
                decimal x = (poid * ro) / 100;
                return Math.Round(x, 2);
            }
            return 0;
        }

        private decimal formuleB(decimal poid, decimal ro)
        {
            return Math.Round((poid * ro) / 100, 2);
        }



        public void ExportToPdf()
        {
            DataTable dt = _dataService.getAll("Trigramme");

            if (dt.Rows.Count == 0)
            {
                MessageBox.Show("Aucune donnée trouvée pour l'exportation.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Fichiers PDF (*.pdf)|*.pdf";
            saveFileDialog.Title = "Enregistrer le PDF";
            saveFileDialog.FileName = "Commission " + dateTrigram + ".pdf";

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                string filePath = saveFileDialog.FileName;
                Document doc = new Document(PageSize.A4.Rotate());  // Orientation paysage

                try
                {
                    PdfWriter writer = PdfWriter.GetInstance(doc, new FileStream(filePath, FileMode.Create));
                    doc.Open();

                    iTextSharp.text.Font titleFont = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 16);
                    iTextSharp.text.Font boldFont = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 12);
                    iTextSharp.text.Font normalFont = FontFactory.GetFont(FontFactory.HELVETICA, 10);

                    int lignes = dt.Rows.Count;
                    int cpt = 1;
                    foreach (DataRow row in dt.Rows)
                    {
                        lcount.Visible = true;
                        lcount.Text = cpt + " / " + lignes;
                        cpt++;
                        // Toujours ajouter une nouvelle page avant d'ajouter un trigramme.
                        doc.NewPage(); // Saut de page forcé à chaque itération

                        string trigramme = row["Trigramme"].ToString();
                        getTrigrammeData(trigramme);

                        if (isTrigramExport())
                        {
                            if (idAgentSelected==type.Trigram.idAgent || idAgentSelected=="")
                            {
                                // Ajout du titre pour chaque trigramme
                                Paragraph title = new Paragraph("Commission " + trigramme + " en mois de " + this.dateTrigram + " \n", titleFont);
                                title.Alignment = Element.ALIGN_CENTER;
                                doc.Add(title);

                                // Ajout des informations du trigramme
                                doc.Add(new Paragraph("Nom et prénoms : " + type.Trigram.nom, normalFont));
                                doc.Add(new Paragraph("Id Agent : " + type.Trigram.idAgent, normalFont));
                                doc.Add(new Paragraph("Type de Contrat : " + type.Trigram.typeContrat, normalFont));
                                doc.Add(new Paragraph("Date de prise d'appel : " + type.Trigram.date, normalFont));
                                doc.Add(new Paragraph("Ancienneté en mois : " + type.Trigram.mois, normalFont));
                                doc.Add(new Paragraph("Superviseur : " + type.Trigram.superviseur + "\n", normalFont));

                                PdfPTable table = new PdfPTable(7);
                                table.WidthPercentage = 100; // Ajuste la largeur du tableau pour prendre toute la page
                                table.SetWidths(new float[] { 3f, 1f, 1f, 1f, 1f, 1f, 1f }); // Ajuste la largeur des colonnes

                                // Ajoutez une marge de 3px (en points, où 1px = 0.75 point)
                                table.SpacingBefore = 5f; // Définit une marge de 3 pixels (en points)

                                table.AddCell(new PdfPCell(new Phrase("Indicateurs", boldFont)) { HorizontalAlignment = Element.ALIGN_CENTER });
                                table.AddCell(new PdfPCell(new Phrase("Poids", boldFont)) { HorizontalAlignment = Element.ALIGN_CENTER });
                                table.AddCell(new PdfPCell(new Phrase("Objectif", boldFont)) { HorizontalAlignment = Element.ALIGN_CENTER });
                                table.AddCell(new PdfPCell(new Phrase("Coefficient", boldFont)) { HorizontalAlignment = Element.ALIGN_CENTER });
                                table.AddCell(new PdfPCell(new Phrase("Réalisation", boldFont)) { HorizontalAlignment = Element.ALIGN_CENTER });
                                table.AddCell(new PdfPCell(new Phrase("R/0", boldFont)) { HorizontalAlignment = Element.ALIGN_CENTER });
                                table.AddCell(new PdfPCell(new Phrase("Résultats pondérés", boldFont)) { HorizontalAlignment = Element.ALIGN_CENTER });

                                // Ajout des données dans le tableau
                                table.AddCell(new PdfPCell(new Phrase("Taux d'absentéisme")) { HorizontalAlignment = Element.ALIGN_LEFT });
                                table.AddCell(new PdfPCell(new Phrase("")) { HorizontalAlignment = Element.ALIGN_RIGHT });
                                table.AddCell(new PdfPCell(new Phrase("")) { HorizontalAlignment = Element.ALIGN_RIGHT });
                                table.AddCell(new PdfPCell(new Phrase("")) { HorizontalAlignment = Element.ALIGN_RIGHT });
                                table.AddCell(new PdfPCell(new Phrase(realisation.taux.ToString("F2") + "%")) { HorizontalAlignment = Element.ALIGN_RIGHT });
                                table.AddCell(new PdfPCell(new Phrase("")) { HorizontalAlignment = Element.ALIGN_RIGHT });
                                table.AddCell(new PdfPCell(new Phrase("")) { HorizontalAlignment = Element.ALIGN_RIGHT });

                                table.AddCell(new PdfPCell(new Phrase("Productivité (R/O en contribution individuelle aux appels traités)")) { HorizontalAlignment = Element.ALIGN_LEFT });
                                table.AddCell(new PdfPCell(new Phrase(tContributionPoid.ToString("F2") + "%")) { HorizontalAlignment = Element.ALIGN_RIGHT });
                                table.AddCell(new PdfPCell(new Phrase(tContributionObjectif.ToString("F2") + "%")) { HorizontalAlignment = Element.ALIGN_RIGHT });
                                table.AddCell(new PdfPCell(new Phrase(tContributionCoef.ToString("F2") + "%")) { HorizontalAlignment = Element.ALIGN_RIGHT });
                                table.AddCell(new PdfPCell(new Phrase(realisation.contribution.ToString("F2") + "%")) { HorizontalAlignment = Element.ALIGN_RIGHT });
                                table.AddCell(new PdfPCell(new Phrase(ro.contribution.ToString("F2") + "%")) { HorizontalAlignment = Element.ALIGN_RIGHT });
                                table.AddCell(new PdfPCell(new Phrase(resultatPondere.contribution.ToString("F2") + "%")) { HorizontalAlignment = Element.ALIGN_RIGHT });

                                table.AddCell(new PdfPCell(new Phrase("DMT")) { HorizontalAlignment = Element.ALIGN_LEFT });
                                table.AddCell(new PdfPCell(new Phrase(tDmtPoid.ToString("F2") + "%")) { HorizontalAlignment = Element.ALIGN_RIGHT });
                                table.AddCell(new PdfPCell(new Phrase(tDmtObjectif.ToString("F2"))) { HorizontalAlignment = Element.ALIGN_RIGHT });
                                table.AddCell(new PdfPCell(new Phrase(tDmtCoef.ToString("F2") + "%")) { HorizontalAlignment = Element.ALIGN_RIGHT });
                                table.AddCell(new PdfPCell(new Phrase(realisation.dmt.ToString("F2"))) { HorizontalAlignment = Element.ALIGN_RIGHT });
                                table.AddCell(new PdfPCell(new Phrase(ro.dmt.ToString("F2") + "%")) { HorizontalAlignment = Element.ALIGN_RIGHT });
                                table.AddCell(new PdfPCell(new Phrase(resultatPondere.dmt.ToString("F2") + "%")) { HorizontalAlignment = Element.ALIGN_RIGHT });

                                table.AddCell(new PdfPCell(new Phrase("Evaluation qualitative (note moyenne)")) { HorizontalAlignment = Element.ALIGN_LEFT });
                                table.AddCell(new PdfPCell(new Phrase(tQualitePoid.ToString("F2") + "%")) { HorizontalAlignment = Element.ALIGN_RIGHT });
                                table.AddCell(new PdfPCell(new Phrase(tQualiteObjectif.ToString("F2") + "%")) { HorizontalAlignment = Element.ALIGN_RIGHT });
                                table.AddCell(new PdfPCell(new Phrase(tQualiteCoef.ToString("F2") + "%")) { HorizontalAlignment = Element.ALIGN_RIGHT });
                                table.AddCell(new PdfPCell(new Phrase(realisation.qualite.ToString("F2") + "%")) { HorizontalAlignment = Element.ALIGN_RIGHT });
                                table.AddCell(new PdfPCell(new Phrase(ro.qualite.ToString("F2") + "%")) { HorizontalAlignment = Element.ALIGN_RIGHT });
                                table.AddCell(new PdfPCell(new Phrase(resultatPondere.qualite.ToString("F2") + "%")) { HorizontalAlignment = Element.ALIGN_RIGHT });

                                table.AddCell(new PdfPCell(new Phrase("Quizz")) { HorizontalAlignment = Element.ALIGN_LEFT });
                                table.AddCell(new PdfPCell(new Phrase(tQuizPoid.ToString("F2") + "%")) { HorizontalAlignment = Element.ALIGN_RIGHT });
                                table.AddCell(new PdfPCell(new Phrase(tQuizObjectif.ToString("F2"))) { HorizontalAlignment = Element.ALIGN_RIGHT });
                                table.AddCell(new PdfPCell(new Phrase(tQuizCoef.ToString("F2"))) { HorizontalAlignment = Element.ALIGN_RIGHT });
                                table.AddCell(new PdfPCell(new Phrase(realisation.quizz.ToString("F2"))) { HorizontalAlignment = Element.ALIGN_RIGHT });
                                table.AddCell(new PdfPCell(new Phrase(ro.quizz.ToString("F2") + "%")) { HorizontalAlignment = Element.ALIGN_RIGHT });
                                table.AddCell(new PdfPCell(new Phrase(resultatPondere.quizz.ToString("F2") + "%")) { HorizontalAlignment = Element.ALIGN_RIGHT });

                                table.AddCell(new PdfPCell(new Phrase("Satisfaction client  sur la personnalisation du traitement  (TTB)")) { HorizontalAlignment = Element.ALIGN_LEFT });
                                table.AddCell(new PdfPCell(new Phrase(tT2bPersPoid.ToString("F2") + "%")) { HorizontalAlignment = Element.ALIGN_RIGHT });
                                table.AddCell(new PdfPCell(new Phrase(tT2bPersObjectif.ToString("F2") + "%")) { HorizontalAlignment = Element.ALIGN_RIGHT });
                                table.AddCell(new PdfPCell(new Phrase(tT2bPersCoef.ToString("F2") + "%")) { HorizontalAlignment = Element.ALIGN_RIGHT });
                                table.AddCell(new PdfPCell(new Phrase(realisation.t2bPersonalisation.ToString("F2") + "%")) { HorizontalAlignment = Element.ALIGN_RIGHT });
                                table.AddCell(new PdfPCell(new Phrase(ro.t2bPersonalisation.ToString("F2") + "%")) { HorizontalAlignment = Element.ALIGN_RIGHT });
                                table.AddCell(new PdfPCell(new Phrase(resultatPondere.t2bPersonalisation.ToString("F2") + "%")) { HorizontalAlignment = Element.ALIGN_RIGHT });

                                table.AddCell(new PdfPCell(new Phrase("Satisfaction client sur solution proposée (TTB)")) { HorizontalAlignment = Element.ALIGN_LEFT });
                                table.AddCell(new PdfPCell(new Phrase(tT2bSolutionPoid.ToString("F2") + "%")) { HorizontalAlignment = Element.ALIGN_RIGHT });
                                table.AddCell(new PdfPCell(new Phrase(tT2bSolutionObjectif.ToString("F2") + "%")) { HorizontalAlignment = Element.ALIGN_RIGHT });
                                table.AddCell(new PdfPCell(new Phrase(tT2bSolutionCoef.ToString("F2") + "%")) { HorizontalAlignment = Element.ALIGN_RIGHT });
                                table.AddCell(new PdfPCell(new Phrase(realisation.t2bSolution.ToString("F2") + "%")) { HorizontalAlignment = Element.ALIGN_RIGHT });
                                table.AddCell(new PdfPCell(new Phrase(ro.t2bSolution.ToString("F2") + "%")) { HorizontalAlignment = Element.ALIGN_RIGHT });
                                table.AddCell(new PdfPCell(new Phrase(resultatPondere.t2bSolution.ToString("F2") + "%")) { HorizontalAlignment = Element.ALIGN_RIGHT });

                                table.AddCell(new PdfPCell(new Phrase("Appréciation managériale ")) { HorizontalAlignment = Element.ALIGN_LEFT });
                                table.AddCell(new PdfPCell(new Phrase(tNotePoid.ToString("F2") + "%")) { HorizontalAlignment = Element.ALIGN_RIGHT });
                                table.AddCell(new PdfPCell(new Phrase("")) { HorizontalAlignment = Element.ALIGN_RIGHT });
                                table.AddCell(new PdfPCell(new Phrase("")) { HorizontalAlignment = Element.ALIGN_RIGHT });
                                table.AddCell(new PdfPCell(new Phrase(realisation.notemanageriale.ToString("F2") + "%")) { HorizontalAlignment = Element.ALIGN_RIGHT });
                                table.AddCell(new PdfPCell(new Phrase("")) { HorizontalAlignment = Element.ALIGN_RIGHT });
                                table.AddCell(new PdfPCell(new Phrase(realisation.notemanageriale.ToString("F2") + "%")) { HorizontalAlignment = Element.ALIGN_RIGHT });

                                // Ajoute le tableau au document
                                doc.Add(table);

                                // Ajout du score final
                                doc.Add(new Paragraph("PV1 : " + ab + "%", normalFont));
                                doc.Add(new Paragraph("PV2 : " + cdef + "%", normalFont));
                                doc.Add(new Paragraph("PV3 : " + g + "%", normalFont));
                                doc.Add(new Paragraph("% d'atteinte des objectifs : " + atteinte + "%", normalFont));
                                doc.Add(new Paragraph("Base numéraire : " + baseNumeriale.ToString("#,##0") + " Ar", normalFont));
                                doc.Add(new Paragraph("PVV final : " + pvvFinal.ToString("#,##0") + " Ar", normalFont));
                                doc.Add(new Paragraph("Montant commission en Ar arrondi : " + montant.ToString("#,##0") + " Ar", normalFont));
                                doc.Add(new Paragraph("Observation : " + type.Trigram.observation, normalFont));   
                            }
                        }
                    }

                    // Fermeture du document et de l'écrivain
                    doc.Close();
                    writer.Close();
                    MessageBox.Show("Exportation réussie !", "Succès", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    this.Close();
                }
                catch (Exception ex)
                {
                    //MessageBox.Show("Erreur lors de la génération du PDF : " + ex.Message, "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void ExportType_Load(object sender, EventArgs e)
        {

        }
    }
}

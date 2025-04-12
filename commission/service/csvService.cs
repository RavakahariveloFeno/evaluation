using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;

namespace evaluation.service
{
    class CsvService
    {
        private string filePath = "";

        public CsvService(string table = "csv.csv")
        {
            this.filePath = "data/" + table;
        }

        // 📌 Lire les valeurs du fichier CSV
        public decimal[] ReadValues()
        {
            if (File.Exists(filePath))
            {
                string[] lines = File.ReadAllLines(filePath);
                if (lines.Length > 0)
                {
                    string[] values = lines[0].Split(';'); // Séparation par ";"
                    if (values.Length == 2)
                    {
                        decimal[] response = new decimal[2];
                        response[0] = Convert.ToDecimal(values[0]);
                        response[1] = Convert.ToDecimal(values[1]);
                        return response;
                    }
                }
            }
            // Retourne un tableau vide au lieu de `null`
            return new decimal[0];
        }

        // 📌 Modifier et enregistrer de nouvelles valeurs dans le fichier CSV
        public void UpdateValues(decimal valeur1, decimal valeur2)
        {
            File.WriteAllText(filePath, valeur1+";"+valeur2);
        }

        public List<type.Bareme> ReadBarem()
        {
            List<type.Bareme> dataList = new List<type.Bareme>();
            try
            {
                using (StreamReader reader = new StreamReader(filePath))
                {
                    string line;
                    while ((line = reader.ReadLine()) != null)
                    {
                        string[] values = line.Split(';');

                        if (values.Length == 3) // Vérifie qu'on a bien 3 colonnes
                        {
                            dataList.Add(new type.Bareme
                            {
                                Pourcentage = values[0],
                                Min = values[1],
                                Max = values[2]
                            });
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Erreur lors de la lecture du fichier : " + ex.Message);
            }

            return dataList;
        }

        public void SaveToCsv(List<type.Bareme> dataList)
        {
            try
            {
                using (StreamWriter writer = new StreamWriter(filePath))
                {
                    foreach (var data in dataList)
                    {
                        writer.WriteLine(data.Pourcentage+";"+data.Min+";"+data.Max);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erreur lors de l'enregistrement : " + ex.Message, "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public string GetPourcentageBarem(decimal x)
        {
            List<type.Bareme> dataList = this.ReadBarem();
            foreach (var item in dataList)
            {
                decimal min;
                decimal max;
                if (item.Min == "min")
                {
                    min = decimal.MinValue;
                }
                else {
                    min = decimal.Parse(item.Min);
                }
                if (item.Max == "max")
                {
                    max = decimal.MaxValue;
                }
                else {
                    max = decimal.Parse(item.Max);
                }
                
                if (x >= min && x <= max)
                {
                    return item.Pourcentage;
                }
            }
            return "Valeur hors barème";
        }
    }
}

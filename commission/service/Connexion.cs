
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace evaluation.service
{
    class Connexion
    {
        string filePath = "data\\MATRICE_PVV.xlsx";
        OleDbConnection connection;

        public Connexion(string table)
        {
            filePath = "data\\" + table + ".xlsx";

            string connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath + ";Extended Properties=\"Excel 12.0 Xml;HDR=YES;\"";
            connection = new OleDbConnection(connectionString);
        }

        public DataTable getAll(string colonne = "*") {
            try
            {
                // Ouverture de la connexion
                connection.Open();

                // Récupération du nom de la première feuille
                DataTable sheets = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                if (sheets == null || sheets.Rows.Count == 0)
                {
                    MessageBox.Show("Aucune feuille trouvée dans le fichier Excel.");
                    return null;
                }

                string sheetName = sheets.Rows[0]["TABLE_NAME"].ToString(); // Nom de la première feuille

                // Création de l'adaptateur de données
                OleDbDataAdapter adapter = new OleDbDataAdapter("SELECT "+colonne+" FROM [" + sheetName + "]", connection);
                DataTable dt = new DataTable();

                // Remplissage du DataTable
                adapter.Fill(dt);

                // Affichage dans le DataGridView
                return dt;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erreur : " + ex.Message);
                return null;
            }
            finally
            {
                // Fermeture de la connexion dans tous les cas
                if (connection.State == ConnectionState.Open)
                {
                    connection.Close();
                }
            }
        }

        public DataRow getByTrigramme(string trigramme)
        {
            try
            {
                // Ouverture de la connexion
                connection.Open();

                // Récupération du nom de la première feuille
                DataTable sheets = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                if (sheets == null || sheets.Rows.Count == 0)
                {
                    MessageBox.Show("Aucune feuille trouvée dans le fichier Excel.");
                    return null;
                }

                string sheetName = sheets.Rows[0]["TABLE_NAME"].ToString(); // Nom de la première feuille

                // Création de l'adaptateur de données pour récupérer toute la table
                OleDbDataAdapter adapter = new OleDbDataAdapter("SELECT * FROM [" + sheetName + "]", connection);
                DataTable dt = new DataTable();

                // Remplissage du DataTable
                adapter.Fill(dt);

                // Rechercher la ligne correspondant au trigramme
                var rows = dt.AsEnumerable().Where(row => row.Field<string>("Trigramme") == trigramme).ToList();

                // Si une ligne est trouvée, retourner cette ligne
                if (rows.Count > 0)
                {
                    return rows[0]; // Retourner la première ligne correspondante
                }
                else
                {
                    return null;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erreur : " + ex.Message);
                return null;
            }
            finally
            {
                // Fermeture de la connexion dans tous les cas
                if (connection.State == ConnectionState.Open)
                {
                    connection.Close();
                }
            }
        }

        public DataRow getByIndicateur(string id)
        {
            try
            {
                // Ouverture de la connexion
                connection.Open();

                // Récupération du nom de la première feuille
                DataTable sheets = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                if (sheets == null || sheets.Rows.Count == 0)
                {
                    MessageBox.Show("Aucune feuille trouvée dans le fichier Excel.");
                    return null;
                }

                string sheetName = sheets.Rows[0]["TABLE_NAME"].ToString(); // Nom de la première feuille

                // Création de l'adaptateur de données pour récupérer toute la table
                OleDbDataAdapter adapter = new OleDbDataAdapter("SELECT * FROM [" + sheetName + "]", connection);
                DataTable dt = new DataTable();

                // Remplissage du DataTable
                adapter.Fill(dt);

                // Rechercher la ligne correspondant au trigramme
                var rows = dt.AsEnumerable().Where(row => row.Field<string>("id") == id).ToList();

                // Si une ligne est trouvée, retourner cette ligne
                if (rows.Count > 0)
                {
                    return rows[0]; // Retourner la première ligne correspondante
                }
                else
                {
                    return null;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erreur : " + ex.Message);
                return null;
            }
            finally
            {
                // Fermeture de la connexion dans tous les cas
                if (connection.State == ConnectionState.Open)
                {
                    connection.Close();
                }
            }
        }

        public bool updateIndicateur(string id, int poid, int objectif, int coef)
        {
            try
            {
                // Ouvrir la connexion
                connection.Open();

                // Récupération du nom de la première feuille
                DataTable sheets = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                if (sheets == null || sheets.Rows.Count == 0)
                {
                    MessageBox.Show("Aucune feuille trouvée dans le fichier Excel.");
                    return false;
                }

                string sheetName = sheets.Rows[0]["TABLE_NAME"].ToString(); // Nom de la première feuille

                // Requête UPDATE pour modifier la valeur de la colonne correspondante
                string query = "UPDATE [" + sheetName + "] SET Poids = ?, Objectif = ?, Coefficient = ? WHERE id = ?";

                using (OleDbCommand cmd = new OleDbCommand(query, connection))
                {
                    // Ajouter les paramètres
                    cmd.Parameters.AddWithValue("@Poids", poid);
                    cmd.Parameters.AddWithValue("@Objectif", objectif);
                    cmd.Parameters.AddWithValue("@Coefficient", coef);
                    cmd.Parameters.AddWithValue("@id", id);

                    // Exécuter la requête
                    int rowsAffected = cmd.ExecuteNonQuery();

                    // Vérifier si la mise à jour a été effectuée
                    return rowsAffected > 0;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erreur lors de la mise à jour : " + ex.Message);
                return false;
            }
            finally
            {
                // Fermer la connexion
                if (connection.State == ConnectionState.Open)
                {
                    connection.Close();
                }
            }
        }

        public bool updatePv(string id, decimal pv1, decimal pv2, decimal pv3,decimal montant,string observation)
        {
            try
            {
                // Ouvrir la connexion
                connection.Open();

                // Récupération du nom de la première feuille
                DataTable sheets = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                if (sheets == null || sheets.Rows.Count == 0)
                {
                    MessageBox.Show("Aucune feuille trouvée dans le fichier Excel.");
                    return false;
                }

                string sheetName = sheets.Rows[0]["TABLE_NAME"].ToString(); // Nom de la première feuille

                // Requête UPDATE pour modifier la valeur de la colonne correspondante
                string query = "UPDATE [" + sheetName + "] SET PV1 = ?,PV2 = ?,PV3 = ?,Montant_commission= ?,observations=? WHERE Trigramme = ?";

                using (OleDbCommand cmd = new OleDbCommand(query, connection))
                {
                    // Ajouter les paramètres
                    cmd.Parameters.AddWithValue("@PV1", pv1);
                    cmd.Parameters.AddWithValue("@PV2", pv2);
                    cmd.Parameters.AddWithValue("@PV3", pv3);
                    cmd.Parameters.AddWithValue("@Montant_commission", montant.ToString());
                    cmd.Parameters.AddWithValue("@observations", observation);
                    cmd.Parameters.AddWithValue("@Trigramme", id);

                    // Exécuter la requête
                    int rowsAffected = cmd.ExecuteNonQuery();

                    // Vérifier si la mise à jour a été effectuée
                    return rowsAffected > 0;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erreur lors de la mise à jour : " + ex.Message);
                return false;
            }
            finally
            {
                // Fermer la connexion
                if (connection.State == ConnectionState.Open)
                {
                    connection.Close();
                }
            }
        }
    }
}

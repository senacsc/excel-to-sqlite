using System;
using System.Collections.Generic;
using System.IO;
using System.Windows;
using NPOI.HSSF.UserModel;
using Microsoft.Win32;
using System.Data.SQLite;
using System.Globalization;
using NPOI.SS.UserModel;

namespace Auphan_Converter
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private readonly OpenFileDialog _dialog;
        private List<Medicamento> _medicamentos;
        private List<ClasseTerapeutica> _classesTerapeuticas;

        public MainWindow()
        {
            InitializeComponent();

            _dialog = new OpenFileDialog
            {
                Multiselect = false,
                CheckFileExists = true,
                CheckPathExists = true,
                ShowReadOnly = true,
                AddExtension = true,
                Filter = "Excel Files|*.xls;*.xlsx"
            };

            FileTextBox.IsReadOnly = true;
        }

        class Planilha
        {
            public ISheet Value { get; }

            public Planilha(ISheet sheet)
            {
                Value = sheet;
            }

            public override string ToString()
            {
                return Value.SheetName;
            }
        }

        private void SetupSheetsList(HSSFWorkbook wb)
        {
            SheetsComboBox.Items.Clear();

            for (int i = 0; i < wb.Count; i++)
                SheetsComboBox.Items.Add(new Planilha(wb.GetSheetAt(i)));

            if (SheetsComboBox.HasItems)
                SheetsComboBox.SelectedIndex = 0;
        }

        class ClasseTerapeutica
        {
            public double Id { get; set; }

            public string Nome { get; set; }

            public string Descricao { get; set; }

            public override string ToString()
            {
                return $"[ Id: {Id}, Nome: {Nome}, Descricao: {Descricao} ]";
            }

            protected bool Equals(ClasseTerapeutica other)
            {
                return Id.Equals(other.Id);
            }

            public override bool Equals(object obj)
            {
                if (ReferenceEquals(null, obj)) return false;
                if (ReferenceEquals(this, obj)) return true;
                if (obj.GetType() != this.GetType()) return false;
                return Equals((ClasseTerapeutica) obj);
            }

            public override int GetHashCode()
            {
                return Id.GetHashCode();
            }
        }

        class Medicamento
        {
            public ClasseTerapeutica ClasseTerapeutica { get; set; }

            public double IdMedicamento { get; set; }

            public string NomeMedicamento { get; set; }

            public string FormaApresentacao { get; set; }

            public string Narrativa { get; set; }

            public override string ToString()
            {
                return $"ClasseTerapeutica: {ClasseTerapeutica}, IdMedicamento: {IdMedicamento}, NomeMedicamento: {NomeMedicamento}, FormaApresentacao: {FormaApresentacao}, Narrativa: {Narrativa}";
            }
        }


        private void StartEditing(string filename)
        {
            var wb = OpenExcelFile(filename);
            SetupSheetsList(wb);

            ReloadSheet();
        }

        private void ReloadSheet()
        {
            var planilha = (SheetsComboBox.SelectedItem as Planilha)?.Value;

            if (planilha != null)
            {
                IRow labelRow = planilha.GetRow(0);

                _medicamentos = new List<Medicamento>();
                _classesTerapeuticas = new List<ClasseTerapeutica>();

                for (int i = 1; i < planilha.LastRowNum - 1; i++)
                {
                    IRow nRow = planilha.GetRow(i);

                    double idMedicamento = nRow.GetCell(0).NumericCellValue;
                    string nomeMedicamento = nRow.GetCell(1).StringCellValue;
                    string formaApresentacao = nRow.GetCell(3).StringCellValue;
                    //string narrativa = nRow.GetCell(28).StringCellValue;
                    double classeId = nRow.GetCell(7).NumericCellValue;
                    string classeNome = nRow.GetCell(27).StringCellValue;
                    string classeDescricao = nRow.GetCell(8).StringCellValue;

                    ClasseTerapeutica classe = new ClasseTerapeutica()
                    {
                        Id = classeId,
                        Nome = classeNome.Trim(),
                        Descricao = classeDescricao.Trim()
                    };

                    if (!_classesTerapeuticas.Contains(classe))
                    {
                        _classesTerapeuticas.Add(classe);
                    }

                    _medicamentos.Add(new Medicamento
                    {
                        IdMedicamento = idMedicamento,
                        NomeMedicamento = nomeMedicamento.Trim(),
                        FormaApresentacao = formaApresentacao.Trim(),
                        ClasseTerapeutica = classe
                    });
                }

                MedicamentosListView.Items.Clear();

                foreach (Medicamento m in _medicamentos)
                {
                    MedicamentosListView.Items.Add(m.ToString());
                }
            }
        }

        private void button_Click(object sender, RoutedEventArgs e)
        {
            var result = _dialog.ShowDialog();

            if (result == true)
            {
                FileTextBox.Text = _dialog.FileName;
                StartEditing(_dialog.FileName);
            }
        }

        private HSSFWorkbook OpenExcelFile(string filepath)
        {
            var file = new FileStream(filepath, FileMode.Open, FileAccess.Read);
            return new HSSFWorkbook(file);
        }

        private void button_Click_1(object sender, RoutedEventArgs e)
        {
            OpenFileDialog saveDialog = new OpenFileDialog
            {
                Multiselect = false,
                CheckFileExists = false,
                CheckPathExists = true,
                ShowReadOnly = true,
                AddExtension = true,
                Filter = "SQLite|*.db"
            };

            bool? result = saveDialog.ShowDialog();

            if (result == true)
            {
                if (MessageBox.Show($"Deseja salvar o banco de dados em {saveDialog.FileName}?", "Exportação para SQLite", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                {
                    MessageBox.Show("O processo pode demorar um pouco, clique em OK para continuar");

                    if (!File.Exists(saveDialog.FileName))
                    {
                        SQLiteConnection.CreateFile(saveDialog.FileName);
                    }

                    var dbConnection = new SQLiteConnection($"Data Source={saveDialog.FileName};Version=3;");
                    dbConnection.Open();

                    string sql = "";

                    sql = "CREATE TABLE IF NOT EXISTS classe_terapeutica (idClasse INT(7), nomeClasse VARCHAR(200), descricaoClasse VARCHAR(300))";
                    new SQLiteCommand(sql, dbConnection).ExecuteNonQuery();

                    sql = "CREATE TABLE IF NOT EXISTS medicamento (idMedicamento INT(7), nomeMedicamento VARCHAR(150), formaApresentacao VARCHAR(100), classeTerapeutica INT (7))";
                    new SQLiteCommand(sql, dbConnection).ExecuteNonQuery();

                    foreach (ClasseTerapeutica ct in _classesTerapeuticas)
                    {
                        sql = "insert into classe_terapeutica (idClasse, nomeClasse, descricaoClasse) values (@idClasse, @nomeClasse, @descricaoClasse)";

                        SQLiteCommand command = new SQLiteCommand(sql, dbConnection);
                        command.Parameters.AddWithValue("@idClasse", ct.Id);
                        command.Parameters.AddWithValue("@nomeClasse", ct.Nome);
                        command.Parameters.AddWithValue("@descricaoClasse", ct.Descricao);

                        try
                        {
                            command.ExecuteNonQuery();
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.ToString());
                        }
                    }

                    dbConnection.Close();

                    dbConnection.Open();

                    foreach (Medicamento m in _medicamentos)
                    {
                        sql = "insert into medicamento (idMedicamento, nomeMedicamento, formaApresentacao, classeTerapeutica) values (@idMedicamento, @nomeMedicamento, @formaApresentacao, @classeTerapeutica)";

                        SQLiteCommand command = new SQLiteCommand(sql, dbConnection);
                        command.Parameters.AddWithValue("@idMedicamento", m.IdMedicamento);
                        command.Parameters.AddWithValue("@nomeMedicamento", m.NomeMedicamento);
                        command.Parameters.AddWithValue("@formaApresentacao", m.FormaApresentacao);
                        command.Parameters.AddWithValue("@classeTerapeutica", m.ClasseTerapeutica.Id);

                        try
                        {
                            command.ExecuteNonQuery();
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.ToString());
                        }
                    }

                    MessageBox.Show("Exportação concluida com sucesso!");
                }
            }
        }

        private void SheetsComboBox_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            var planilha = SheetsComboBox.SelectedItem as Planilha;

            if (planilha?.Value.SheetName.ToUpper().Equals("PLANEJADOS") == true)
            {
                ReloadSheet();
            }
            else
            {
                MessageBox.Show("Planilha não suportada.");
                SheetsComboBox.SelectedIndex = 0;
            }
        }
    }
}

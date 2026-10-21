using System;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace DergesaReports
{
    public class MainForm : Form
    {
        private TextBox dbPath;
        private DateTimePicker fromDate, toDate;
        private ComboBox articleCombo, reportCombo;
        private DataGridView grid;
        private Label status;
        private string ConnectionString
        {
            get { return "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + dbPath.Text.Trim() + ";Persist Security Info=False;"; }
        }

        public MainForm()
        {
            Text = "Dergesa Reports - Hyrje / Dalje / Gjendje";
            StartPosition = FormStartPosition.CenterScreen;
            MinimumSize = new Size(950, 600);
            Size = new Size(1180, 740);
            Font = new Font("Segoe UI", 9F);
            BuildUi();
        }

        private void BuildUi()
        {
            var header = new Panel { Dock = DockStyle.Top, Height = 76, BackColor = Color.FromArgb(28, 55, 87) };
            var title = new Label { Text = "RAPORTET E MAGAZINËS", ForeColor = Color.White, Font = new Font("Segoe UI", 18, FontStyle.Bold), AutoSize = true, Location = new Point(20, 12) };
            var subtitle = new Label { Text = "Hyrje-dalje sipas artikujve • Gjendjet • Kartela e artikullit", ForeColor = Color.Gainsboro, AutoSize = true, Location = new Point(23, 47) };
            header.Controls.Add(title); header.Controls.Add(subtitle); Controls.Add(header);

            var dbLabel = new Label { Text = "Databaza Access:", AutoSize = true, Location = new Point(20, 98) };
            dbPath = new TextBox { Location = new Point(135, 94), Width = 760, Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right };
            var browse = new Button { Text = "Zgjidh MDB...", Location = new Point(905, 92), Width = 120, Anchor = AnchorStyles.Top | AnchorStyles.Right };
            var connect = new Button { Text = "Lidhu / Ngarko", Location = new Point(1035, 92), Width = 120, Anchor = AnchorStyles.Top | AnchorStyles.Right };
            browse.Click += BrowseDb; connect.Click += (s, e) => LoadArticles();
            Controls.Add(dbLabel); Controls.Add(dbPath); Controls.Add(browse); Controls.Add(connect);

            var filters = new GroupBox { Text = "Filtrat", Location = new Point(20, 132), Size = new Size(1135, 105), Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right };
            filters.Controls.Add(new Label { Text = "Nga data:", AutoSize = true, Location = new Point(15, 28) });
            fromDate = new DateTimePicker { Format = DateTimePickerFormat.Short, Location = new Point(85, 24), Width = 130 };
            filters.Controls.Add(fromDate);
            filters.Controls.Add(new Label { Text = "Deri më:", AutoSize = true, Location = new Point(235, 28) });
            toDate = new DateTimePicker { Format = DateTimePickerFormat.Short, Location = new Point(295, 24), Width = 130, Value = DateTime.Today };
            fromDate.Value = new DateTime(DateTime.Today.Year, DateTime.Today.Month, 1);
            filters.Controls.Add(toDate);
            filters.Controls.Add(new Label { Text = "Artikulli:", AutoSize = true, Location = new Point(445, 28) });
            articleCombo = new ComboBox { Location = new Point(500, 24), Width = 360, DropDownStyle = ComboBoxStyle.DropDownList, DisplayMember = "Display", ValueMember = "Idart" };
            filters.Controls.Add(articleCombo);
            filters.Controls.Add(new Label { Text = "Raporti:", AutoSize = true, Location = new Point(15, 65) });
            reportCombo = new ComboBox { Location = new Point(85, 61), Width = 340, DropDownStyle = ComboBoxStyle.DropDownList };
            reportCombo.Items.AddRange(new object[] { "Hyrje dhe dalje sipas artikullit", "Gjendja aktuale e artikujve", "Kartela e artikullit" });
            reportCombo.SelectedIndex = 0;
            filters.Controls.Add(reportCombo);
            var run = new Button { Text = "Shfaq raportin", Location = new Point(445, 58), Width = 150, Height = 30 };
            run.Click += (s, e) => RunReport();
            var export = new Button { Text = "Eksporto CSV", Location = new Point(605, 58), Width = 130, Height = 30 };
            export.Click += ExportCsv;
            var print = new Button { Text = "Printo", Location = new Point(745, 58), Width = 100, Height = 30 };
            print.Click += (s, e) => MessageBox.Show("Për printim, përdor Eksporto CSV dhe hape skedarin në Excel. Print Preview do të shtohet në versionin pasues.", "Printim");
            filters.Controls.Add(run); filters.Controls.Add(export); filters.Controls.Add(print);
            Controls.Add(filters);

            grid = new DataGridView { Location = new Point(20, 250), Size = new Size(1135, 390), Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right, ReadOnly = true, AllowUserToAddRows = false, AllowUserToDeleteRows = false, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells, BackgroundColor = Color.White, RowHeadersVisible = false };
            Controls.Add(grid);
            status = new Label { Text = "Zgjidh databazën .mdb dhe kliko Lidhu / Ngarko.", Dock = DockStyle.Bottom, Height = 26, TextAlign = ContentAlignment.MiddleLeft, Padding = new Padding(8, 0, 0, 0) };
            Controls.Add(status);
            dbPath.Text = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "dergesa.mdb");
        }

        private void BrowseDb(object sender, EventArgs e)
        {
            using (var dlg = new OpenFileDialog())
            {
                dlg.Filter = "Access Database (*.mdb;*.accdb)|*.mdb;*.accdb";
                dlg.Title = "Zgjidh databazën Access";
                if (dlg.ShowDialog() == DialogResult.OK) { dbPath.Text = dlg.FileName; LoadArticles(); }
            }
        }

        private OleDbConnection OpenConnection()
        {
            if (!File.Exists(dbPath.Text.Trim())) throw new FileNotFoundException("Nuk u gjet databaza: " + dbPath.Text);
            return new OleDbConnection(ConnectionString);
        }

        private void LoadArticles()
        {
            try
            {
                using (var cn = OpenConnection())
                using (var da = new OleDbDataAdapter("SELECT Idart, Kodi, Artikulli FROM Artikujt ORDER BY Artikulli", cn))
                {
                    var dt = new DataTable(); da.Fill(dt);
                    var all = dt.NewRow(); all["Idart"] = 0; all["Kodi"] = ""; all["Artikulli"] = "(Të gjithë artikujt)"; dt.Rows.InsertAt(all, 0);
                    dt.Columns.Add("Display", typeof(string), "IIF(Kodi = '', Artikulli, Kodi + ' - ' + Artikulli)");
                    articleCombo.DataSource = dt;
                    articleCombo.SelectedIndex = 0;
                    status.Text = "Lidhja u krye. U ngarkuan " + (dt.Rows.Count - 1) + " artikuj.";
                }
            }
            catch (Exception ex) { ShowDbError(ex); }
        }

        private void RunReport()
        {
            if (toDate.Value.Date < fromDate.Value.Date) { MessageBox.Show("Data 'Deri më' nuk mund të jetë më e vogël se data 'Nga'."); return; }
            try
            {
                DataTable dt;
                int report = reportCombo.SelectedIndex;
                int id = articleCombo.SelectedValue == null ? 0 : Convert.ToInt32(articleCombo.SelectedValue);
                using (var cn = OpenConnection())
                {
                    cn.Open();
                    if (report == 0) dt = GetMovementSummary(cn, id);
                    else if (report == 1) dt = GetStock(cn, id);
                    else
                    {
                        if (id == 0) { MessageBox.Show("Për kartelën, zgjidh një artikull të caktuar."); return; }
                        dt = GetCard(cn, id);
                    }
                }
                grid.DataSource = dt;
                status.Text = "Rreshta: " + dt.Rows.Count + " | Raporti: " + reportCombo.Text;
            }
            catch (Exception ex) { ShowDbError(ex); }
        }

        private DataTable Query(OleDbConnection cn, string sql, params OleDbParameter[] parameters)
        {
            using (var cmd = new OleDbCommand(sql, cn))
            {
                if (parameters != null) cmd.Parameters.AddRange(parameters);
                using (var da = new OleDbDataAdapter(cmd))
                {
                    var dt = new DataTable(); da.Fill(dt); return dt;
                }
            }
        }

        private DataTable GetMovementSummary(OleDbConnection cn, int id)
        {
            string sql = @"
SELECT A.Idart, A.Kodi, A.Artikulli,
 Nz((SELECT Sum(L.SasiHyrje) FROM FaturehyrjeArtikujt AS L
 INNER JOIN Faturehyrje AS F ON L.idfaturehyrje=F.Idfaturehyrje
 WHERE L.idart=A.Idart AND F.Datafatures>=? AND F.Datafatures<? AND Nz(F.Anulluar,False)=False),0) AS Hyrje,
 Nz((SELECT Sum(L.SasiShitje) FROM FatureshitjeArtikujt AS L
 INNER JOIN FatureShitje AS F ON L.idfatureshitje=F.IdfatureShitje
 WHERE L.idart=A.Idart AND F.Datafatures>=? AND F.Datafatures<? AND Nz(F.Anulluar,False)=False),0) AS Dalje
 FROM Artikujt AS A WHERE (?=0 OR A.Idart=?) ORDER BY A.Artikulli";
            return Query(cn, sql,
                new OleDbParameter { OleDbType = OleDbType.Date, Value = fromDate.Value.Date },
                new OleDbParameter { OleDbType = OleDbType.Date, Value = toDate.Value.Date.AddDays(1) },
                new OleDbParameter { OleDbType = OleDbType.Date, Value = fromDate.Value.Date },
                new OleDbParameter { OleDbType = OleDbType.Date, Value = toDate.Value.Date.AddDays(1) },
                new OleDbParameter { OleDbType = OleDbType.Integer, Value = id },
                new OleDbParameter { OleDbType = OleDbType.Integer, Value = id });
        }

        private DataTable GetStock(OleDbConnection cn, int id)
        {
            string sql = "SELECT Idart, Kodi, Barcode, Artikulli, Njesia, Gjendjemalli AS Gjendja, Cmimihyrjes, CmimiShitjes FROM Artikujt WHERE (?=0 OR Idart=?) ORDER BY Artikulli";
            return Query(cn, sql, new OleDbParameter { OleDbType = OleDbType.Integer, Value = id }, new OleDbParameter { OleDbType = OleDbType.Integer, Value = id });
        }

        private DataTable GetCard(OleDbConnection cn, int id)
        {
            string sql = @"
SELECT * FROM (
 SELECT F.Datafatures AS DataLev, F.nrfature AS Dokument, L.idart, L.EmertimArtikulli AS Artikulli, L.SasiHyrje AS Hyrje, 0 AS Dalje
 FROM FaturehyrjeArtikujt AS L INNER JOIN Faturehyrje AS F ON L.idfaturehyrje=F.Idfaturehyrje
 WHERE L.idart=? AND F.Datafatures>=? AND F.Datafatures<? AND Nz(F.Anulluar,False)=False
 UNION ALL
 SELECT F.Datafatures AS DataLev, F.nrfature AS Dokument, L.idart, L.EmertimArtikulli AS Artikulli, 0 AS Hyrje, L.SasiShitje AS Dalje
 FROM FatureshitjeArtikujt AS L INNER JOIN FatureShitje AS F ON L.idfatureshitje=F.IdfatureShitje
 WHERE L.idart=? AND F.Datafatures>=? AND F.Datafatures<? AND Nz(F.Anulluar,False)=False
) AS M ORDER BY DataLev, Dokument";
            return Query(cn, sql,
                new OleDbParameter { OleDbType = OleDbType.Integer, Value = id },
                new OleDbParameter { OleDbType = OleDbType.Date, Value = fromDate.Value.Date },
                new OleDbParameter { OleDbType = OleDbType.Date, Value = toDate.Value.Date.AddDays(1) },
                new OleDbParameter { OleDbType = OleDbType.Integer, Value = id },
                new OleDbParameter { OleDbType = OleDbType.Date, Value = fromDate.Value.Date },
                new OleDbParameter { OleDbType = OleDbType.Date, Value = toDate.Value.Date.AddDays(1) });
        }

        private void ExportCsv(object sender, EventArgs e)
        {
            if (grid.DataSource == null) { MessageBox.Show("Shfaq fillimisht një raport."); return; }
            using (var dlg = new SaveFileDialog { Filter = "CSV (*.csv)|*.csv", FileName = "RaportMagazine.csv" })
            {
                if (dlg.ShowDialog() != DialogResult.OK) return;
                var dt = grid.DataSource as DataTable;
                using (var writer = new StreamWriter(dlg.FileName, false, System.Text.Encoding.UTF8))
                {
                    for (int c = 0; c < dt.Columns.Count; c++) { if (c > 0) writer.Write(";"); writer.Write(EscapeCsv(dt.Columns[c].ColumnName)); }
                    writer.WriteLine();
                    foreach (DataRow row in dt.Rows)
                    {
                        for (int c = 0; c < dt.Columns.Count; c++) { if (c > 0) writer.Write(";"); writer.Write(EscapeCsv(Convert.ToString(row[c]))); }
                        writer.WriteLine();
                    }
                }
                status.Text = "U eksportua: " + dlg.FileName;
            }
        }

        private string EscapeCsv(string value)
        {
            if (value == null) return "";
            return "\"" + value.Replace("\"", "\"\"") + "\"";
        }

        private void ShowDbError(Exception ex)
        {
            MessageBox.Show("Nuk u krye veprimi.\r\n\r\n" + ex.Message +
                "\r\n\r\nKontrollo që Microsoft Access Database Engine (ACE OLEDB) të jetë i instaluar dhe të ketë të njëjtën arkitekturë (32-bit) me programin.", "Gabim databaze", MessageBoxButtons.OK, MessageBoxIcon.Error);
            status.Text = "Gabim: " + ex.Message;
        }
    }
}

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Data.Sqlite;
using System.Globalization;
using System.Threading;

namespace PersonalFinanceTracker
{
    public partial class MainForm : Form
    {
        private BindingList<FinancialOperation> operations;
        private List<string> categories;
        private string connectionString = "Data Source=finance.db";
        private Color[] chartColors = {
            Color.Red, Color.Blue, Color.Green, Color.Orange, Color.Purple,
            Color.Teal, Color.Magenta, Color.Gold, Color.Brown, Color.Pink,
            Color.LightBlue, Color.LightGreen, Color.LightCoral, Color.LightSalmon,
            Color.Violet, Color.DeepSkyBlue, Color.Lime, Color.Coral
        };
        private Panel legendScrollPanel;

        public MainForm()
        {
            InitializeComponent();
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            Thread.CurrentThread.CurrentCulture = new CultureInfo("ru-RU");
            Thread.CurrentThread.CurrentUICulture = new CultureInfo("ru-RU");

            InitializeDatabase();
            LoadCategories();
            LoadOperations();

            comboBoxType.Items.Add("Доход");
            comboBoxType.Items.Add("Расход");
            comboBoxType.SelectedIndex = 0;

            // ИСПРАВЛЕНИЕ: Добавляем элементы только один раз
            comboBoxChartType.Items.Clear(); // Очищаем перед добавлением
            comboBoxChartType.Items.Add("Доходы по категориям");
            comboBoxChartType.Items.Add("Расходы по категориям");
            comboBoxChartType.Items.Add("Общая статистика");
            comboBoxChartType.SelectedIndex = 0;

            ConfigureDataGridView();
            UpdateBalance();
            UpdateCategoriesComboBox();
            InitializeMonthComboBox();
            UpdateMonthlyStatistics();
        }

        // ДОБАВЛЕННЫЕ МЕТОДЫ ДЛЯ ИСПРАВЛЕНИЯ ОШИБОК:

        private void InitializeDatabase()
        {
            using (var connection = new SqliteConnection(connectionString))
            {
                connection.Open();

                var command1 = connection.CreateCommand();
                command1.CommandText = @"
                    CREATE TABLE IF NOT EXISTS Operations (
                        Id INTEGER PRIMARY KEY AUTOINCREMENT,
                        Type TEXT NOT NULL,
                        Amount DECIMAL NOT NULL,
                        Category TEXT NOT NULL,
                        Date TEXT NOT NULL
                    )";
                command1.ExecuteNonQuery();

                var command2 = connection.CreateCommand();
                command2.CommandText = @"
                    CREATE TABLE IF NOT EXISTS Categories (
                        Id INTEGER PRIMARY KEY AUTOINCREMENT,
                        Name TEXT NOT NULL UNIQUE
                    )";
                command2.ExecuteNonQuery();

                InitializeDefaultCategories();
            }
        }

        private void InitializeDefaultCategories()
        {
            var defaultCategories = new List<string>
            {
                "Зарплата", "Премия", "Инвестиции", "Подарки",
                "Продукты", "Транспорт", "Жилье", "Развлечения",
                "Здоровье", "Одежда", "Образование", "Прочее"
            };

            using (var connection = new SqliteConnection(connectionString))
            {
                connection.Open();

                foreach (var category in defaultCategories)
                {
                    var command = connection.CreateCommand();
                    command.CommandText = "INSERT OR IGNORE INTO Categories (Name) VALUES (@name)";
                    command.Parameters.AddWithValue("@name", category);
                    command.ExecuteNonQuery();
                }
            }
        }

        private void LoadCategories()
        {
            categories = new List<string>();

            using (var connection = new SqliteConnection(connectionString))
            {
                connection.Open();
                var command = connection.CreateCommand();
                command.CommandText = "SELECT Name FROM Categories ORDER BY Name";

                using (var reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        categories.Add(reader.GetString(0));
                    }
                }
            }
        }

        private void LoadOperations()
        {
            operations = new BindingList<FinancialOperation>();

            using (var connection = new SqliteConnection(connectionString))
            {
                connection.Open();
                var command = connection.CreateCommand();
                command.CommandText = "SELECT * FROM Operations ORDER BY Date DESC";

                using (var reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        operations.Add(new FinancialOperation(
                            reader.GetInt32(0),
                            reader.GetString(1),
                            reader.GetDecimal(2),
                            reader.GetString(3),
                            DateTime.Parse(reader.GetString(4))
                        ));
                    }
                }
            }

            dataGridViewOperations.DataSource = operations;
        }

        private void ConfigureDataGridView()
        {
            dataGridViewOperations.AutoGenerateColumns = false;
            dataGridViewOperations.Columns.Clear();

            dataGridViewOperations.Columns.Add(new DataGridViewTextBoxColumn()
            {
                DataPropertyName = "Id",
                HeaderText = "ID",
                Visible = false
            });
            dataGridViewOperations.Columns.Add(new DataGridViewTextBoxColumn()
            {
                DataPropertyName = "Date",
                HeaderText = "Дата",
                Width = 120
            });
            dataGridViewOperations.Columns.Add(new DataGridViewTextBoxColumn()
            {
                DataPropertyName = "Type",
                HeaderText = "Тип",
                Width = 80
            });
            dataGridViewOperations.Columns.Add(new DataGridViewTextBoxColumn()
            {
                DataPropertyName = "Amount",
                HeaderText = "Сумма",
                Width = 100,
                DefaultCellStyle = new DataGridViewCellStyle { Format = "C2" }
            });
            dataGridViewOperations.Columns.Add(new DataGridViewTextBoxColumn()
            {
                DataPropertyName = "Category",
                HeaderText = "Категория",
                Width = 150
            });
        }

        private void UpdateBalance()
        {
            decimal balance = 0;

            foreach (var op in operations)
            {
                if (op.Type == "Доход")
                {
                    balance += op.Amount;
                }
                else if (op.Type == "Расход")
                {
                    balance -= op.Amount;
                }
            }

            textBoxBalance.Text = balance.ToString("C2");
            textBoxBalance.ForeColor = balance >= 0 ? Color.Green : Color.Red;
        }

        private void UpdateCategoriesComboBox()
        {
            comboBoxCategory.Items.Clear();

            foreach (var category in categories)
            {
                comboBoxCategory.Items.Add(category);
            }

            if (comboBoxCategory.Items.Count > 0)
            {
                comboBoxCategory.SelectedIndex = 0;
            }
        }

        private void InitializeMonthComboBox()
        {
            comboBoxMonth.Items.Clear();
            comboBoxMonth.Items.Add("Все месяцы");

            var allMonths = operations
                .Select(op => new DateTime(op.Date.Year, op.Date.Month, 1))
                .Distinct()
                .OrderByDescending(d => d)
                .ToList();

            if (!allMonths.Any())
            {
                allMonths.Add(new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1));
            }

            foreach (var monthDate in allMonths)
            {
                var monthName = monthDate.ToString("MMMM yyyy");
                comboBoxMonth.Items.Add(monthName);
            }

            if (comboBoxMonth.Items.Count > 0)
            {
                comboBoxMonth.SelectedIndex = 0;
            }
        }

        private void UpdateMonthlyStatistics()
        {
            if (comboBoxMonth.SelectedItem == null) return;

            var selectedMonth = comboBoxMonth.SelectedItem.ToString();

            if (selectedMonth == "Все месяцы")
            {
                ShowAllMonthsStatistics();
            }
            else
            {
                ShowSelectedMonthStatistics(selectedMonth);
            }
        }

        private void ShowAllMonthsStatistics()
        {
            decimal totalIncome = operations.Where(op => op.Type == "Доход").Sum(op => op.Amount);
            decimal totalExpenses = operations.Where(op => op.Type == "Расход").Sum(op => op.Amount);
            decimal totalBalance = totalIncome - totalExpenses;

            var currentMonth = DateTime.Now.Month;
            var currentYear = DateTime.Now.Year;
            decimal incomeThisMonth = operations
                .Where(op => op.Type == "Доход" && op.Date.Month == currentMonth && op.Date.Year == currentYear)
                .Sum(op => op.Amount);
            decimal expensesThisMonth = operations
                .Where(op => op.Type == "Расход" && op.Date.Month == currentMonth && op.Date.Year == currentYear)
                .Sum(op => op.Amount);

            var lastMonthDate = DateTime.Now.AddMonths(-1);
            decimal incomeLastMonth = operations
                .Where(op => op.Type == "Доход" && op.Date.Month == lastMonthDate.Month && op.Date.Year == lastMonthDate.Year)
                .Sum(op => op.Amount);
            decimal expensesLastMonth = operations
                .Where(op => op.Type == "Расход" && op.Date.Month == lastMonthDate.Month && op.Date.Year == lastMonthDate.Year)
                .Sum(op => op.Amount);

            lblIncomeThisMonth.Text = totalIncome.ToString("C2");
            lblExpensesThisMonth.Text = totalExpenses.ToString("C2");
            lblBalanceThisMonth.Text = totalBalance.ToString("C2");
            lblBalanceThisMonth.ForeColor = totalBalance >= 0 ? Color.Green : Color.Red;

            lblIncomeLastMonth.Text = incomeThisMonth.ToString("C2");
            lblExpensesLastMonth.Text = expensesThisMonth.ToString("C2");

            label5.Text = "Доход:";
            label6.Text = "Расход:";
            label7.Text = "Расход:";
            /*label8.Text = "Текущий месяц:";*/
            label9.Text = "Доход:";

            label7.Visible = true;
            label9.Visible = true;
        }

        private void ShowSelectedMonthStatistics(string selectedMonth)
        {
            try
            {
                var monthDate = DateTime.ParseExact(selectedMonth, "MMMM yyyy", CultureInfo.GetCultureInfo("ru-RU"));

                decimal income = operations
                    .Where(op => op.Type == "Доход" && op.Date.Month == monthDate.Month && op.Date.Year == monthDate.Year)
                    .Sum(op => op.Amount);
                decimal expenses = operations
                    .Where(op => op.Type == "Расход" && op.Date.Month == monthDate.Month && op.Date.Year == monthDate.Year)
                    .Sum(op => op.Amount);
                decimal balance = income - expenses;

                var previousMonthDate = monthDate.AddMonths(-1);
                decimal incomePrevious = operations
                    .Where(op => op.Type == "Доход" && op.Date.Month == previousMonthDate.Month && op.Date.Year == previousMonthDate.Year)
                    .Sum(op => op.Amount);
                decimal expensesPrevious = operations
                    .Where(op => op.Type == "Расход" && op.Date.Month == previousMonthDate.Month && op.Date.Year == previousMonthDate.Year)
                    .Sum(op => op.Amount);

                lblIncomeThisMonth.Text = income.ToString("C2");
                lblExpensesThisMonth.Text = expenses.ToString("C2");
                lblBalanceThisMonth.Text = balance.ToString("C2");
                lblBalanceThisMonth.ForeColor = balance >= 0 ? Color.Green : Color.Red;

                lblIncomeLastMonth.Text = incomePrevious.ToString("C2");
                lblExpensesLastMonth.Text = expensesPrevious.ToString("C2");

                label5.Text = "Доход:";
                label6.Text = "Расход:";
                label7.Text = "Пред. месяц:";
                /*label8.Text = "Пред. месяц:";*/
                label9.Text = "Пред. месяц:";

                label7.Visible = true;
                label9.Visible = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при обработке месяца: {ex.Message}", "Ошибка",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void AddOperationToDatabase(FinancialOperation operation)
        {
            using (var connection = new SqliteConnection(connectionString))
            {
                connection.Open();
                var command = connection.CreateCommand();
                command.CommandText = @"
                    INSERT INTO Operations (Type, Amount, Category, Date)
                    VALUES (@type, @amount, @category, @date)";

                command.Parameters.AddWithValue("@type", operation.Type);
                command.Parameters.AddWithValue("@amount", operation.Amount);
                command.Parameters.AddWithValue("@category", operation.Category);
                command.Parameters.AddWithValue("@date", operation.Date.ToString("yyyy-MM-dd"));

                command.ExecuteNonQuery();
            }
        }

        private void DeleteOperationFromDatabase(int id)
        {
            using (var connection = new SqliteConnection(connectionString))
            {
                connection.Open();
                var command = connection.CreateCommand();
                command.CommandText = "DELETE FROM Operations WHERE Id = @id";
                command.Parameters.AddWithValue("@id", id);
                command.ExecuteNonQuery();
            }
        }

        // ОБРАБОТЧИКИ СОБЫТИЙ:

        private void comboBoxType_SelectedIndexChanged(object sender, EventArgs e)
        {
            // Можно добавить логику фильтрации категорий по типу
        }

        private void buttonAdd_Click(object sender, EventArgs e)
        {
            if (decimal.TryParse(textBoxAmount.Text, out decimal amount) && amount > 0)
            {
                if (comboBoxCategory.SelectedItem == null)
                {
                    MessageBox.Show("Выберите категорию.", "Ошибка",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                var newOperation = new FinancialOperation(
                    0,
                    comboBoxType.SelectedItem.ToString(),
                    amount,
                    comboBoxCategory.SelectedItem.ToString(),
                    dateTimePicker.Value
                );

                AddOperationToDatabase(newOperation);
                LoadOperations();

                textBoxAmount.Clear();
                UpdateBalance();
                UpdateMonthlyStatistics();
                panelChart.Invalidate();

                InitializeMonthComboBox();
            }
            else
            {
                MessageBox.Show("Пожалуйста, введите корректную сумму.", "Ошибка ввода",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void buttonUpdateBalance_Click(object sender, EventArgs e)
        {
            LoadOperations();
            UpdateBalance();
            UpdateMonthlyStatistics();
            panelChart.Invalidate();
        }

        private void buttonDelete_Click(object sender, EventArgs e)
        {
            if (dataGridViewOperations.CurrentRow != null)
            {
                var selectedOperation = (FinancialOperation)dataGridViewOperations.CurrentRow.DataBoundItem;

                var result = MessageBox.Show($"Удалить операцию '{selectedOperation.Category}' на сумму {selectedOperation.Amount:C2}?",
                    "Подтверждение удаления", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    DeleteOperationFromDatabase(selectedOperation.Id);
                    LoadOperations();
                    UpdateBalance();
                    UpdateMonthlyStatistics();
                    panelChart.Invalidate();
                    InitializeMonthComboBox();
                }
            }
            else
            {
                MessageBox.Show("Выберите операцию для удаления.", "Информация",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void buttonManageCategories_Click(object sender, EventArgs e)
        {
            using (var form = new CategoriesForm(categories, connectionString))
            {
                if (form.ShowDialog() == DialogResult.OK)
                {
                    LoadCategories();
                    UpdateCategoriesComboBox();
                    LoadOperations();
                    panelChart.Invalidate();
                }
            }
        }

        private void comboBoxMonth_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdateMonthlyStatistics();
            panelChart.Invalidate();
        }

        private void comboBoxChartType_SelectedIndexChanged(object sender, EventArgs e)
        {
            panelChart.Invalidate();
        }

        // МЕТОДЫ ДЛЯ ДИАГРАММЫ (с исправлением проблемы с dynamic):

        private void panelChart_Paint(object sender, PaintEventArgs e)
        {
            // Очищаем старые контролы легенды
            var oldPanels = panelChart.Controls.OfType<Panel>().Where(p => p.Name == "legendPanel").ToList();
            foreach (var panel in oldPanels)
            {
                panelChart.Controls.Remove(panel);
                panel.Dispose();
            }

            DrawChart(e.Graphics);
        }

        private void DrawChart(Graphics g)
        {
            g.Clear(Color.White);
            g.SmoothingMode = SmoothingMode.AntiAlias;

            var chartType = comboBoxChartType.SelectedItem?.ToString();
            if (string.IsNullOrEmpty(chartType)) return;

            var selectedMonth = comboBoxMonth.SelectedItem?.ToString();
            var operationsToUse = GetOperationsForSelectedPeriod(selectedMonth);

            if (!operationsToUse.Any())
            {
                DrawNoDataMessage(g);
                return;
            }

            switch (chartType)
            {
                case "Доходы по категориям":
                    DrawIncomePieChart(g, operationsToUse);
                    break;
                case "Расходы по категориям":
                    DrawExpensesPieChart(g, operationsToUse);
                    break;
                case "Общая статистика":
                    DrawGeneralStatistics(g, operationsToUse);
                    break;
            }
        }

        private List<FinancialOperation> GetOperationsForSelectedPeriod(string selectedMonth)
        {
            if (selectedMonth == "Все месяцы")
                return operations.ToList();

            try
            {
                var monthDate = DateTime.ParseExact(selectedMonth, "MMMM yyyy", CultureInfo.GetCultureInfo("ru-RU"));
                return operations.Where(op => op.Date.Month == monthDate.Month && op.Date.Year == monthDate.Year).ToList();
            }
            catch
            {
                return operations.ToList();
            }
        }

        private void DrawNoDataMessage(Graphics g)
        {
            var font = new Font("Arial", 12, FontStyle.Bold);
            var brush = new SolidBrush(Color.Gray);
            var text = "Нет данных для отображения";
            var textSize = g.MeasureString(text, font);
            var x = (panelChart.Width - textSize.Width) / 2;
            var y = (panelChart.Height - textSize.Height) / 2;

            g.DrawString(text, font, brush, x, y);
        }

        private void DrawIncomePieChart(Graphics g, List<FinancialOperation> operations)
        {
            var incomeOperations = operations.Where(op => op.Type == "Доход").ToList();
            if (!incomeOperations.Any())
            {
                DrawNoDataMessage(g);
                return;
            }

            // Используем конкретный тип вместо dynamic
            var incomeByCategory = incomeOperations
                .GroupBy(op => op.Category)
                .Select(g => new CategoryAmount { Category = g.Key, Amount = g.Sum(op => op.Amount) })
                .OrderByDescending(x => x.Amount)
                .ToList();

            DrawPieChart(g, incomeByCategory, "Доходы по категориям");
        }

        private void DrawExpensesPieChart(Graphics g, List<FinancialOperation> operations)
        {
            var expenseOperations = operations.Where(op => op.Type == "Расход").ToList();
            if (!expenseOperations.Any())
            {
                DrawNoDataMessage(g);
                return;
            }

            // Используем конкретный тип вместо dynamic
            var expensesByCategory = expenseOperations
                .GroupBy(op => op.Category)
                .Select(g => new CategoryAmount { Category = g.Key, Amount = g.Sum(op => op.Amount) })
                .OrderByDescending(x => x.Amount)
                .ToList();

            DrawPieChart(g, expensesByCategory, "Расходы по категориям");
        }

        // Вспомогательный класс для замены dynamic
        private class CategoryAmount
        {
            public string Category { get; set; }
            public decimal Amount { get; set; }
        }

        private void DrawPieChart(Graphics g, List<CategoryAmount> data, string title)
        {
            var total = data.Sum(x => x.Amount);
            var centerX = panelChart.Width / 2;
            var centerY = panelChart.Height / 2 + 100; // Опускаем диаграмму еще ниже

            // Рассчитываем радиус в зависимости от количества данных
            var radius = Math.Min(centerX, panelChart.Height - centerY) - 20;
            radius = Math.Min(radius, 150); // Ограничиваем максимальный радиус

            // Рисуем заголовок
            var titleFont = new Font("Arial", 12, FontStyle.Bold);
            var titleSize = g.MeasureString(title, titleFont);
            g.DrawString(title, titleFont, Brushes.Black, (panelChart.Width - titleSize.Width) / 2, 10);

            // Рисуем общую сумму ЧЕРНЫМ цветом
            var totalFont = new Font("Arial", 10, FontStyle.Bold);
            var totalText = $"Всего: {total:C2}";
            var totalTextSize = g.MeasureString(totalText, totalFont);
            var textX = centerX - totalTextSize.Width / 2;
            var textY = 35; // Позиция под заголовком
            g.DrawString(totalText, totalFont, Brushes.Black, textX, textY);

            // Создаем прокручиваемую легенду
            CreateScrollableLegend(data, total);

            // Рисуем секторы диаграммы (только если есть данные)
            if (data.Count > 0)
            {
                float startAngle = 0;

                for (int i = 0; i < data.Count; i++)
                {
                    var item = data[i];
                    var sweepAngle = (float)((double)item.Amount / (double)total * 360);

                    using (var brush = new SolidBrush(chartColors[i % chartColors.Length]))
                    {
                        g.FillPie(brush, centerX - radius, centerY - radius, radius * 2, radius * 2, startAngle, sweepAngle);
                        g.DrawPie(Pens.Black, centerX - radius, centerY - radius, radius * 2, radius * 2, startAngle, sweepAngle);
                    }

                    startAngle += sweepAngle;
                }
            }
        }

        private void DrawLegendAboveChart(Graphics g, List<CategoryAmount> data, decimal total)
        {
            // Определяем размеры и отступы
            var startX = 10;
            var startY = 50; // Начинаем ниже, чтобы освободить место для заголовка и текста "Всего"
            var lineHeight = 18;
            var itemsPerColumn = 6; // Увеличиваем количество элементов в колонке для более компактного отображения
            var columnWidth = 220; // Уменьшаем ширину колонки

            // Вычисляем необходимое количество колонок
            var columnCount = Math.Min(3, (int)Math.Ceiling((double)data.Count / itemsPerColumn)); // Максимум 3 колонки

            // Создаем контейнер для легенды с прокруткой
            var legendPanel = new Panel();
            legendPanel.Width = panelChart.Width - 20;
            legendPanel.Height = Math.Min(120, (int)Math.Ceiling((double)data.Count / columnCount) * lineHeight + 10);
            legendPanel.Location = new Point(startX, startY);
            legendPanel.BorderStyle = BorderStyle.FixedSingle;
            legendPanel.AutoScroll = true;

            // Временный контроль для отрисовки
            var legendHeight = 0;

            for (int i = 0; i < data.Count; i++)
            {
                var item = data[i];
                var column = i / itemsPerColumn;
                var row = i % itemsPerColumn;

                var x = column * columnWidth;
                var y = row * lineHeight;

                // Обновляем высоту контейнера если нужно
                if (y + lineHeight > legendHeight)
                {
                    legendHeight = y + lineHeight;
                }

                // Цветной квадрат
                var colorRect = new Rectangle(x, y, 12, 12);
                using (var brush = new SolidBrush(chartColors[i % chartColors.Length]))
                {
                    g.FillRectangle(brush, colorRect);
                }
                g.DrawRectangle(Pens.Black, colorRect);

                // Текст легенды
                var percentage = (item.Amount / total * 100).ToString("F1");
                var legendText = $"{item.Category}: {item.Amount:C2} ({percentage}%)";
                var font = new Font("Arial", 8);
                g.DrawString(legendText, font, Brushes.Black, x + 15, y - 2);
            }

            // Если элементов много, показываем полосу прокрутки
            if (legendHeight > 120)
            {
                // Рисуем рамку для области легенды
                var legendRect = new Rectangle(startX, startY, panelChart.Width - 20, 120);
                g.DrawRectangle(Pens.Gray, legendRect);

                // Рисуем текст "Прокрутите для просмотра всех категорий"
                var scrollHint = "Прокрутите для просмотра всех категорий";
                var hintFont = new Font("Arial", 8, FontStyle.Italic);
                var hintSize = g.MeasureString(scrollHint, hintFont);
                g.DrawString(scrollHint, hintFont, Brushes.Gray,
                             panelChart.Width - hintSize.Width - 10, startY + 125);
            }
        }

        // Создаем отдельный контрол с прокруткой для легенды
        private void CreateScrollableLegend(List<CategoryAmount> data, decimal total)
        {
            // Создаем панель для легенды
            var legendPanel = new Panel();
            legendPanel.Name = "legendPanel";
            legendPanel.Location = new Point(10, 50);
            legendPanel.Size = new Size(panelChart.Width - 20, 120);
            legendPanel.BackColor = Color.White;
            legendPanel.BorderStyle = BorderStyle.FixedSingle;
            legendPanel.AutoScroll = true;

            // Добавляем элементы легенды
            var yPos = 5;
            for (int i = 0; i < data.Count; i++)
            {
                var item = data[i];
                var percentage = (item.Amount / total * 100).ToString("F1");

                // Создаем панель для элемента легенды
                var itemPanel = new Panel();
                itemPanel.Size = new Size(legendPanel.Width - 25, 18);
                itemPanel.Location = new Point(5, yPos);

                // Цветной квадрат
                var colorBox = new PictureBox();
                colorBox.Size = new Size(12, 12);
                colorBox.Location = new Point(0, 3);
                colorBox.BackColor = chartColors[i % chartColors.Length];
                colorBox.BorderStyle = BorderStyle.FixedSingle;

                // Текст
                var label = new Label();
                label.Text = $"{item.Category}: {item.Amount:C2} ({percentage}%)";
                label.Location = new Point(18, 0);
                label.Size = new Size(itemPanel.Width - 20, 18);
                label.Font = new Font("Arial", 8);

                itemPanel.Controls.Add(colorBox);
                itemPanel.Controls.Add(label);
                legendPanel.Controls.Add(itemPanel);

                yPos += 20;
            }

            // Устанавливаем высоту контента
            legendPanel.AutoScrollMinSize = new Size(legendPanel.Width - 25, yPos + 5);

            // Добавляем панель на panelChart
            panelChart.Controls.Add(legendPanel);
            legendPanel.BringToFront();
        }

        private void DrawGeneralStatistics(Graphics g, List<FinancialOperation> operations)
        {
            var totalIncome = operations.Where(op => op.Type == "Доход").Sum(op => op.Amount);
            var totalExpenses = operations.Where(op => op.Type == "Расход").Sum(op => op.Amount);
            var balance = totalIncome - totalExpenses;

            var centerX = panelChart.Width / 2;
            var centerY = panelChart.Height / 2;

            // Заголовок
            var titleFont = new Font("Arial", 12, FontStyle.Bold);
            var title = "Общая статистика";
            var titleSize = g.MeasureString(title, titleFont);
            g.DrawString(title, titleFont, Brushes.Black, (panelChart.Width - titleSize.Width) / 2, 20);

            // Данные
            var dataFont = new Font("Arial", 10);
            var y = 60;

            g.DrawString($"Доходы: {totalIncome:C2}", dataFont, Brushes.Green, centerX - 100, y);
            y += 25;
            g.DrawString($"Расходы: {totalExpenses:C2}", dataFont, Brushes.Red, centerX - 100, y);
            y += 25;
            g.DrawString($"Баланс: {balance:C2}", dataFont, balance >= 0 ? Brushes.Green : Brushes.Red, centerX - 100, y);

            // Простая круговая диаграмма доходов/расходов
            if (totalIncome + totalExpenses > 0)
            {
                var pieRadius = 80;
                var pieY = centerY + 50;

                if (totalIncome > 0)
                {
                    var incomeAngle = (float)(totalIncome / (totalIncome + totalExpenses) * 360);
                    g.FillPie(Brushes.Green, centerX - pieRadius, pieY - pieRadius, pieRadius * 2, pieRadius * 2, 0, incomeAngle);
                }

                if (totalExpenses > 0)
                {
                    var expensesAngle = (float)(totalExpenses / (totalIncome + totalExpenses) * 360);
                    g.FillPie(Brushes.Red, centerX - pieRadius, pieY - pieRadius, pieRadius * 2, pieRadius * 2,
                        (float)(totalIncome / (totalIncome + totalExpenses) * 360), expensesAngle);
                }

                g.DrawPie(Pens.Black, centerX - pieRadius, pieY - pieRadius, pieRadius * 2, pieRadius * 2, 0, 360);

                // Легенда для общей диаграммы
                var legendY = pieY + pieRadius + 20;
                g.FillRectangle(Brushes.Green, centerX - 40, legendY, 15, 15);
                g.DrawString("Доходы", dataFont, Brushes.Black, centerX - 20, legendY);

                g.FillRectangle(Brushes.Red, centerX + 30, legendY, 15, 15);
                g.DrawString("Расходы", dataFont, Brushes.Black, centerX + 50, legendY);
            }
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void lblExpensesLastMonth_Click(object sender, EventArgs e)
        {

        }

        private void dateTimePicker_ValueChanged(object sender, EventArgs e)
        {

        }
        private void label1_Click(object sender, EventArgs e)
        {

        }
        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void label10_Click(object sender, EventArgs e)
        {

        }
    }


    public class FinancialOperation
    {
        public int Id { get; set; }
        public string Type { get; set; }
        public decimal Amount { get; set; }
        public string Category { get; set; }
        public DateTime Date { get; set; }

        public FinancialOperation(int id, string type, decimal amount, string category, DateTime date)
        {
            Id = id;
            Type = type;
            Amount = amount;
            Category = category;
            Date = date;
        }


        public FinancialOperation() { }
    }
}
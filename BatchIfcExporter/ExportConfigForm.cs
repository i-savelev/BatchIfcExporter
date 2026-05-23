using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace BatchExportIfc
{
    /// <summary>
    /// Окно настройки экспорта: выбор папки, файла конфигурации, запуск.
    /// Без рамок у полей, с полноценным диалогом выбора папки.
    /// </summary>
    public class ExportConfigForm : Form
    {
        // 🔑 Публичные элементы для подписки из команды
        public Button BtnSelectIfcFolder { get; private set; }
        public TextBox TxtIfcFolder { get; private set; }

        public Button BtnSelectConfigFile { get; private set; }
        public TextBox TxtConfigFile { get; private set; }

        public Button BtnRunExport { get; private set; }
        public Button BtnSaveTemplate { get; private set; }

        public Button BtnClose { get; private set; }
        public Label LblStatus { get; private set; }

        public ExportConfigForm()
        {
            SetupForm();
            SetupControls();
        }

        private void SetupForm()
        {
            Text = "Настройка экспорта IFC";
            Size = new Size(650, 300);
            MinimumSize = new Size(550, 280);
            StartPosition = FormStartPosition.CenterParent;
            Font = new Font("Segoe UI", 9F);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
        }

        private void SetupControls()
        {
            var mainPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 3,
                RowCount = 5,
                Padding = new Padding(15, 10, 15, 10),
                AutoSize = true
            };

            // Настройка колонок: [Кнопка 180] [TextBox *] [Пусто 40]
            mainPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 180));
            mainPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            mainPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 40));

            // === Строка 1: Папка с моделями ===
            var lblIfc = new Label
            {
                Text = "Папка с моделями Revit:",
                Dock = DockStyle.Fill,
                Margin = new Padding(0, 0, 0, 3),
                AutoSize = true
            };
            mainPanel.Controls.Add(lblIfc, 0, 0);
            mainPanel.SetColumnSpan(lblIfc, 3);

            BtnSelectIfcFolder = new Button
            {
                Text = "Выбрать папку...",
                Width = 160,
                FlatStyle = FlatStyle.Flat,
                TextAlign = ContentAlignment.MiddleLeft
            };

            // 🔥 TextBox без рамки, с прозрачным фоном — выглядит как просто текст
            TxtIfcFolder = new TextBox
            {
                ReadOnly = true,
                BorderStyle = BorderStyle.None,           // 👈 Убрана рамка
                BackColor = SystemColors.Window,          // 👈 Белый фон как у формы
                ForeColor = SystemColors.ControlText,
                Dock = DockStyle.Fill,
                TabStop = false,
                Padding = new Padding(2, 4, 2, 4),        // 👈 Отступы для текста
                Margin = new Padding(5, 0, 0, 0)          // 👈 Отступ от кнопки
            };
            TxtIfcFolder.MouseEnter += (s, e) => { TxtIfcFolder.Cursor = Cursors.Default; };

            mainPanel.Controls.Add(BtnSelectIfcFolder, 0, 1);
            mainPanel.Controls.Add(TxtIfcFolder, 1, 1);
            mainPanel.SetColumnSpan(TxtIfcFolder, 2);

            // === Строка 2: Файл конфигурации ===
            var lblConfig = new Label
            {
                Text = "Файл конфигурации (Excel):",
                Dock = DockStyle.Fill,
                Margin = new Padding(0, 8, 0, 3),
                AutoSize = true
            };
            mainPanel.Controls.Add(lblConfig, 0, 2);
            mainPanel.SetColumnSpan(lblConfig, 3);

            BtnSelectConfigFile = new Button
            {
                Text = "Выбрать файл...",
                Width = 160,
                FlatStyle = FlatStyle.Flat,
                TextAlign = ContentAlignment.MiddleLeft
            };

            TxtConfigFile = new TextBox
            {
                ReadOnly = true,
                BorderStyle = BorderStyle.None,
                BackColor = SystemColors.Window,
                ForeColor = SystemColors.ControlText,
                Dock = DockStyle.Fill,
                TabStop = false,
                Padding = new Padding(2, 4, 2, 4),
                Margin = new Padding(5, 0, 0, 0)
            };
            TxtConfigFile.MouseEnter += (s, e) => { TxtConfigFile.Cursor = Cursors.Default; };

            mainPanel.Controls.Add(BtnSelectConfigFile, 0, 3);
            mainPanel.Controls.Add(TxtConfigFile, 1, 3);
            mainPanel.SetColumnSpan(TxtConfigFile, 2);

            // === Строка 3: Кнопки действий ===
            var actionPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                Padding = new Padding(0, 12, 0, 0),
                AutoSize = true
            };
            mainPanel.SetColumnSpan(actionPanel, 3);
            mainPanel.Controls.Add(actionPanel, 0, 4);

            BtnSaveTemplate = new Button
            {
                Text = "Сохранить шаблон",
                Width = 160,
                FlatStyle = FlatStyle.Flat,
                Margin = new Padding(0, 0, 10, 0)
            };

            BtnRunExport = new Button
            {
                Text = "Запустить экспорт",
                Width = 140,
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(220, 240, 220)
            };

            BtnClose = new Button
            {
                Text = "Закрыть",
                Width = 90,
                FlatStyle = FlatStyle.Flat,
                DialogResult = DialogResult.Cancel
            };

            actionPanel.Controls.Add(BtnSaveTemplate);
            actionPanel.Controls.Add(BtnRunExport);
            actionPanel.Controls.Add(BtnClose);

            // === Строка 4: Статус ===
            LblStatus = new Label
            {
                Dock = DockStyle.Bottom,
                Height = 24,
                BackColor = SystemColors.Control,
                ForeColor = SystemColors.ControlText,
                Padding = new Padding(5, 4, 0, 0),
                Text = "Готов",
                TextAlign = ContentAlignment.MiddleLeft
            };

            // === Сборка формы ===
            Controls.Add(mainPanel);
            Controls.Add(LblStatus);
        }

        /// <summary>
        /// Устанавливает текст в статус-бар.
        /// </summary>
        public void SetStatus(string message) => LblStatus.Text = message;

        /// <summary>
        /// Возвращает выбранный путь к папке с моделями.
        /// </summary>
        public string IfcFolderPath => TxtIfcFolder.Text;

        /// <summary>
        /// Возвращает выбранный путь к файлу конфигурации.
        /// </summary>
        public string ConfigFilePath => TxtConfigFile.Text;
    }
}
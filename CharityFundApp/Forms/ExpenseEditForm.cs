using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using CharityFundApp.Models;

namespace CharityFundApp.Forms
{
    public class ExpenseEditForm : Form
    {
        private readonly ComboBox _cmbProject;
        private readonly ComboBox _cmbRecipient;
        private readonly CheckBox _chkHasRecipient;
        private readonly ComboBox _cmbCategory;
        private readonly NumericUpDown _numAmount;
        private readonly DateTimePicker _dtpDate;
        private readonly TextBox _txtDocNumber;
        private readonly TextBox _txtDescription;
        private readonly Button _btnSave;
        private readonly Button _btnCancel;

        public Expense Expense { get; private set; }

        public ExpenseEditForm(Expense? expense, List<Project> projects, List<Recipient> recipients)
        {
            Expense = expense ?? new Expense();
            Text = expense == null ? "Регистрация расхода средств" : "Редактирование расхода";
            Size = new Size(540, 530);
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            Font = new Font("Segoe UI", 9.5f);

            int y = 15;

            var lblProj = new Label { Text = "Проект списания средств*:", Location = new Point(20, y), AutoSize = true, Font = new Font("Segoe UI", 9.5f, FontStyle.Bold) };
            y += 25;
            _cmbProject = new ComboBox { Location = new Point(20, y), Width = 480, DropDownStyle = ComboBoxStyle.DropDownList };
            _cmbProject.DataSource = projects;
            _cmbProject.DisplayMember = "Name";
            _cmbProject.ValueMember = "Id";
            if (Expense.ProjectId > 0) _cmbProject.SelectedValue = Expense.ProjectId;
            y += 35;

            _chkHasRecipient = new CheckBox { Text = "Связать с конкретным получателем помощи", Location = new Point(20, y), AutoSize = true, Checked = Expense.RecipientId.HasValue };
            y += 25;
            _cmbRecipient = new ComboBox { Location = new Point(20, y), Width = 480, DropDownStyle = ComboBoxStyle.DropDownList, Enabled = Expense.RecipientId.HasValue };
            _cmbRecipient.DataSource = recipients;
            _cmbRecipient.DisplayMember = "FullName";
            _cmbRecipient.ValueMember = "Id";
            if (Expense.RecipientId.HasValue) _cmbRecipient.SelectedValue = Expense.RecipientId.Value;
            _chkHasRecipient.CheckedChanged += (s, e) => _cmbRecipient.Enabled = _chkHasRecipient.Checked;
            y += 35;

            var lblCat = new Label { Text = "Статья расхода*:", Location = new Point(20, y), AutoSize = true };
            var lblAmount = new Label { Text = "Сумма (руб.)*:", Location = new Point(270, y), AutoSize = true };
            y += 25;

            _cmbCategory = new ComboBox { Location = new Point(20, y), Width = 230, DropDownStyle = ComboBoxStyle.DropDownList };
            _cmbCategory.Items.AddRange(new object[] {
                "Адресная помощь",
                "Закупка медикаментов",
                "Оборудование",
                "Транспорт",
                "Административные расходы"
            });
            _cmbCategory.SelectedItem = Expense.ExpenseCategory;
            if (_cmbCategory.SelectedIndex < 0) _cmbCategory.SelectedIndex = 0;

            _numAmount = new NumericUpDown { Location = new Point(270, y), Width = 230, DecimalPlaces = 2, Maximum = 100000000m, Value = Expense.Amount > 0 ? Expense.Amount : 5000m };
            y += 35;

            var lblDoc = new Label { Text = "Номер первичного документа (чек / акт / накладная)*:", Location = new Point(20, y), AutoSize = true };
            y += 25;
            _txtDocNumber = new TextBox { Location = new Point(20, y), Width = 480, Text = Expense.DocumentNumber };
            y += 35;

            var lblDate = new Label { Text = "Дата расхода:", Location = new Point(20, y), AutoSize = true };
            _dtpDate = new DateTimePicker { Location = new Point(130, y - 3), Width = 150, Format = DateTimePickerFormat.Short, Value = Expense.ExpenseDate };
            y += 35;

            var lblDesc = new Label { Text = "Назначение / Детализация расхода:", Location = new Point(20, y), AutoSize = true };
            y += 25;
            _txtDescription = new TextBox { Location = new Point(20, y), Width = 480, Height = 55, Multiline = true, Text = Expense.Description ?? "" };
            y += 70;

            _btnSave = new Button { Text = "Сохранить", Location = new Point(290, y), Width = 100, Height = 32, DialogResult = DialogResult.OK, BackColor = Color.FromArgb(37, 99, 235), ForeColor = Color.White, FlatStyle = FlatStyle.Flat };
            _btnCancel = new Button { Text = "Отмена", Location = new Point(400, y), Width = 100, Height = 32, DialogResult = DialogResult.Cancel };

            _btnSave.Click += BtnSave_Click;

            Controls.AddRange(new Control[] {
                lblProj, _cmbProject, _chkHasRecipient, _cmbRecipient,
                lblCat, _cmbCategory, lblAmount, _numAmount,
                lblDoc, _txtDocNumber, lblDate, _dtpDate,
                lblDesc, _txtDescription, _btnSave, _btnCancel
            });

            AcceptButton = _btnSave;
            CancelButton = _btnCancel;
        }

        private void BtnSave_Click(object? sender, EventArgs e)
        {
            if (_cmbProject.SelectedValue == null)
            {
                MessageBox.Show("Выберите проект!", "Валидация", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                DialogResult = DialogResult.None;
                return;
            }

            if (_numAmount.Value <= 0)
            {
                MessageBox.Show("Сумма расхода должна быть больше нуля!", "Валидация", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                DialogResult = DialogResult.None;
                return;
            }

            if (string.IsNullOrWhiteSpace(_txtDocNumber.Text))
            {
                MessageBox.Show("Введите номер подтверждающего документа!", "Валидация", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                DialogResult = DialogResult.None;
                _txtDocNumber.Focus();
                return;
            }

            Expense.ProjectId = Convert.ToInt32(_cmbProject.SelectedValue);
            Expense.RecipientId = _chkHasRecipient.Checked && _cmbRecipient.SelectedValue != null 
                ? Convert.ToInt32(_cmbRecipient.SelectedValue) 
                : null;
            Expense.ExpenseCategory = _cmbCategory.SelectedItem?.ToString() ?? "Адресная помощь";
            Expense.Amount = _numAmount.Value;
            Expense.ExpenseDate = _dtpDate.Value.Date;
            Expense.DocumentNumber = _txtDocNumber.Text.Trim();
            Expense.Description = string.IsNullOrWhiteSpace(_txtDescription.Text) ? null : _txtDescription.Text.Trim();
        }
    }
}

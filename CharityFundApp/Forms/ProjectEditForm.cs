using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using CharityFundApp.Models;

namespace CharityFundApp.Forms
{
    public class ProjectEditForm : Form
    {
        private readonly TextBox _txtName;
        private readonly ComboBox _cmbCategory;
        private readonly NumericUpDown _numTarget;
        private readonly NumericUpDown _numCurrent;
        private readonly DateTimePicker _dtpStart;
        private readonly CheckBox _chkHasEndDate;
        private readonly DateTimePicker _dtpEnd;
        private readonly ComboBox _cmbStatus;
        private readonly TextBox _txtDescription;
        private readonly Button _btnSave;
        private readonly Button _btnCancel;

        public Project Project { get; private set; }

        public ProjectEditForm(Project? project, List<Category> categories)
        {
            Project = project ?? new Project();
            Text = project == null ? "Добавление благотворительного проекта" : "Редактирование проекта";
            Size = new Size(540, 520);
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            Font = new Font("Segoe UI", 9.5f);

            int y = 15;

            var lblName = new Label { Text = "Название проекта*:", Location = new Point(20, y), AutoSize = true, Font = new Font("Segoe UI", 9.5f, FontStyle.Bold) };
            y += 25;
            _txtName = new TextBox { Location = new Point(20, y), Width = 480, Text = Project.Name };
            y += 35;

            var lblCat = new Label { Text = "Категория благотворительности*:", Location = new Point(20, y), AutoSize = true };
            y += 25;
            _cmbCategory = new ComboBox { Location = new Point(20, y), Width = 480, DropDownStyle = ComboBoxStyle.DropDownList };
            _cmbCategory.DataSource = categories;
            _cmbCategory.DisplayMember = "Name";
            _cmbCategory.ValueMember = "Id";
            if (Project.CategoryId > 0)
            {
                _cmbCategory.SelectedValue = Project.CategoryId;
            }
            y += 35;

            var lblTarget = new Label { Text = "Целевая сумма (руб.)*:", Location = new Point(20, y), AutoSize = true };
            var lblCurrent = new Label { Text = "Собранная сумма (руб.):", Location = new Point(270, y), AutoSize = true };
            y += 25;

            _numTarget = new NumericUpDown { Location = new Point(20, y), Width = 230, DecimalPlaces = 2, Maximum = 1000000000m, Value = Project.TargetAmount > 0 ? Project.TargetAmount : 100000m };
            _numCurrent = new NumericUpDown { Location = new Point(270, y), Width = 230, DecimalPlaces = 2, Maximum = 1000000000m, Value = Project.CurrentAmount };
            y += 35;

            var lblStart = new Label { Text = "Дата начала*:", Location = new Point(20, y), AutoSize = true };
            var lblEnd = new Label { Text = "Дата окончания:", Location = new Point(270, y), AutoSize = true };
            y += 25;

            _dtpStart = new DateTimePicker { Location = new Point(20, y), Width = 230, Format = DateTimePickerFormat.Short, Value = Project.StartDate };
            _chkHasEndDate = new CheckBox { Text = "Указать дату", Location = new Point(270, y), AutoSize = true, Checked = Project.EndDate.HasValue };
            _dtpEnd = new DateTimePicker { Location = new Point(370, y), Width = 130, Format = DateTimePickerFormat.Short, Enabled = Project.EndDate.HasValue, Value = Project.EndDate ?? DateTime.Today.AddMonths(6) };
            _chkHasEndDate.CheckedChanged += (s, e) => _dtpEnd.Enabled = _chkHasEndDate.Checked;
            y += 35;

            var lblStatus = new Label { Text = "Статус проекта*:", Location = new Point(20, y), AutoSize = true };
            y += 25;
            _cmbStatus = new ComboBox { Location = new Point(20, y), Width = 480, DropDownStyle = ComboBoxStyle.DropDownList };
            _cmbStatus.Items.AddRange(new object[] { "Активен", "Завершен", "Приостановлен" });
            _cmbStatus.SelectedItem = Project.Status;
            if (_cmbStatus.SelectedIndex < 0) _cmbStatus.SelectedIndex = 0;
            y += 35;

            var lblDesc = new Label { Text = "Описание и цели проекта:", Location = new Point(20, y), AutoSize = true };
            y += 25;
            _txtDescription = new TextBox { Location = new Point(20, y), Width = 480, Height = 60, Multiline = true, Text = Project.Description ?? "" };
            y += 75;

            _btnSave = new Button { Text = "Сохранить", Location = new Point(290, y), Width = 100, Height = 32, DialogResult = DialogResult.OK, BackColor = Color.FromArgb(37, 99, 235), ForeColor = Color.White, FlatStyle = FlatStyle.Flat };
            _btnCancel = new Button { Text = "Отмена", Location = new Point(400, y), Width = 100, Height = 32, DialogResult = DialogResult.Cancel };

            _btnSave.Click += BtnSave_Click;

            Controls.AddRange(new Control[] {
                lblName, _txtName, lblCat, _cmbCategory,
                lblTarget, lblCurrent, _numTarget, _numCurrent,
                lblStart, lblEnd, _dtpStart, _chkHasEndDate, _dtpEnd,
                lblStatus, _cmbStatus, lblDesc, _txtDescription,
                _btnSave, _btnCancel
            });

            AcceptButton = _btnSave;
            CancelButton = _btnCancel;
        }

        private void BtnSave_Click(object? sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(_txtName.Text))
            {
                MessageBox.Show("Введите название проекта!", "Валидация", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                DialogResult = DialogResult.None;
                _txtName.Focus();
                return;
            }

            if (_cmbCategory.SelectedValue == null)
            {
                MessageBox.Show("Выберите категорию проекта!", "Валидация", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                DialogResult = DialogResult.None;
                return;
            }

            if (_numTarget.Value <= 0)
            {
                MessageBox.Show("Целевая сумма сбора должна быть больше нуля!", "Валидация", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                DialogResult = DialogResult.None;
                return;
            }

            Project.Name = _txtName.Text.Trim();
            Project.CategoryId = Convert.ToInt32(_cmbCategory.SelectedValue);
            Project.TargetAmount = _numTarget.Value;
            Project.CurrentAmount = _numCurrent.Value;
            Project.StartDate = _dtpStart.Value.Date;
            Project.EndDate = _chkHasEndDate.Checked ? _dtpEnd.Value.Date : null;
            Project.Status = _cmbStatus.SelectedItem?.ToString() ?? "Активен";
            Project.Description = string.IsNullOrWhiteSpace(_txtDescription.Text) ? null : _txtDescription.Text.Trim();
        }
    }
}

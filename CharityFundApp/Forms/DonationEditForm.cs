using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using CharityFundApp.Models;

namespace CharityFundApp.Forms
{
    public class DonationEditForm : Form
    {
        private readonly ComboBox _cmbDonor;
        private readonly ComboBox _cmbProject;
        private readonly NumericUpDown _numAmount;
        private readonly DateTimePicker _dtpDate;
        private readonly ComboBox _cmbPaymentMethod;
        private readonly TextBox _txtNotes;
        private readonly Button _btnSave;
        private readonly Button _btnCancel;

        public Donation Donation { get; private set; }

        public DonationEditForm(Donation? donation, List<Donor> donors, List<Project> projects)
        {
            Donation = donation ?? new Donation();
            Text = donation == null ? "Регистрация пожертвования" : "Редактирование пожертвования";
            Size = new Size(520, 460);
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            Font = new Font("Segoe UI", 9.5f);

            int y = 15;

            var lblDonor = new Label { Text = "Благотворитель (донор)*:", Location = new Point(20, y), AutoSize = true, Font = new Font("Segoe UI", 9.5f, FontStyle.Bold) };
            y += 25;
            _cmbDonor = new ComboBox { Location = new Point(20, y), Width = 460, DropDownStyle = ComboBoxStyle.DropDownList };
            _cmbDonor.DataSource = donors;
            _cmbDonor.DisplayMember = "FullName";
            _cmbDonor.ValueMember = "Id";
            if (Donation.DonorId > 0) _cmbDonor.SelectedValue = Donation.DonorId;
            y += 35;

            var lblProj = new Label { Text = "Проект назначения*:", Location = new Point(20, y), AutoSize = true, Font = new Font("Segoe UI", 9.5f, FontStyle.Bold) };
            y += 25;
            _cmbProject = new ComboBox { Location = new Point(20, y), Width = 460, DropDownStyle = ComboBoxStyle.DropDownList };
            _cmbProject.DataSource = projects;
            _cmbProject.DisplayMember = "Name";
            _cmbProject.ValueMember = "Id";
            if (Donation.ProjectId > 0) _cmbProject.SelectedValue = Donation.ProjectId;
            y += 35;

            var lblAmount = new Label { Text = "Сумма пожертвования (руб.)*:", Location = new Point(20, y), AutoSize = true };
            var lblMethod = new Label { Text = "Способ оплаты*:", Location = new Point(260, y), AutoSize = true };
            y += 25;

            _numAmount = new NumericUpDown { Location = new Point(20, y), Width = 220, DecimalPlaces = 2, Maximum = 100000000m, Value = Donation.Amount > 0 ? Donation.Amount : 1000m };
            _cmbPaymentMethod = new ComboBox { Location = new Point(260, y), Width = 220, DropDownStyle = ComboBoxStyle.DropDownList };
            _cmbPaymentMethod.Items.AddRange(new object[] { "Банковская карта", "Банковский перевод", "Наличные", "Онлайн-платеж" });
            _cmbPaymentMethod.SelectedItem = Donation.PaymentMethod;
            if (_cmbPaymentMethod.SelectedIndex < 0) _cmbPaymentMethod.SelectedIndex = 0;
            y += 35;

            var lblDate = new Label { Text = "Дата и время взноса*:", Location = new Point(20, y), AutoSize = true };
            y += 25;
            _dtpDate = new DateTimePicker { Location = new Point(20, y), Width = 460, CustomFormat = "dd.MM.yyyy HH:mm", Format = DateTimePickerFormat.Custom, Value = Donation.DonationDate };
            y += 35;

            var lblNotes = new Label { Text = "Примечание / Назначение платежа:", Location = new Point(20, y), AutoSize = true };
            y += 25;
            _txtNotes = new TextBox { Location = new Point(20, y), Width = 460, Height = 55, Multiline = true, Text = Donation.Notes ?? "" };
            y += 70;

            _btnSave = new Button { Text = "Сохранить", Location = new Point(270, y), Width = 100, Height = 32, DialogResult = DialogResult.OK, BackColor = Color.FromArgb(37, 99, 235), ForeColor = Color.White, FlatStyle = FlatStyle.Flat };
            _btnCancel = new Button { Text = "Отмена", Location = new Point(380, y), Width = 100, Height = 32, DialogResult = DialogResult.Cancel };

            _btnSave.Click += BtnSave_Click;

            Controls.AddRange(new Control[] {
                lblDonor, _cmbDonor, lblProj, _cmbProject,
                lblAmount, _numAmount, lblMethod, _cmbPaymentMethod,
                lblDate, _dtpDate, lblNotes, _txtNotes,
                _btnSave, _btnCancel
            });

            AcceptButton = _btnSave;
            CancelButton = _btnCancel;
        }

        private void BtnSave_Click(object? sender, EventArgs e)
        {
            if (_cmbDonor.SelectedValue == null)
            {
                MessageBox.Show("Выберите донора!", "Валидация", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                DialogResult = DialogResult.None;
                return;
            }

            if (_cmbProject.SelectedValue == null)
            {
                MessageBox.Show("Выберите проект!", "Валидация", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                DialogResult = DialogResult.None;
                return;
            }

            if (_numAmount.Value <= 0)
            {
                MessageBox.Show("Сумма пожертвования должна быть больше 0!", "Валидация", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                DialogResult = DialogResult.None;
                return;
            }

            Donation.DonorId = Convert.ToInt32(_cmbDonor.SelectedValue);
            Donation.ProjectId = Convert.ToInt32(_cmbProject.SelectedValue);
            Donation.Amount = _numAmount.Value;
            Donation.PaymentMethod = _cmbPaymentMethod.SelectedItem?.ToString() ?? "Банковская карта";
            Donation.DonationDate = _dtpDate.Value;
            Donation.Notes = string.IsNullOrWhiteSpace(_txtNotes.Text) ? null : _txtNotes.Text.Trim();
        }
    }
}

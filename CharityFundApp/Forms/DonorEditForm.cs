using System;
using System.Drawing;
using System.Windows.Forms;
using CharityFundApp.Models;

namespace CharityFundApp.Forms
{
    public class DonorEditForm : Form
    {
        private readonly TextBox _txtName;
        private readonly ComboBox _cmbType;
        private readonly TextBox _txtPhone;
        private readonly TextBox _txtEmail;
        private readonly TextBox _txtAddress;
        private readonly Button _btnSave;
        private readonly Button _btnCancel;

        public Donor Donor { get; private set; }

        public DonorEditForm(Donor? donor = null)
        {
            Donor = donor ?? new Donor();
            Text = donor == null ? "Добавление донора" : "Редактирование донора";
            Size = new Size(480, 390);
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            Font = new Font("Segoe UI", 9.5f);

            int y = 20;

            var lblName = new Label { Text = "ФИО / Наименование организации*:", Location = new Point(20, y), AutoSize = true, Font = new Font("Segoe UI", 9.5f, FontStyle.Bold) };
            y += 25;
            _txtName = new TextBox { Location = new Point(20, y), Width = 420, Text = Donor.FullName };
            y += 35;

            var lblType = new Label { Text = "Тип донора*:", Location = new Point(20, y), AutoSize = true };
            y += 25;
            _cmbType = new ComboBox { Location = new Point(20, y), Width = 420, DropDownStyle = ComboBoxStyle.DropDownList };
            _cmbType.Items.AddRange(new object[] { "Физическое лицо", "Юридическое лицо" });
            _cmbType.SelectedItem = Donor.DonorType;
            if (_cmbType.SelectedIndex < 0) _cmbType.SelectedIndex = 0;
            y += 35;

            var lblPhone = new Label { Text = "Телефон:", Location = new Point(20, y), AutoSize = true };
            y += 25;
            _txtPhone = new TextBox { Location = new Point(20, y), Width = 420, Text = Donor.Phone ?? "" };
            y += 35;

            var lblEmail = new Label { Text = "Email:", Location = new Point(20, y), AutoSize = true };
            y += 25;
            _txtEmail = new TextBox { Location = new Point(20, y), Width = 420, Text = Donor.Email ?? "" };
            y += 35;

            var lblAddress = new Label { Text = "Адрес:", Location = new Point(20, y), AutoSize = true };
            y += 25;
            _txtAddress = new TextBox { Location = new Point(20, y), Width = 420, Text = Donor.Address ?? "" };
            y += 45;

            _btnSave = new Button { Text = "Сохранить", Location = new Point(230, y), Width = 100, Height = 32, DialogResult = DialogResult.OK, BackColor = Color.FromArgb(37, 99, 235), ForeColor = Color.White, FlatStyle = FlatStyle.Flat };
            _btnCancel = new Button { Text = "Отмена", Location = new Point(340, y), Width = 100, Height = 32, DialogResult = DialogResult.Cancel };

            _btnSave.Click += BtnSave_Click;

            Controls.AddRange(new Control[] {
                lblName, _txtName, lblType, _cmbType,
                lblPhone, _txtPhone, lblEmail, _txtEmail,
                lblAddress, _txtAddress, _btnSave, _btnCancel
            });

            AcceptButton = _btnSave;
            CancelButton = _btnCancel;
        }

        private void BtnSave_Click(object? sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(_txtName.Text))
            {
                MessageBox.Show("Введите ФИО или наименование донора!", "Валидация", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                DialogResult = DialogResult.None;
                _txtName.Focus();
                return;
            }

            Donor.FullName = _txtName.Text.Trim();
            Donor.DonorType = _cmbType.SelectedItem?.ToString() ?? "Физическое лицо";
            Donor.Phone = string.IsNullOrWhiteSpace(_txtPhone.Text) ? null : _txtPhone.Text.Trim();
            Donor.Email = string.IsNullOrWhiteSpace(_txtEmail.Text) ? null : _txtEmail.Text.Trim();
            Donor.Address = string.IsNullOrWhiteSpace(_txtAddress.Text) ? null : _txtAddress.Text.Trim();
        }
    }
}

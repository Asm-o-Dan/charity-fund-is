using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using CharityFundApp.Models;

namespace CharityFundApp.Forms
{
    public class RecipientEditForm : Form
    {
        private readonly TextBox _txtName;
        private readonly ComboBox _cmbCategory;
        private readonly TextBox _txtPhone;
        private readonly TextBox _txtAddress;
        private readonly TextBox _txtNeed;
        private readonly ComboBox _cmbStatus;
        private readonly DateTimePicker _dtpRegDate;
        private readonly Button _btnSave;
        private readonly Button _btnCancel;

        public Recipient Recipient { get; private set; }

        public RecipientEditForm(Recipient? recipient, List<Category> categories)
        {
            Recipient = recipient ?? new Recipient();
            Text = recipient == null ? "Добавление получателя помощи" : "Редактирование получателя помощи";
            Size = new Size(520, 510);
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            Font = new Font("Segoe UI", 9.5f);

            int y = 15;

            var lblName = new Label { Text = "ФИО получателя / Организация*:", Location = new Point(20, y), AutoSize = true, Font = new Font("Segoe UI", 9.5f, FontStyle.Bold) };
            y += 25;
            _txtName = new TextBox { Location = new Point(20, y), Width = 460, Text = Recipient.FullName };
            y += 35;

            var lblCat = new Label { Text = "Категория благотворительности*:", Location = new Point(20, y), AutoSize = true };
            y += 25;
            _cmbCategory = new ComboBox { Location = new Point(20, y), Width = 460, DropDownStyle = ComboBoxStyle.DropDownList };
            _cmbCategory.DataSource = categories;
            _cmbCategory.DisplayMember = "Name";
            _cmbCategory.ValueMember = "Id";
            if (Recipient.CategoryId > 0) _cmbCategory.SelectedValue = Recipient.CategoryId;
            y += 35;

            var lblPhone = new Label { Text = "Телефон:", Location = new Point(20, y), AutoSize = true };
            var lblStatus = new Label { Text = "Статус рассмотрения*:", Location = new Point(260, y), AutoSize = true };
            y += 25;

            _txtPhone = new TextBox { Location = new Point(20, y), Width = 220, Text = Recipient.Phone ?? "" };
            _cmbStatus = new ComboBox { Location = new Point(260, y), Width = 220, DropDownStyle = ComboBoxStyle.DropDownList };
            _cmbStatus.Items.AddRange(new object[] { "На рассмотрении", "Одобрен", "Помощь оказана", "Отклонен" });
            _cmbStatus.SelectedItem = Recipient.Status;
            if (_cmbStatus.SelectedIndex < 0) _cmbStatus.SelectedIndex = 0;
            y += 35;

            var lblAddress = new Label { Text = "Адрес проживания / нахождения:", Location = new Point(20, y), AutoSize = true };
            y += 25;
            _txtAddress = new TextBox { Location = new Point(20, y), Width = 460, Text = Recipient.Address ?? "" };
            y += 35;

            var lblNeed = new Label { Text = "Описание потребности в помощи*:", Location = new Point(20, y), AutoSize = true };
            y += 25;
            _txtNeed = new TextBox { Location = new Point(20, y), Width = 460, Height = 60, Multiline = true, Text = Recipient.NeedDescription };
            y += 70;

            var lblDate = new Label { Text = "Дата регистрации:", Location = new Point(20, y), AutoSize = true };
            _dtpRegDate = new DateTimePicker { Location = new Point(160, y - 3), Width = 150, Format = DateTimePickerFormat.Short, Value = Recipient.RegistrationDate };
            y += 40;

            _btnSave = new Button { Text = "Сохранить", Location = new Point(270, y), Width = 100, Height = 32, DialogResult = DialogResult.OK, BackColor = Color.FromArgb(37, 99, 235), ForeColor = Color.White, FlatStyle = FlatStyle.Flat };
            _btnCancel = new Button { Text = "Отмена", Location = new Point(380, y), Width = 100, Height = 32, DialogResult = DialogResult.Cancel };

            _btnSave.Click += BtnSave_Click;

            Controls.AddRange(new Control[] {
                lblName, _txtName, lblCat, _cmbCategory,
                lblPhone, _txtPhone, lblStatus, _cmbStatus,
                lblAddress, _txtAddress, lblNeed, _txtNeed,
                lblDate, _dtpRegDate, _btnSave, _btnCancel
            });

            AcceptButton = _btnSave;
            CancelButton = _btnCancel;
        }

        private void BtnSave_Click(object? sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(_txtName.Text))
            {
                MessageBox.Show("Введите ФИО или наименование получателя!", "Валидация", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                DialogResult = DialogResult.None;
                _txtName.Focus();
                return;
            }

            if (_cmbCategory.SelectedValue == null)
            {
                MessageBox.Show("Выберите категорию помощи!", "Валидация", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                DialogResult = DialogResult.None;
                return;
            }

            if (string.IsNullOrWhiteSpace(_txtNeed.Text))
            {
                MessageBox.Show("Укажите описание потребности в помощи!", "Валидация", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                DialogResult = DialogResult.None;
                _txtNeed.Focus();
                return;
            }

            Recipient.FullName = _txtName.Text.Trim();
            Recipient.CategoryId = Convert.ToInt32(_cmbCategory.SelectedValue);
            Recipient.Phone = string.IsNullOrWhiteSpace(_txtPhone.Text) ? null : _txtPhone.Text.Trim();
            Recipient.Address = string.IsNullOrWhiteSpace(_txtAddress.Text) ? null : _txtAddress.Text.Trim();
            Recipient.NeedDescription = _txtNeed.Text.Trim();
            Recipient.Status = _cmbStatus.SelectedItem?.ToString() ?? "На рассмотрении";
            Recipient.RegistrationDate = _dtpRegDate.Value.Date;
        }
    }
}

using System;
using System.Drawing;
using System.Windows.Forms;
using CharityFundApp.Models;

namespace CharityFundApp.Forms
{
    public class CategoryEditForm : Form
    {
        private readonly TextBox _txtName;
        private readonly TextBox _txtDescription;
        private readonly Button _btnSave;
        private readonly Button _btnCancel;

        public Category Category { get; private set; }

        public CategoryEditForm(Category? category = null)
        {
            Category = category ?? new Category();
            Text = category == null ? "Добавление категории" : "Редактирование категории";
            Size = new Size(460, 280);
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            Font = new Font("Segoe UI", 9.5f);

            var lblName = new Label { Text = "Название категории*:", Location = new Point(20, 20), AutoSize = true, Font = new Font("Segoe UI", 9.5f, FontStyle.Bold) };
            _txtName = new TextBox { Location = new Point(20, 45), Width = 400, Text = Category.Name };

            var lblDesc = new Label { Text = "Описание:", Location = new Point(20, 85), AutoSize = true };
            _txtDescription = new TextBox { Location = new Point(20, 110), Width = 400, Height = 65, Multiline = true, Text = Category.Description ?? "" };

            _btnSave = new Button { Text = "Сохранить", Location = new Point(210, 195), Width = 100, Height = 32, DialogResult = DialogResult.OK, BackColor = Color.FromArgb(37, 99, 235), ForeColor = Color.White, FlatStyle = FlatStyle.Flat };
            _btnCancel = new Button { Text = "Отмена", Location = new Point(320, 195), Width = 100, Height = 32, DialogResult = DialogResult.Cancel };

            _btnSave.Click += BtnSave_Click;

            Controls.AddRange(new Control[] { lblName, _txtName, lblDesc, _txtDescription, _btnSave, _btnCancel });
            AcceptButton = _btnSave;
            CancelButton = _btnCancel;
        }

        private void BtnSave_Click(object? sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(_txtName.Text))
            {
                MessageBox.Show("Введите название категории!", "Валидация", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                DialogResult = DialogResult.None;
                _txtName.Focus();
                return;
            }

            Category.Name = _txtName.Text.Trim();
            Category.Description = string.IsNullOrWhiteSpace(_txtDescription.Text) ? null : _txtDescription.Text.Trim();
        }
    }
}

using System;
using System.Windows.Forms;
using CharityFundApp.Forms;

namespace CharityFundApp
{
    static class Program
    {
        /// <summary>
        /// Главная точка входа для приложения «Информационная система Благотворительный фонд».
        /// </summary>
        [STAThread]
        static void Main()
        {
            ApplicationConfiguration.Initialize();
            Application.Run(new MainForm());
        }
    }
}
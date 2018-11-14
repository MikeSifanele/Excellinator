using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Threading.Tasks;

namespace Excellinator
{
    class Form1x
    {
        public void Start()
        {
            if (DateTime.Now > DateTime.Parse("03/02/2019"))
            {
                MessageBox.Show("Application trial expired.\nPlease contact Justice", "Access denied");
                Application.Exit();
            }
        }
    }
}

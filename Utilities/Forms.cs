using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Utilities
{
    class Forms
    {
        public Boolean confirmationBox(string message, string title)
        {
            bool confirm = false;
            DialogResult result = MessageBox.Show(message, title, MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2);
            if (result == DialogResult.Yes)
            {
                confirm = true;
            }
            return confirm;
        }
    }
}

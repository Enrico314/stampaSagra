using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace StampaSagra
{
    static class Program
    {
        /// <summary>
        /// Punto di ingresso principale dell'applicazione.
        /// </summary>
        [STAThread]
        static void Main()
        {

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
#if DEBUG
            double timeout_second = 500;
#else
            double timeout_second = 5;
#endif

            // lo faccio andare per quel che è settato
            Form1 form1 = new Form1();
            Task task = Task.Run(() => Application.Run(form1));

            /* commendate da qui per debaggure altrimenti si chiude dopo 5s*/
            
            if (task.Wait(TimeSpan.FromSeconds(timeout_second)))
            {

            }

            else
            {
                //ucciditi!
                form1.killMyself();
            }
            

        }
    }
}

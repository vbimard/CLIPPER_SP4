using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using AF_Import_ODBC_Clipper_AlmaCam;
using AF_ImportTools;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Wpm.Implement.Manager;



namespace ConsoleApp2
{

   

    class Program
    {
        static void Main(string[] args)
        {
            bool import_matiere = false;
            bool tube_rond = false;
            bool rond = false;
            bool tube_rectangle = false;
            bool tube_carre = false;
            bool tube_flat = false;
            bool tube_speciaux = false;
            bool Vis = true;

            /*
            Clipper_ImportTubes_Processor tubeimporter = new Clipper_ImportTubes_Processor(import_matiere,  tube_rond,  rond,  tube_rectangle,  tube_carre, tube_flat,  tube_speciaux);
            //tubeimporter.Execute();
            tubeimporter.Execute();*/
            IContext contextlocal = null;
            contextlocal=AlmaCamTool.GetContext(contextlocal);
            Clipper_Import_Fournitures_Divers_Processor FournitureImporter = new Clipper_Import_Fournitures_Divers_Processor();
            FournitureImporter.Import(contextlocal);



            //tubeimporter.Execute();

            //VisImporter.Execute();

            //VisImporter.ExecuteContext(IContext);

        }
    }
}

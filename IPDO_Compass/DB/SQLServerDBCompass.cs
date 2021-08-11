using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IPDO_Compass.DB
{
    public class SQLServerDBCompass : SQLServerDB
    {
        public SQLServerDBCompass(string banco = "IPDO",string server = "local")
        {
            string endereco = "10.206.194.187";
            
            //this.SetPorta(1433);
            if (server == "local")
            {
                endereco = "10.206.194.187";
                this.SetUsuario("sa");


            }
            else if (server == "azure")
            {

                endereco = "bdcompass.database.windows.net";
                this.SetUsuario("compass");
            }
            this.SetServidor(endereco);
            //this.SetServidor("bdcompass.database.windows.net");
            //this.SetPorta(1433);
            //  this.SetUsuario("compass");
            this.SetSenha(this.GetPassword(this.GetUsuario()));
            this.SetDatabase(this.GetDatabase(banco));
        }

        private string GetPassword(string p_strUsuario)
        {
            switch (p_strUsuario)
            {
                case "sa":
                    return "cp@s9876";
                case "captura":
                    return "c@ptura9876";

                case "captura_read":
                    return "captur@leitur@";
                case "compass":
                    return "cpas#9876";

                default:
                    return "";
            }
        }

        private string GetDatabase(string p_banco)
        {
            switch (p_banco)
            {
                case "IPDO":
                    return "IPDO";
                case "ESTUDO_PV":
                    return "ESTUDO_PV";
                case "CHUVAS":
                    return "CHUVAS";



                default:
                    return "";
            }
        }
    }
}

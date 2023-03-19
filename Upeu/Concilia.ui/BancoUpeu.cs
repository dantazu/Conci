using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Concilia.ui
{
    public class BancoUpeu
    {
        public string NroOpe { get; set; }

        public DateTime FechaRegistro { get; set; }

        public string ReferenciaLibros { get; set; }

        public string Descripcion { get; set; }

        public string FechaOperacion { get; set; }

        public decimal Importe { get; set; }

        public int Dh { get; set; }

        public string NivelConta { get; set; }

        public decimal Saldo { get; set; }

        public string CodigoPos { get; set; }

        public bool Pendiente { get; set; }

        public string Whoyo { get; set; }
    }
}

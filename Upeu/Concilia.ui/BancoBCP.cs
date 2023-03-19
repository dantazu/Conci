using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Concilia.ui
{
    public class BancoBCP : IEquatable<BancoBCP>
    {
        public DateTime FechaOperacion { get; set; }

        public string NroOpe { get; set; }

        public string Descripcion { get; set; }

        public decimal Importe { get; set; }

        public decimal Diferencia { get; set; }

        public decimal NetoAbonar { get; set; }

        public int Dh { get; set; }

        public string CodigoPos { get; set; }

        public string Whoyo { get; set; }

        public string Terminal { get; set; }

        public string FechaAbono { get; set; }

        public string ReferenciaVoucher { get; set; }

        public bool Equals(BancoBCP other)
        {
            if (other == null)
            {
                return false;
            }
            return NroOpe == other.NroOpe && CodigoPos == other.CodigoPos && Importe == other.Importe;
        }

        public override bool Equals(object obj)
        {
            return Equals(obj as BancoBCP);
        }

        public override int GetHashCode()
        {
            return (NroOpe, CodigoPos, Importe).GetHashCode();
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MapaReduce.Modelos
{
    internal class mdlCategorias
    {
        private String Nombre;
        private int Min;

        public String getName()
        {
            return this.Nombre;
        }
        public void setName(String Nombre)
        {
            this.Nombre  = Nombre ;
        }
        public int getMin()
        {
            return this.Min;
        }
        public void setMin(int Min)
        {
            this.Min = Min;
        }
    }
    
}

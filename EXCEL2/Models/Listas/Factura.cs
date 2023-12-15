using Microsoft.Ajax.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace EXCEL2.Models.Listas
{
    public class Factura
    {
        public object SERIE { get; set; }
        public object CORRELATIVO { get; set; }
        public object TIPO_DE_COMPROBANTE { get; set; }
        public object RAZON_SOCIAL { get; set; }
        public object CUENTA { get; set; }
    }
}
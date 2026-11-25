# -*- coding: utf-8 -*-
{
    'name': "Exportador Facturas XLSX",

    'summary': """
        Libro XLSX de facturas agrupadas por Contacto con impuestos en columnas.
        """,

    'description': """
Exportación de facturas a XLSX agrupadas por Contacto
=====================================================

Añade un informe en formato Excel (.xlsx) que se puede imprimir desde las
facturas de cliente y de proveedor seleccionadas, con el desglose de bases
imponibles e impuestos de cada documento.

Funcionalidades
---------------
* Agrupación de las facturas por contacto comercial (NIF y nombre), ordenado
  alfabéticamente.
* Una columna por cada impuesto aplicado en las facturas seleccionadas
  (IVA, recargo de equivalencia, retenciones, etc.), incluso cuando su
  importe es cero. Primero se muestran los impuestos positivos, de mayor a
  menor porcentaje, y después las retenciones.
* Columnas de Base Imponible y Total Factura por documento.
* Separación de facturas y rectificativas dentro de cada partner, con sus
  subtotales correspondientes.
* Totales por partner y totales generales (facturas, rectificativas y total
  general) calculados mediante fórmulas de Excel.
* Cabecera informativa con la empresa y el rango de fechas de las facturas
  incluidas.
* Se ignoran automáticamente los asientos que no son facturas y las facturas
  canceladas.
    """,

    'author': "motsoft",
    'website': "",

    # Categories can be used to filter modules in modules listing
    # Check https://github.com/odoo/odoo/blob/15.0/odoo/addons/base/data/ir_module_category_data.xml
    # for the full list
    'category': 'Invoicing Management',
    'version': '15.0.0.2',

    'license': 'LGPL-3',
    # any module necessary for this one to work correctly
    'depends': ['account', 'report_xlsx'],

    # always loaded
    'data': [
        'report/invoice_xlsx.xml',
    ],
    'installable': True,
    'application': False,
}

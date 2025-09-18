{
    'name': 'Importación de Operaciones Bancarias',
    'version': '18.0.1.0.0',
    'category': 'Accounting',
    'summary': 'Importar y procesar archivos bancarios TXT y Excel',
    'description': '''
        Módulo para importar operaciones bancarias desde archivos TXT y Excel,
        realizar matching con pagos existentes en el sistema basado en número
        de operación y monto para validación.
    ''',
    'depends': ['account', 'base'],
    'data': [
        'views/bank_import_views.xml',
        # 'wizards/bank_import_wizard_views.xml',
        'views/menu_views.xml',
        'security/ir.model.access.csv',
    ],
    'installable': True,
    'application': False,
    'auto_install': False,
}
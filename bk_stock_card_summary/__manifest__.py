{
    "name": "Stock In/Out Report (Excel Export)",
    "summary": "Generate detailed stock movement (In/Out) reports with initial balance, purchases, returns, sales, adjustments, and ending balance.",
    "description": """
Stock In/Out Report with Excel Export
=====================================

This module provides a custom stock movement analysis report for Odoo.
    """,
    "version": "1.0.0",
    "author": "Biniyam K|info.biniyamkg@gmail.com",
    "website": "https://yourcompany.com",
    "category": "Inventory/Reporting",
    "license": "LGPL-3",
    "depends": ["stock", "web", "account", "point_of_sale"],
    "data": [
        "security/ir.model.access.csv",
        "wizard/bk_stock_inout_wizard_view.xml",
        "views/account_account_view.xml",
    ],
    "assets": {},
    "application": False,
    "installable": True,
    "auto_install": False,
}

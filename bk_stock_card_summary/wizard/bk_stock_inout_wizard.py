from odoo import api, fields, models


class StockInOutWizard(models.TransientModel):
    _name = "bk.stock.inout.wizard"
    _description = "Stock In/Out Report Wizard"

    date_start = fields.Date("Start Date", required=True)
    date_end = fields.Date("End Date", required=True)
    location_ids = fields.Many2many("stock.location", string="Locations", domain=[('usage', '=', 'internal')])
    pos_config_ids = fields.Many2many("pos.config", string="Excluded PoS")
    categ_ids = fields.Many2many("product.category", string="Product Categories")
    product_ids = fields.Many2many("product.product", string="Products")
    state = fields.Selection([
        ('done', 'Done'),
        ('ready', 'Ready'),
        ('all', 'All')
    ], string="Move Status", default='done', required=True)
    report_type = fields.Selection([
        ('detailed', 'Detailed Movement'),
        ('summary', 'Forecast'),
        ('valuation', 'Valuation')
    ], string="Report Type", default="detailed", required=True)

    def action_export_excel(self):
        return {
            "type": "ir.actions.act_url",
            "url": f"/web/binary/export_xlsx?wizard_id={self.id}",
            "target": "self",
        }
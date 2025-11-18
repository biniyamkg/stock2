from odoo import fields, models, api


class AccountAccount(models.Model):
    _inherit = 'account.account'
    _description = 'Account'

    is_cogs_account = fields.Boolean(string="Is COGS Account")
    is_inventory_account = fields.Boolean(string="Is Inventory Account")


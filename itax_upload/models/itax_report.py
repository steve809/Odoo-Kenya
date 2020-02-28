# -*- coding: utf-8 -*-
##############################################################################
# Part of Hyperthink Systems Limited (www.hyperthinkkenya.co.ke)
#
##############################################################################

from odoo import api, fields, models, _
from odoo.exceptions import UserError, ValidationError

class ResPartner(models.Model):
    _inherit = 'res.partner'

    customer_flag = fields.Selection([('import', 'IMPORT'), ('local', 'LOCAL')], string='Customer Flag')
    
class AccountMove(models.Model):
    _inherit = "account.move"

    custom_entry_number = fields.Char('Custom Entry Number')

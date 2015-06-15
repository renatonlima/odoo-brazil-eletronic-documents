# -*- encoding: utf-8 -*-
###############################################################################
#                                                                             #
# Copyright (C) 2015 Luis Felipe Mileo - KMEE                                 #
#                                                                             #
#This program is free software: you can redistribute it and/or modify         #
#it under the terms of the GNU Affero General Public License as published by  #
#the Free Software Foundation, either version 3 of the License, or            #
#(at your option) any later version.                                          #
#                                                                             #
#This program is distributed in the hope that it will be useful,              #
#but WITHOUT ANY WARRANTY; without even the implied warranty of               #
#MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the                #
#GNU General Public License for more details.                                 #
#                                                                             #
#You should have received a copy of the GNU General Public License            #
#along with this program.  If not, see <http://www.gnu.org/licenses/>.        #
###############################################################################

import logging
from openerp import tools
from openerp.osv import fields, osv
from openerp.tools.translate import _
import time
from openerp.tools import float_is_zero
import openerp.addons.decimal_precision as dp
import openerp.addons.product.product

from ..sped.nfce.nfce_factory import NfCeFactory
from openerp.addons.nfe.sped.nfe.processing.xml import send
from ..sped.nfce.validator.config_check import validate_nfce_configuration

_logger = logging.getLogger(__name__)

#
# class res_partner(osv.osv):
#
#     _inherit = "res.partner"
#
#     _columns = {
#         'leticia' : fields.char(u'Leticiaaaa'),
#         }


class pos_order(osv.osv):

    _inherit = "pos.order"

    _columns = {
        'account_document_event_ids': fields.one2many(
            'l10n_br_account.document_event', 'document_event_ids',
            u'Eventos'),
        }

    def _get_nfe_factory(self, nfe_version):
        return NfCeFactory().get_nfe(nfe_version)

    def action_nfce(self, cr, uid, ids, context=None):
        event_obj = self.pool.get('l10n_br_account.document_event')
        product_obj = self.pool.get('product.product')
        nfce_ids = []

        for order in self.pool.get('pos.order').browse(cr, uid, ids, context=context):
            if order.invoice_id:
                raise osv.except_osv(_('Error!'), _(
                    'Não é permitida a Emissão de NFC-E para uma ordem com '
                    'NF-E já emitida'))

            company = order.company_id

            validate_nfce_configuration(company)
            nfe_obj = self._get_nfe_factory(company.nfe_version)

            nfes = nfe_obj.get_nfce(cr, uid, ids, int(company.nfe_environment), context)
            xml = nfes[0].xml
            eventos = [ ]
            for processo in send(company, nfes):
               vals = {
                            'type': str(processo.webservice),
                            'status': processo.resposta.cStat.valor,
                            'response': '',
                            'company_id': company.id,
                  #          'origin': '[NF-E]' + inv.internal_number,
                            #TODO: Manipular os arquivos manualmente
                            # 'file_sent': processo.arquivos[0]['arquivo'],
                            # 'file_returned': processo.arquivos[1]['arquivo'],
                            'message': processo.resposta.xMotivo.valor,
                            'state': 'done',
                   #         'document_event_ids': inv.id
               }
               eventos.append(vals)
            return True


    def create_from_ui(self, cr, uid, orders, context=None):
        # Keep only new orders
        submitted_references = [o['data']['name'] for o in orders]
        existing_order_ids = self.search(cr, uid, [('pos_reference', 'in', submitted_references)], context=context)
        existing_orders = self.read(cr, uid, existing_order_ids, ['pos_reference'], context=context)
        existing_references = set([o['pos_reference'] for o in existing_orders])
        orders_to_save = [o for o in orders if o['data']['name'] not in existing_references]

        order_ids = []

        for tmp_order in orders_to_save:
            to_invoice = tmp_order['to_invoice']
            order = tmp_order['data']
            order_id = self._process_order(cr, uid, order, context=context)
            order_ids.append(order_id)

            try:
                self.signal_workflow(cr, uid, [order_id], 'paid')
            except Exception as e:
                _logger.error('Could not fully process the POS Order: %s', tools.ustr(e))

            if to_invoice:
                self.action_invoice(cr, uid, [order_id], context)
                order_obj = self.browse(cr, uid, order_id, context)
                self.pool['account.invoice'].signal_workflow(cr, uid, [order_obj.invoice_id.id], 'invoice_open')
            else:
                self.action_nfce(cr, uid, [order_id], context)
        return order_ids
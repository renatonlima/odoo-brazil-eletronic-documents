/**
 * Created by mileo on 27/05/15.
 */

openerp.nfce = function (instance) {

    var module = instance.point_of_sale;

    var QWeb = instance.web.qweb;
    var _t = instance.web._t;

    var round_di = instance.web.round_decimals;
    var round_pr = instance.web.round_precision;

    Backbone.Model.prototype._super = function (funcName) {
        return this.constructor.__super__[funcName].apply(this, _.rest(arguments));
    };

    module.Order = module.Order.extend({
        initialize: function(attributes){
            Backbone.Model.prototype.initialize.apply(this, arguments);
            this.pos = attributes.pos;
            this.sequence_number = this.pos.pos_session.sequence_number++;
            this.uid =     this.generateUniqueId();
            this.set({
                creationDate:   new Date(),
                orderLines:     new module.OrderlineCollection(),
                paymentLines:   new module.PaymentlineCollection(),
                name:           _t("Order ") + this.uid,
                client:         null,
            });
            this.selected_orderline   = undefined;
            this.selected_paymentline = undefined;
            this.screen_data = {};  // see ScreenSelector
            this.receipt_type = 'receipt';  // 'receipt' || 'invoice'
            this.temporary = attributes.temporary || false;
            // l10n_br_openerp_pos_models
            this.chNFe = 'chave123456';
            this.nVersao = null;
            this.tpAmb = this.pos.company.nfe_environment;
            this.cDest = null;
            this.dhEmi = null;
            this.vNF = null;
            this.vICMS = null;
            this.digVal = null;
            this.cldToken = null;
            this.cHashQR = null;
            this.legal_name = this.pos.company.legal_name;
            //
            this.internal_number = null;
            this.serie_nfce = null;
            this.protocolo = null;
            this.vTribTotal = null;
            this.qtItens = null;

            return this;
        },

        //get_vICMS: function () {
        //    var vICMS = 0.0;
        //    var pos = this.pos;
        //    this.get('orderLines').each(function (line) {
        //        var ldetails = line.get_tax_details();
        //        console.log(JSON.stringify(this));
        //        console.log(JSON.stringify(this));
        //        for (var id in ldetails) {
        //            if (ldetails.hasOwnProperty(id)) {
        //                if (pos.taxes_by_id[id] == 'icms') {
        //                    vICMS += ldetails[id]
        //                }
        //            }
        //        }
        //    });
        //    return vICMS;
        //},

        get_qtItens: function () {
            var qtItens = 0;
            this.get('orderLines').each(function(line){
                var lineQtd = line.get_quantity();
                qtItens += lineQtd;
            });
            return qtItens;
        },
        export_as_JSON: function() {
            var orderLines, paymentLines;
            orderLines = [];
            (this.get('orderLines')).each(_.bind( function(item) {
                return orderLines.push([0, 0, item.export_as_JSON()]);
            }, this));
            paymentLines = [];
            (this.get('paymentLines')).each(_.bind( function(item) {
                return paymentLines.push([0, 0, item.export_as_JSON()]);
            }, this));
            return {
                name: this.getName(),
                amount_paid: this.getPaidTotal(),
                amount_total: this.getTotalTaxIncluded(),
                amount_tax: this.getTax(),
                amount_return: this.getChange(),
                lines: orderLines,
                statement_ids: paymentLines,
                pos_session_id: this.pos.pos_session.id,
                partner_id: this.get_client() ? this.get_client().id : false,
                user_id: this.pos.cashier ? this.pos.cashier.id : this.pos.user.id,
                uid: this.uid,
                sequence_number: this.sequence_number,
                vICMS: 6,
                qtItens: this.get_qtItens(),
                chNFe: this.chNFe,
            };
        }

    });

    //var _initialize_ = module.PosModel.prototype.initialize;
    //module.PosModel.prototype.initialize = function(session, attributes){
    //    self = this;
    //    model = {
    //        model: 'product.template',
    //        fields: [
    //            'name',
    //            'display_name',
    //            'product_variant_ids',
    //            'product_variant_count',
    //            ],
    //        domain:  function(self){
    //            return [
    //                ['sale_ok','=',true],
    //                ['available_in_pos','=',true],
    //            ];},
    //        context: function(self){
    //            return {
    //                pricelist: self.pricelist.id,
    //                display_default_code: false,
    //            };},
    //        loaded: function(self, templates){
    //             self.db.add_templates(templates);
    //        },
    //    }
    //    this.models.push(model);
    //    return _initialize_.call(this, session, attributes);
    //};

    module.PosModel.prototype.models.push({
         model:  'res.company',
            fields: [ 'currency_id', 'email', 'website', 'company_registry', 'vat', 'name', 'phone', 'partner_id' , 'country_id','legal_name', 'nfe_environment'],
            ids:    function(self){ return [self.user.company_id[0]] },
            loaded: function(self,companies) {
                self.company = companies[0];
            }
    });
    //teste de criacao de botao
    //module.NfcePaymentScreenWidget = module.PaymentScreenWidget.include({
    //    template: 'PaymentScreenWidget',
    //    show: function() {
    //        this._super();
    //        var self = this;
    //
    //        if (this.pos.config.iface_nfce) {
    //            this.add_action_button({
    //                label: _t('NFC-e'),
    //                name: 'nfce',
    //                icon: '/point_of_sale/static/src/img/icons/png48/invoice.png',
    //                click: function () {
    //                    self.validate_order({invoice: true});
    //                }
    //            });
    //        }
    //    }
    //});
};




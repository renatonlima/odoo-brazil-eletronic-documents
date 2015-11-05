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
            this.cDest = null;
            this.dhEmi = null;
            this.vNF = null;
            this.vICMS = null;
            this.digVal = null;
            this.cldToken = null;
            this.cHashQR = null;
            this.infCpl = null;
            //
            this.internal_number = null;
            this.serie_nfce = null;
            this.protocolo = null;
            this.vTribTotal = null;

            this.qtItens = null;
            this.tpAmb = this.pos.company.nfe_environment;
            this.cnpj = this.pos.company.cnpj_cpf;
            this.inscr_est = this.pos.company.inscr_est;
            this.inscr_mun = this.pos.company.inscr_mun;
            this.legal_name = this.pos.company.legal_name;
            this.company_street = this.pos.company.street;
            this.company_number = this.pos.company.number;
            this.company_street2 = this.pos.company.street2;
            this.company_district = this.pos.company.district;
            this.company_state = this.pos.company.state_id[1];
            this.company_city = this.pos.company.l10n_br_city_id[1];
            this.company_zip = this.pos.company.zip;

            return this;
        },

        get_vTribTotal: function () {
            var vTribTotal = 0.0;
            var pos = this.pos;
            this.get('orderLines').each(function (line) {
                var ldetails = line.get_tax_details();
                for (var id in ldetails) {
                    vTribTotal += ldetails[id]
                }
            });
            return vTribTotal;
        },

        get_mensagemFiscal: function (){
            var mensagemFiscal = '';
            var tpAmb = this.tpAmb;

            if (tpAmb == 2){
                mensagemFiscal = "EMITIDA EM AMBIENTE DE HOMOLOGAÇÃO – SEM VALOR FISCAL";
            }
            if (tpAmb == 9){
                mensagemFiscal = "EMITIDA EM CONTINGÊNCIA"
            }
            return mensagemFiscal;
        },

        get_danfe_dhEmi: function (){
            var data = new Date();
            var dhEmi = data.toLocaleDateString()+ ' '+ data.toLocaleTimeString();
            return dhEmi;
        },

        get_qtItens: function () {
            var qtItens = 0;
            this.get('orderLines').each(function(line){
                var lineQtd = line.get_quantity();
                qtItens += lineQtd;
            });
            return qtItens;
        },

        get_via: function () {
            var via = '';
            if (this.tpAmb == 9){
                via = {1: "Via do Consumidor", 2: "Via Estabelecimento"};
            }
            return via;
        },

        get_endEletronico: function () {
            var endEletronico = '';
            if (this.company_state == "Rio Grande do Sul") {
                endEletronico = 'www.sefaz.rs.gov.br';
            }
            return endEletronico;
        },

        get_company_address: function() {
            var company_address = '';
            if (this.company_street) {
                company_address += this.company_street;
            }
             if (this.company_number) {
                company_address += ", "+this.company_number;
            }
             if (this.company_street2) {
                company_address += ", "+this.company_street2;
            }
             if (this.company_district) {
                company_address += ", "+ this.company_district;
            }
             if (this.company_city) {
                company_address +=", "+this.company_city;
            }
             if (this.company_state) {
                company_address +=", "+this.company_state;
            }
             if (this.company_zip) {
                 company_address += " - "+this.company_zip;
             }
            return company_address
        },

    });
    var _super_export_for_printing_ = module.Order.prototype.export_for_printing;
    module.Order.prototype.export_for_printing = function(){
        res = _super_export_for_printing_.call(this);
        res.vTribTotal = this.get_vTribTotal();
        res.company.address = this.get_company_address();
        res.chNFe = this.chNFe;
        res.nVersao = this.nVersao;
        res.cDest = this.cDest;
        res.dhEmi = this.dhEmi;
        res.vNF = this.vNF;
        res.vICMS = this.vICMS;
        res.digVal = this.digVal;
        res.cldToken = this.cldToken;
        res.cHashQR = this.cHashQR;
        res.infCpl = this.infCpl;
        res.internal_number = this.internal_number;
        res.serie_nfce = this.serie_nfce;
        res.protocolo = this.protocolo;
        res.qtItens = this.get_qtItens();
        res.tpAmb = this.tpAmb;
        res.company.cnpj = this.cnpj;
        res.company.inscr_est = this.inscr_est;
        res.mensagemFiscal = this.get_mensagemFiscal();
        res.via = this.get_via();
        res.endEletronico = this.get_endEletronico();
        res.dhEmi = this.get_danfe_dhEmi();
        res.company.inscr_mun = this.inscr_mun;
        res.company.legal_name = this.legal_name;
        console.log(JSON.stringify(res));
        return res;
    };


    //modificação para exibir o recibo ao inves do ticket
    var _super_receipt_refresh_ = module.ReceiptScreenWidget.prototype.refresh;
    module.ReceiptScreenWidget.prototype.refresh = function() {
        var order = this.pos.get('selectedOrder');
        var receipt = order.export_for_printing();
        $('.pos-receipt-container', this.$el).html(QWeb.render('XmlReceipt',{
                        receipt: receipt, widget: self,
                    }));
    };

    module.PosModel.prototype.models.push({
         model:  'res.company',
            fields: [ 'currency_id', 'email', 'website', 'company_registry', 'vat', 'name', 'phone', 'partner_id' ,
                'country_id', 'legal_name', 'nfe_environment', 'cnpj_cpf', 'inscr_est', 'inscr_mun',
                'street', 'number', 'street2', 'district', 'state_id', 'l10n_br_city_id', 'zip', 'country_id' ],
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

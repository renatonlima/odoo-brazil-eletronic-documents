#coding=utf-8
###############################################################################
#                                                                             #
# Copyright (C) 2015  Danimar Ribeiro www.trustcode.com.br                    #
#                                                                             #
#This program is free software: you can redistribute it and/or modify         #
#it under the terms of the GNU Affero General Public License as published by  #
#the Free Software Foundation, either version 3 of the License, or            #
#(at your option) any later version.                                          #
#                                                                             #
#This program is distributed in the hope that it will be useful,              #
#but WITHOUT ANY WARRANTY; without even the implied warranty of               #
#MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the                #
#GNU Affero General Public License for more details.                          #
#                                                                             #
#You should have received a copy of the GNU Affero General Public License     #
#along with this program.  If not, see <http://www.gnu.org/licenses/>.        #
###############################################################################
import re
import os
import base64
import gzip
import cStringIO
from pysped.nfe import ProcessadorNFe
from datetime import datetime

def __processo(company):
    
    p = ProcessadorNFe()
    p.ambiente = int(company.nfe_environment)
    p.estado = company.partner_id.l10n_br_city_id.state_id.code
    p.certificado.stream_certificado = base64.decodestring(company.nfe_a1_file)
    p.certificado.senha = company.nfe_a1_password
    p.salvar_arquivos      = True
    p.contingencia_SCAN   = False
    p.caminho = company.nfe_export_folder
    return p


def _format_nsu(nsu):
    nsu = long(nsu)
    return "%015d" % (nsu,)

def _create_dirs(company):
    caminho = company.nfe_export_folder
    ambiente = int(company.nfe_environment)
    if ambiente == 1:
        caminho = os.path.join(caminho, 'producao/dfe-resumo/')
    else:
        caminho = os.path.join(caminho, 'homologacao/dfe-resumo/')
    caminho = os.path.join(caminho, datetime.now().strftime('%Y-%m') + '/')
    try:
        os.makedirs(caminho)
    except:
        pass
    return caminho

def distribuicao_nfe(company, ultimo_nsu):
    ultimo_nsu = _format_nsu(ultimo_nsu)
    p = __processo(company)
    cnpj_partner = re.sub('[^0-9]','', company.cnpj_cpf)   
    result = p.consultar_distribuicao(cnpj_cpf=cnpj_partner, ultimo_nsu=ultimo_nsu, nsu='')
    
    if result.resposta.status == 200: #Webservice ok
        if result.resposta.cStat.valor == '137' or result.resposta.cStat.valor == '138':
            
            nfe_list = []
            save_path =_create_dirs(company)
            for doc in result.resposta.loteDistDFeInt.docZip:                        
                orig_file_desc = gzip.GzipFile(mode='r', 
                              fileobj=cStringIO.StringIO(base64.b64decode(doc.base64Binary.valor)))
                orig_file_cont = orig_file_desc.read()
                orig_file_desc.close()                
                
                path = os.path.join(save_path, 'resumo_nfe-' +  doc.NSU.valor + '.xml')
                arq = open(path, 'w')
                arq.write(orig_file_cont.encode('utf-8'))
                arq.close()       
                
                nfe_list.append({ 'path':path, 'xml': orig_file_cont, 'NSU': doc.NSU.valor, 
                                 'schema': doc.schema.valor })                         
            
            return { 'code': result.resposta.cStat.valor, 'message': result.resposta.xMotivo.valor,
               'list_nfe': nfe_list, 'file_returned': result.resposta.xml}
        else:
            return { 'code': result.resposta.cStat.valor, 'message': result.resposta.xMotivo.valor,
               'file_sent': result.envio.xml, 'file_returned': result.resposta.xml }
    else:
        return { 'code': result.resposta.status, 'message': result.resposta.reason, 
                'file_sent': result.envio.xml, 'file_returned': None }
    
def send_event(company, nfe_key, method):
    p = __processo(company)
    cnpj_partner = re.sub('[^0-9]','', company.cnpj_cpf)
    result = {}
    if method == 'ciencia_operacao':    
        result = p.conhecer_operacao_evento(
            cnpj=cnpj_partner, # CNPJ do destinatário/gerador do evento
            chave_nfe=nfe_key )
    elif method == 'confirma_operacao':
        result = p.confirmar_operacao_evento(
            cnpj=cnpj_partner, 
            chave_nfe=nfe_key )
    elif method == 'desconhece_operacao':
        result = p.desconhecer_operacao_evento(
            cnpj=cnpj_partner, 
            chave_nfe=nfe_key )
    elif method == 'nao_realizar_operacao':
        result = p.nao_realizar_operacao_evento(
            cnpj=cnpj_partner, 
            chave_nfe=nfe_key )
    
    if result.resposta.status == 200: #Webservice ok
        if result.resposta.cStat.valor =='128': 
            inf_evento = result.resposta.retEvento[0].infEvento
            return { 'code': inf_evento.cStat.valor, 'message': inf_evento.xMotivo.valor,
                    'file_sent': result.envio.xml, 'file_returned': result.resposta.xml }
        else:
            return { 'code': result.resposta.cStat.valor, 'message': result.resposta.xMotivo.valor,
                    'file_sent': result.envio.xml, 'file_returned': result.resposta.xml }
    else:
        return { 'code': result.resposta.status, 'message': result.resposta.reason, 
                'file_sent': result.envio.xml, 'file_returned': None }

def download_nfe(company, list_nfe):
    p = __processo(company)
    cnpj_partner = re.sub('[^0-9]','', company.cnpj_cpf)   
    result = p.baixar_notas_destinadas(cnpj=cnpj_partner, lista_chaves=list_nfe)
    import_folder = company.nfe_import_folder
    
    if result.resposta.status == 200: #Webservice ok
        if result.resposta.cStat.valor == '139':
            nfe = result.resposta.retNFe[0]
            if nfe.cStat.valor == '140':             
                
                nome_arq = os.path.join(import_folder, 'download_nfe/')
                try:
                    os.makedirs(nome_arq)
                except:
                    pass
                nome_arq = nome_arq + nfe.chNFe.valor +  'download-nfe.xml'
                arq = open(nome_arq, 'w')
                arq.write(nfe.procNFe.valor.encode('utf-8'))
                arq.close()
                
                return { 'code': nfe.cStat.valor, 'message': nfe.xMotivo.valor ,
                       'file_sent': result.envio.xml, 'file_returned': nome_arq,
                       'nfe': nfe}
            else:                       
                return { 'code': nfe.cStat.valor, 'message': nfe.xMotivo.valor,
                   'file_sent': result.envio.xml, 'file_returned': result.resposta.xml }
            
        else:
            return { 'code': result.resposta.cStat.valor, 'message': result.resposta.xMotivo.valor,
                   'file_sent': result.envio.xml, 'file_returned': result.resposta.xml }
    else:
        return { 'code': result.resposta.status, 'message': result.resposta.reason, 
                'file_sent': result.envio.xml, 'file_returned': None }
 
    
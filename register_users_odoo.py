import xmlrpc.client
import pandas as pd


url = "http://127.0.0.1"
db = ""
username = 'admin'
password = 'admin'

common = xmlrpc.client.ServerProxy('{}/xmlrpc/2/common'.format(url))
output = common.version()

# Autenticação / Retonar ID do usuário
uid = common.authenticate(db, username, password, {})

models = xmlrpc.client.ServerProxy('{}/xmlrpc/2/object'.format(url))


dependentes = pd.read_excel(
    open('document.xlsx', 'rb'))


qnt = len(dependentes['e-mail'])
i = 0
dependentes_dict = {}

while i < qnt:

    dependentes_dict.update(
        {
            dependentes['e-mail'][i].lower(): [
                dependentes['estado_civil_funcionario'][i], # 0
                dependentes['condicao_saude_funcionario'][i],
                dependentes['condicao_saude2_funcionario'][i],
                dependentes['escolaridade_funcionario'][i],

                dependentes['tipo_dp1'][i],  #4
                dependentes['nome_dp1'][i],
                dependentes['email_dp1'][i],
                dependentes['estado_civil_dp1'][i],
                dependentes['escolaridade_dp1'][i],  
                dependentes['condicao_saude1_dp1'][i],
                dependentes['condicao_saude2_dp1'][i],
                
                dependentes['tipo_dp2'][i], # 11
                dependentes['nome_dp2'][i],
                dependentes['email_dp2'][i],
                dependentes['estado_civil_dp2'][i],
                dependentes['escolaridade_dp2'][i], 
                dependentes['condicao_saude1_dp2'][i],
                dependentes['condicao_saude2_dp2'][i],


                dependentes['tipo_dp3'][i], # 18
                dependentes['nome_dp3'][i],
                dependentes['email_dp3'][i],
                dependentes['estado_civil_dp3'][i],
                dependentes['escolaridade_dp3'][i],
                dependentes['condicao_saude1_dp3'][i],
                dependentes['condicao_saude2_dp3'][i],

                dependentes['tipo_dp4'][i], # 25
                dependentes['nome_dp4'][i],
                dependentes['email_dp4'][i],
                dependentes['estado_civil_dp4'][i], 
                dependentes['escolaridade_dp4'][i],
                dependentes['condicao_saude1_dp4'][i],
                dependentes['condicao_saude2_dp4'][i],

                dependentes['tipo_dp5'][i], # 32
                dependentes['nome_dp5'][i],
                dependentes['email_dp5'][i],
                dependentes['estado_civil_dp5'][i],
                dependentes['escolaridade_dp5'][i],
                dependentes['condicao_saude1_dp5'][i],
                dependentes['condicao_saude2_dp5'][i],
            ]
        }
    )

    i += 1


users = models.execute_kw(
    db, uid, password, 'res.partner', 'search_read',
    [], {'fields': ['id', 'email']})


for user in users:

    email = user['email'].lower()
    user_id = user['id']

    emails_cadastrados = []

    if email in dependentes_dict:
    
        dp_user = dependentes_dict[email]
        prox = 4

        while prox < 38:
            
            estado_civil_funcionario = dp_user[0]
            condicao_saude_funcionario = dp_user[1]
            condicao_saude2_funcionario = dp_user[2]
            escolaridade_funcionario = dp_user[3]


            tipo_dp = dp_user[prox]
            nome_dp = dp_user[prox + 1]
            email_dp = dp_user[prox + 2]
            estadocivil_dp = dp_user[prox + 3]
            escolaridade_dp = dp_user[prox + 4]
            condicao_saude1_dp = dp_user[prox + 5]
            condicao_saude2_dp = dp_user[prox + 6]

            if pd.isna(tipo_dp) == False:

                funcionario = models.execute_kw(
                    db, uid, password, 'res.partner', 'search_read',
                    [[['id', '=', user_id],]],
                    {'fields': ['id', 'email']})

                funcionario_id = funcionario[0]['id']
                funcionario_email = funcionario[0]['email']

                # VERIFICACAO FUNCIONARIOS

                if estado_civil_funcionario == "Casado(a)":
                    estado_civil_funcionario = "casado"
                elif estado_civil_funcionario == "Solteiro(a)":
                    estado_civil_funcionario = "solteiro"
                elif estado_civil_funcionario == "Em união estável":
                    estado_civil_funcionario = "uniaoestavel"
                elif estado_civil_funcionario == "Divorciado(a)":
                    estado_civil_funcionario = "divorciado"
                else:
                    estado_civil_funcionario = False

                if escolaridade_funcionario == "Fundamental - Incompleto":
                    escolaridade_funcionario = "fundamentalincompleto"
                elif escolaridade_funcionario == "Fundamental - Completo":
                    escolaridade_funcionario = "fundamentalcompleto"
                elif escolaridade_funcionario == "Médio - Completo":
                    escolaridade_funcionario = "mediocompleto"
                elif escolaridade_funcionario == "Médio - Incompleto":
                    escolaridade_funcionario = "medioincompleto"
                elif escolaridade_funcionario == "Superior - Completo":
                    escolaridade_funcionario = "superiorcompleto"
                elif escolaridade_funcionario == "Superior - Incompleto":
                    escolaridade_funcionario = "superiorincompleto"
                elif escolaridade_funcionario == "Mestrado - Completo":
                    escolaridade_funcionario = "mestradocompleto"
                elif escolaridade_funcionario == "Mestrado - Incompleto":
                    escolaridade_funcionario = "mestradoincompleto"
                elif escolaridade_funcionario == "Doutorado - Completo":
                    escolaridade_funcionario = "doutoradocompleto"
                elif escolaridade_funcionario == "Doutorado - Incompleto":
                    escolaridade_funcionario = "doutoradoincompleto"
                else:
                    escolaridade_funcionario = False
                
                # Condicao saude 1
                if condicao_saude_funcionario == "Diabetes":
                    # Tipo ??
                    pass
                if condicao_saude_funcionario == "Doenças Reumáticas ou Auto Imunes":
                    reumatismo_funcionario = True
                else:
                    reumatismo_funcionario = False

                if condicao_saude_funcionario == "Em tratamento para Câncer (quimioterapia)":
                    # não possui campo
                    pass
                if condicao_saude_funcionario == "HIpertensão":
                    hipertensao_funcionario = True
                else:
                    hipertensao_funcionario = False
                
                if condicao_saude_funcionario == "Transplantado de órgãos":
                    # nao possui campo
                    pass

                # Condicao saude 2
                if condicao_saude2_funcionario == "Crise de Asma Frequentes (Semanais)":
                    asma_funcionario = True
                else:
                    asma_funcionario = False

                if condicao_saude2_funcionario == "Epilepsia (Convulsões)":
                    # não possui campo
                    pass
            
                if condicao_saude2_funcionario == "Doença pulmonar obstrutiva crônica (enfisema)":
                    # não possui campo
                    pass

                if condicao_saude2_funcionario == "Outras":
                    # não possui campo
                    pass

                # Verificação de Dependentes
                if tipo_dp == "Cônjuje":
                    tipo_dp = "conjuge"
                elif tipo_dp == "Filho/Filha":
                    tipo_dp = "filho"
                else:
                    tipo_dp = "outros"

                if estadocivil_dp == "Casado(a)":
                    estadocivil_dp = "casado"
                elif estadocivil_dp == "Solteiro(a)":
                    estadocivil_dp = "solteiro"
                elif estadocivil_dp == "Em união estável":
                    estadocivil_dp = "uniaoestavel"
                elif estadocivil_dp == "Divorciado(a)":
                    estadocivil_dp = "divorciado"
                else:
                    estadocivil_dp = False

                if escolaridade_dp == "Fundamental - Incompleto":
                    escolaridade_dp = "fundamentalincompleto"
                elif escolaridade_dp == "Fundamental - Completo":
                    escolaridade_dp = "fundamentalcompleto"
                elif escolaridade_dp == "Médio - Completo":
                    escolaridade_dp = "mediocompleto"
                elif escolaridade_dp == "Médio - Incompleto":
                    escolaridade_dp = "medioincompleto"
                elif escolaridade_dp == "Superior - Completo":
                    escolaridade_dp = "superiorcompleto"
                elif escolaridade_dp == "Superior - Incompleto":
                    escolaridade_dp = "superiorincompleto"
                elif escolaridade_dp == "Mestrado - Completo":
                    escolaridade_dp = "mestradocompleto"
                elif escolaridade_dp == "Mestrado - Incompleto":
                    escolaridade_dp = "mestradoincompleto"
                elif escolaridade_dp == "Doutorado - Completo":
                    escolaridade_dp = "doutoradocompleto"
                elif escolaridade_dp == "Doutorado - Incompleto":
                    escolaridade_dp = "doutoradoincompleto"
                else:
                    escolaridade_dp = False
                
                # Condicao saude 1
                if condicao_saude1_dp == "Diabetes":
                    # Tipo ??
                    pass
                if condicao_saude1_dp == "Doenças Reumáticas ou Auto Imunes":
                    reumatismo = True
                else:
                    reumatismo = False

                if condicao_saude1_dp == "Em tratamento para Câncer (quimioterapia)":
                    # não possui campo
                    pass
                if condicao_saude1_dp == "HIpertensão":
                    hipertensao = True
                else:
                    hipertensao = False
                
                if condicao_saude1_dp == "Transplantado de órgãos":
                    # nao possui campo
                    pass

                # Condicao saude 2
                if condicao_saude2_dp == "Crise de Asma Frequentes (Semanais)":
                    asma = True
                else:
                    asma = False

                if condicao_saude2_dp == "Epilepsia (Convulsões)":
                    # não possui campo
                    pass
            
                if condicao_saude2_dp == "Doença pulmonar obstrutiva crônica (enfisema)":
                    # não possui campo
                    pass

                if condicao_saude2_dp == "Outras":
                    # não possui campo
                    pass

                paciente = models.execute_kw(
                    db, uid, password, 'gh.paciente', 'search_read',
                    [[['email', '=', funcionario_email], ['depende_de', '=', False]]],
                    {'fields': ['id']})

                paciente_id = paciente[0]['id']
                
                models.execute_kw(db, uid, password, 'gh.paciente', 'write', [[paciente_id],
                    {
                        'estadocivil': estado_civil_funcionario,
                        'reumatismo': reumatismo_funcionario,
                        'hipertensao': hipertensao_funcionario,
                        'asma': asma_funcionario,
                        'escolaridade': escolaridade_funcionario
                    }])
                
                # SE dependente não tem email ou o email já foi utilizado apenas cria outro paciente
                if pd.isna(email_dp) or email_dp == email or email_dp in emails_cadastrados:
                    
                    user_func = models.execute_kw(
                        db, uid, password, 'gh.paciente', 'search_read',
                        [[['email', '=', funcionario_email], ['depende_de', '=', False]]],
                        {'fields': ['id', 'user_id']})

                    user_func_id = user_func[0]['user_id']
                    user_func_id = user_func_id[0]

                    paciente = models.execute_kw(db, uid, password, 'gh.paciente', 'create', [{
                        'name': nome_dp,
                        'email': funcionario_email,
                        'user_id': user_func_id,
                        'estadocivil': estadocivil_dp,
                        'depende_de': paciente_id,
                        'escolaridade': escolaridade_dp,
                        'reumatismo': reumatismo,
                        'hipertensao': hipertensao,
                        'asma': asma,
                    }])

                # SE dependente tem email, cria outro usuario e paciente
                elif pd.isna(email_dp) == False and email_dp != email and email_dp != email:

                    new_user = models.execute_kw(db, uid, password, 'res.users', 'create', [{
                        'name': nome_dp,
                        'login': email_dp,
                        'email': email_dp
                    }])
                    
                    paciente = models.execute_kw(db, uid, password, 'gh.paciente', 'create', [{
                        'name': nome_dp,
                        'email': email_dp,
                        'user_id': new_user,
                        'estadocivil': estadocivil_dp,
                        'depende_de': paciente_id,
                        'escolaridade': escolaridade_dp,
                        'reumatismo': reumatismo,
                        'hipertensao': hipertensao,
                        'asma': asma,
                    }])

                    emails_cadastrados.append(email_dp)
                    
            prox = prox + 7

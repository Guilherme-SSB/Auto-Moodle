import datetime
import os
import time

import pandas as pd
from dotenv import load_dotenv

from help_functions import (clean_screen, find_element_by_xpath,
                            iniciate_chromedriver, mandar_email)

TODAY = datetime.datetime.now().date()
TODAY_FORMATADO = TODAY.strftime('%d/%m/%Y')

def clear_string(string: str) -> str:
    return string.strip().replace('\n', ' ').replace('\t', ' ').replace('\r', ' ')


def mes_nominal_para_numero(mes_nominal: str) -> str:
    meses = {
        'janeiro': '01',
        'fevereiro': '02',
        'março': '03',
        'abril': '04',
        'maio': '05',
        'junho': '06',
        'julho': '07',
        'agosto': '08',
        'setembro': '09',
        'outubro': '10',
        'novembro': '11',
        'dezembro': '12',
    }

    return meses[mes_nominal]


def formatar_data_entrega(data_entrega: str) -> str:
    """
    A data_entrega vem no formato: "quarta, 8 março, 23:55", "Hoje, 23:59", "Amanhã, 23:59".
    Formato desejado: YYYY-MM-DD HH:MM
    
    """
    data_entrega = data_entrega.split(', ')

    if len(data_entrega) == 2:
        dia = data_entrega[0]
        hora = data_entrega[1]

        if dia == 'Hoje':
            dia = TODAY

        elif dia == 'Amanhã':
            dia = str((datetime.datetime.now().date() + datetime.timedelta(days=1)))

        else:
            raise Exception(f'Erro ao formatar data_entrega - {data_entrega}')

    elif len(data_entrega) == 3:
        dia = str(data_entrega[1].split(' ')[0])
        dia = dia.zfill(2)

        mes = data_entrega[1].split(' ')[1]
        mes = mes_nominal_para_numero(mes)
        ano = TODAY.year

        dia = f'{ano}-{mes}-{dia}'
        hora = data_entrega[2]    

    else:
        raise Exception(f'Erro ao formatar data_entrega - {data_entrega}')

    return f'{dia} {hora}'


def estilizar_tabela_para_email(df: str) -> str:

    # Formatar tabela para o e-mail
    df = df.replace(
        '<table border="1" class="dataframe">', 
        '<table border="1" class="dataframe" style="width: 100%; border-collapse: collapse; border: 1px solid #ddd; font-size: 12px; font-family: Arial, Helvetica, sans-serif;">'
        )
    
    
    ## Centralizar o texto
    df = df.replace(
        '<td>',
        '<td style="text-align: center;">'
        )
    
    return df


def main():
    
    URL = 'https://imt.myopenlms.net/'
    
    driver = iniciate_chromedriver()
    driver.get(URL)
    time.sleep(5)

    btn_acessar = find_element_by_xpath(driver, '/html/body/header/div/div/a')
    btn_acessar.click()
    time.sleep(2)

    btn_office_365 = find_element_by_xpath(driver, '/html/body/div[3]/div/main/section/div/div[2]/div/div/div/div/a/img')
    btn_office_365.click()
    time.sleep(2)

    input_email = find_element_by_xpath(driver, '/html/body/div/form[1]/div/div/div[2]/div[1]/div/div/div/div/div[1]/div[3]/div/div/div/div[2]/div[2]/div/input[1]')
    input_email.click()
    input_email.send_keys(IMT_EMAIL)

    btn_next = find_element_by_xpath(driver, '/html/body/div/form[1]/div/div/div[2]/div[1]/div/div/div/div/div[1]/div[3]/div/div/div/div[4]/div/div/div/div/input')
    btn_next.click()
    time.sleep(2)

    input_password = find_element_by_xpath(driver, '/html/body/div/form[1]/div/div/div[2]/div[1]/div/div/div/div/div/div[3]/div/div[2]/div/div[3]/div/div[2]/input')
    input_password.click()
    input_password.send_keys(IMT_PASSWORD)

    btn_entrar = find_element_by_xpath(driver, '/html/body/div/form[1]/div/div/div[2]/div[1]/div/div/div/div/div/div[3]/div/div[2]/div/div[4]/div[2]/div/div/div/div/input')
    btn_entrar.click()
    time.sleep(2)

    btn_nao_continuar_conectado = find_element_by_xpath(driver, '/html/body/div/form/div/div/div[2]/div[1]/div/div/div/div/div/div[3]/div/div[2]/div/div[3]/div[2]/div/div/div[1]/input')
    btn_nao_continuar_conectado.click()
    time.sleep(5)

    btn_meus_cursos = find_element_by_xpath(driver, '/html/body/header/div/div/a/span')
    btn_meus_cursos.click()
    time.sleep(1)

    # Carrega todas as tarefas
    while True:
        try:
            btn_ver_mais = find_element_by_xpath(driver, '/html/body/nav/div/div[1]/div/section[1]/snap-feed/a/small')

            if btn_ver_mais.text == 'Ver mais':
                btn_ver_mais.click()
                time.sleep(0.5)
            else:
                break
        except:
            break

    
    # Coletar dados das Tarefas e salva como dataframe
    i = 1
    df_final = pd.DataFrame()
    while True:
        try:
            xpath_tarefa = f'/html/body/nav/div/div[1]/div/section[1]/snap-feed/div/div[{i}]'
            nome = clear_string(find_element_by_xpath(driver, xpath_tarefa+'/div[2]/a/h3').text)
            materia = clear_string(find_element_by_xpath(driver, xpath_tarefa+'/div[2]/a/h3/small').text)
            link = clear_string(find_element_by_xpath(driver, xpath_tarefa+'/div[2]/a').get_attribute('href'))
            data_entrega = formatar_data_entrega(clear_string(find_element_by_xpath(driver, xpath_tarefa+'/div[2]/span/time').text))
            status = clear_string(find_element_by_xpath(driver, xpath_tarefa+'/div[2]/span/div/a').text)

            # Se a materia for ECM307-Sistemas e Sinais, pula
            if (materia == 'ECM307-Sistemas e Sinais') or (materia == 'ECM401-Banco de Dados') or (materia == 'ECM971-Devops: Metodologia de Desenvolvimento de Software'):
                i += 1
                continue
            
            # Se tiver "Professor Ricardo Fernandes" na materia "EFH117-Direito Empresarial", pula
            if materia == 'EFH117-Direito Empresarial' and 'Professor Ricardo Fernandes' in nome:
                i += 1
                continue

            print(f'{nome} - {materia} - {data_entrega} - {status}')

            df_aux = pd.DataFrame({
                'NOME': [nome],
                'MATÉRIA': [materia],
                'DATA ENTREGA': [data_entrega],
                'STATUS': [status],
                'LINK': [link],
            })

            df_final = pd.concat([df_final, df_aux], ignore_index=True)

            i += 1
        except:
            break
    
    df_final.to_csv(f'{OUTPUTS_DIR}/tarefas_{TODAY}.csv', index=False)

    driver.close()

    # Gerar relatório
    ## Ler o arquivo com as tarefas do dia
    df_tarefas_dia = pd.read_csv(f'{OUTPUTS_DIR}/tarefas_{TODAY}.csv')

    ## Filtrar tabelas por status
    df_tarefas_dia_email = df_tarefas_dia[df_tarefas_dia['STATUS'].isin(['Não submetido', 'Sem tentativa'])]

    ## Criar e-mail
    df_tarefas_dia_email.sort_values(by=['DATA ENTREGA'], inplace=True)
    df_tarefas_dia_email['DATA ENTREGA'] = df_tarefas_dia_email['DATA ENTREGA'].astype('datetime64[ns]')
    df_tarefas_dia_email['DATA ENTREGA'] = df_tarefas_dia_email['DATA ENTREGA'].dt.strftime('%d/%m/%Y %H:%M')

    df_tarefas_dia_email = df_tarefas_dia_email.to_html(index=False)
    df_tarefas_dia_email = estilizar_tabela_para_email(df_tarefas_dia_email)

    mensagem = f'''
                <h3>Olá, Gui!</h3>
                <p>Estas são próximas tarefas a serem entregues</p>
                <p>{df_tarefas_dia_email}</p>
            '''

    # Enviar email
    mandar_email(
        to='gui.samuel10@gmail.com',
        subject=f'[IMT Tarefas] Tarefas Moodle - {TODAY_FORMATADO}',
        message=mensagem
    )

    ## Atualizar Base_Tarefas_IMT.csv com as tarefas do dia
    base_tarefas_imt = pd.read_csv(f'{OUTPUTS_DIR}/Base_Tarefas_IMT.csv')

    base_tarefas_imt = pd.concat([base_tarefas_imt, df_tarefas_dia], ignore_index=True)
    base_tarefas_imt.drop_duplicates(inplace=True)

    base_tarefas_imt.to_csv(f'{OUTPUTS_DIR}/Base_Tarefas_IMT.csv', index=False)



if __name__ == '__main__':
    clean_screen()

    OUTPUTS_DIR = 'D:/GitHub/Auto Moodle/outputs'

    load_dotenv()
    IMT_EMAIL = str(os.getenv('IMT_EMAIL'))
    IMT_PASSWORD = str(os.getenv('IMT_PASSWORD'))
    
    main()



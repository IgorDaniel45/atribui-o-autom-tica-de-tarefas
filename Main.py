import Graph_API
import atualizar_planilha
import verificar_maquinas_preparadas
import verificar_resposta_email
import script5

def main():
    # Atualizar a planilha original com dados adicionais
    atualizar_planilha.atualizar_planilha_com_dados_adicionais()
    
    # Verificar se as máquinas foram preparadas
    verificar_maquinas_preparadas.verificar_maquinas()
    
    # Monitorar respostas de e-mail
    verificar_resposta_email.monitorar_respostas()
    
    # Iniciar o agendamento de tarefas
    def run_schedule():
        while True:
            script5.schedule.run_pending()
            time.sleep(1)
    
    schedule_thread = threading.Thread(target=run_schedule)
    schedule_thread.start()
    
if __name__ == "__main__":
    main()

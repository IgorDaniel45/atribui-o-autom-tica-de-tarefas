import schedule
import time
from verificar_maquinas_preparadas import verificar_maquinas
from verificar_resposta_email import monitorar_respostas

# Agendar a tarefa para todos os dias úteis às 9 horas da manhã
schedule.every().monday.at("09:00").do(verificar_maquinas)
schedule.every().tuesday.at("09:00").do(verificar_maquinas)
schedule.every().wednesday.at("09:00").do(verificar_maquinas)
schedule.every().thursday.at("09:00").do(verificar_maquinas)
schedule.every().friday.at("09:00").do(verificar_maquinas)

# Agendar a tarefa para monitorar respostas de e-mail a cada 5 minutos
schedule.every(5).minutes.do(monitorar_respostas)

# Loop para manter o agendamento ativo
while True:
    schedule.run_pending()
    time.sleep(1)
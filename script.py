import time

from func import *

# Início
iniciar_processo()
conectar_servidores()

# Agenda a verificação inicial
agendar_verificacao()

# # Agendar a verificação inicial e executar
verificar_eProcessar()
schedule.every(10).seconds.do(verificar_eProcessar)  # Agora chama a função correta

while True:
    print("Verificando tarefas agendadas...")
    schedule.run_pending()
    time.sleep(3) 

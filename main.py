# import app , nsp , cimatec_experience, cimatec_park, CI, agendaN
import nsi
from time import sleep 
email = 'mateus.freire@fbest.org.br'
password = "Meufelho12*"


while nsi.run_call(email=email, password=password) == 0:
    sleep(2)
    nsi.run_call(email=email, password=password)



# print(email)
# print(password)
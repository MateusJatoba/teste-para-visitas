import sqlite3
import os

def get_conexao():
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    db_path = os.path.join(BASE_DIR, 'snippets', 'banco_visitas.db')

    os.makedirs(os.path.dirname(db_path), exist_ok=True)
    return sqlite3.connect(db_path)

# # Criação da tabela TB_visita
# cursor.execute("""
# CREATE TABLE IF NOT EXISTS TB_visita (
#     id INTEGER PRIMARY KEY AUTOINCREMENT,
#     nome_instituicao VARCHAR(30),
#     data_visita DATE,
#     endereco VARCHAR(30),
#     veiculo VARCHAR(20),
#     almoco BOOLEAN,
#     endereco_almoco VARCHAR(100)
# );
# """)

# # Criação da tabela TB_participante_Externo
# cursor.execute("""
# CREATE TABLE IF NOT EXISTS TB_participante_Externo (
#     id INTEGER PRIMARY KEY AUTOINCREMENT,
#     id_visita INTEGER,
#     nome VARCHAR(100),
#     cargo VARCHAR(100),
#     FOREIGN KEY (id_visita) REFERENCES TB_visita(id)
# );
# """)

# # Criação da tabela TB_participante_Interno
# cursor.execute("""
# CREATE TABLE IF NOT EXISTS TB_participante_Interno (
#     id INTEGER PRIMARY KEY AUTOINCREMENT,
#     matricula INTEGER,
#     nome VARCHAR(100),
#     cargo VARCHAR(100),
#     email VARCHAR(100)
# );
# """)

# # Criação da tabela visita_participante_Interno (ligação N:N entre visitas e participantes internos)
# cursor.execute("""
# CREATE TABLE IF NOT EXISTS visita_participante_Interno (
#     id INTEGER PRIMARY KEY AUTOINCREMENT,
#     id_visita INTEGER,
#     id_participante INTEGER,
#     FOREIGN KEY (id_visita) REFERENCES TB_visita(id),
#     FOREIGN KEY (id_participante) REFERENCES TB_participante_Interno(id)
# );
# """)

# cursor.execute(
#     # "DELETE FROM TB_visita WHERE id = 5;" \
#     "DELETE FROM visita_participante_Interno WHERE id_visita = 5;"
# )

# cursor.execute(
#     "ALTER TABLE TB_visita " \
#     "add descricao_roteiro TEXT;"
# )

# cursor.execute(
#     "INSERT INTO TB_participante_Interno (matricula , nome , cargo , email)" \
#     "VALUES (4558, 'Luiz Fernando Taboada Gomes Amaral', 'Gerente Executivo - Negócios' , 'luiz.amaral@fieb.org.br');"
# )


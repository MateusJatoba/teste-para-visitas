from flask import Flask, render_template, request, redirect, url_for, session
import sqlite3
import threading
import os
import agendaN
from time import sleep


app = Flask(__name__)
app.secret_key = 'your-secret-key'  # Required for session

def shutdown_server():
    func = request.environ.get('werkzeug.server.shutdown')
    if func is None:
        raise RuntimeError('Não está rodando com o Werkzeug Server')
    func()


@app.route("/", methods=["GET"])
def index():
    return render_template("formulario.html")

@app.route("/salvar", methods=["POST"])
def salvar():
    # Store form data in session
    session['form_data'] = {
        'nome_instituicao': request.form.get("nome_instituicao"),
        'data_visita': request.form.get("data_visita"),
        'endereco': request.form.get("endereco"),
        'veiculo': request.form.get("veiculo"),
        'almoco': request.form.get("almoco"),
        'endereco_almoco': request.form.get("endereco_almoco"),
        'descricao_roteiro': request.form.get("roteiro")
    }
    
    # Redirect to participant selection page
    return redirect(url_for('selecao_participantes'))



@app.route("/selecao_participantes")
def selecao_participantes():
    conn = sqlite3.connect('banco_visitas.db')
    cursor = conn.cursor()

    cursor.execute("SELECT id, nome FROM TB_participante_Interno")
    colaboradores = cursor.fetchall()
    
    conn.close()
    return render_template("selecao_participantes.html", colaboradores=colaboradores)




@app.route("/finalizar", methods=["POST"])
def finalizar():
    # Get form data from session
    form_data = session.get('form_data', {})
    
    # Get selected participants
    colaboradores = request.form.getlist('colaboradores')
    
    conn = sqlite3.connect('banco_visitas.db')
    cursor = conn.cursor()

    # Insert visit data
    cursor.execute(
        "INSERT INTO TB_visita (nome_instituicao, data_visita, endereco, veiculo, almoco, endereco_almoco, descricao_roteiro) "
        "VALUES (?, ?, ?, ?, ?, ?, ?)",
        (form_data['nome_instituicao'], form_data['data_visita'], form_data['endereco'],
         form_data['veiculo'], form_data['almoco'], form_data['endereco_almoco'], form_data['descricao_roteiro'])
    )
    
    id_visita = cursor.lastrowid

    for colaborador_id in colaboradores:
        cursor.execute(
            "INSERT INTO visita_participante_interno (id_visita, id_participante) VALUES (?, ?)",
            (id_visita, colaborador_id)
        )
    
    conn.commit()
    conn.close()

    # Clear session data
    session.pop('form_data', None)
    
    return render_template("selecao_externos.html")



@app.route("/selecao_externos", methods=["GET"])
def cadastrar_visitante():
    return render_template("selecao_externos.html")  # seu novo HTML

@app.route("/salvar_visitante", methods=["POST"])
def salvar_visitante(): 
    nomes = request.form.getlist("nome_visitante[]")
    cargos = request.form.getlist("cargo_visitante[]")

    conn = sqlite3.connect('banco_visitas.db')
    cursor = conn.cursor()
    last_id = cursor.execute(
        "SELECT MAX(id) FROM TB_visita"
    )

    last_id = last_id.fetchone()[0]
    for nome, cargo in zip(nomes, cargos):
        cursor.execute(
            "INSERT INTO TB_participante_Externo (nome, cargo, id_visita) VALUES (?, ?, ?)",
            (nome, cargo, last_id)
        )

    conn.commit()
    conn.close()

    agendaN.a.manipula_documento()

    

    return "Visitante cadastrado com sucesso!"



if __name__ == "__main__":
    app.run(debug=True)
    
import requests
def conversar(nome_instituicao):
    response = requests.post(
        "http://localhost:11434/api/generate",
        json={
            "model": "mistral",
            "prompt": f"Faça um parágrafo de tamanho mediano acerca da instituição: {nome_instituicao}",
            "stream": False
        }
    )
    return response.json()['response']




from gpt4all import GPT4All

# Caminho para o modelo baixado (exemplo com Mistral 7B Instruct)
modelo = GPT4All("mistral-7b-instruct-v0.1.Q4_0.gguf", model_path="./models")

# Inicia o modelo (a primeira vez pode demorar um pouco)
with modelo.chat_session():
    resposta = modelo.generate("Me faça um resumo da empresa Siemens Energy em linguagem técnica e objetiva.")
    print(resposta)

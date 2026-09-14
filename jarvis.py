"""Assistente virtual conversacional inspirado no J.A.R.V.I.S.

Conduz uma entrevista por etapas para conhecer o usuário (nome, idade,
profissão, família e saúde) e guarda o perfil resultante em disco.
"""
from datetime import datetime
import json
import os
import re
import unicodedata

from flask import Blueprint, jsonify, render_template, request, session

jarvis_bp = Blueprint('jarvis', __name__, url_prefix='/jarvis')

PASTA_PERFIS = 'perfis'

PERGUNTAS = [
    {
        'chave': 'nome',
        'pergunta': 'Antes de começarmos, como devo chamá-lo?',
    },
    {
        'chave': 'idade',
        'pergunta': 'Qual é a sua idade?',
    },
    {
        'chave': 'profissao',
        'pergunta': 'E a que se dedica? Qual é a sua profissão?',
    },
    {
        'chave': 'familia',
        'pergunta': (
            'Fale-me sobre sua família: quem são as pessoas mais próximas '
            'de você e qual o grau de parentesco de cada uma?'
        ),
    },
    {
        'chave': 'saude',
        'pergunta': (
            'Por último, sobre sua saúde: possui alguma condição, alergia '
            'ou medicação que eu deva sempre levar em conta?'
        ),
    },
]


def saudacao_por_horario():
    hora = datetime.now().hour
    if hora < 12:
        return 'Bom dia'
    if hora < 18:
        return 'Boa tarde'
    return 'Boa noite'


def _estado_inicial():
    return {'indice': 0, 'respostas': {}}


def _resumo(respostas):
    linhas = [
        f"Nome: {respostas.get('nome', '—')}",
        f"Idade: {respostas.get('idade', '—')}",
        f"Profissão: {respostas.get('profissao', '—')}",
        f"Família: {respostas.get('familia', '—')}",
        f"Saúde: {respostas.get('saude', '—')}",
    ]
    return '\n'.join(linhas)


def _salvar_perfil(respostas):
    os.makedirs(PASTA_PERFIS, exist_ok=True)
    nome = respostas.get('nome', 'usuario')
    nome_ascii = unicodedata.normalize('NFKD', nome).encode('ascii', 'ignore').decode('ascii')
    slug = re.sub(r'[^a-zA-Z0-9_-]+', '_', nome_ascii.strip().lower()).strip('_') or 'usuario'
    carimbo = datetime.now().strftime('%Y%m%d_%H%M%S')
    caminho = os.path.join(PASTA_PERFIS, f'{slug}_{carimbo}.json')
    dados = dict(respostas)
    dados['criado_em'] = datetime.now().isoformat()
    with open(caminho, 'w', encoding='utf-8') as arquivo:
        json.dump(dados, arquivo, ensure_ascii=False, indent=2)
    return caminho


@jarvis_bp.route('')
def index():
    session['jarvis'] = _estado_inicial()
    primeira_pergunta = PERGUNTAS[0]['pergunta']
    mensagem_inicial = (
        f"{saudacao_por_horario()}. Sou o seu assistente virtual e gostaria "
        f"de conhecê-lo melhor antes de começarmos a trabalhar juntos. "
        f"{primeira_pergunta}"
    )
    return render_template('jarvis.html', mensagem_inicial=mensagem_inicial)


@jarvis_bp.route('/responder', methods=['POST'])
def responder():
    dados = request.get_json(silent=True) or {}
    mensagem = (dados.get('mensagem') or '').strip()

    estado = session.get('jarvis') or _estado_inicial()
    indice = estado['indice']
    respostas = estado['respostas']

    if not mensagem:
        pergunta_atual = PERGUNTAS[indice]['pergunta'] if indice < len(PERGUNTAS) else ''
        return jsonify({
            'resposta': f'Não recebi nenhuma informação. {pergunta_atual}',
            'concluido': False,
        })

    if indice < len(PERGUNTAS):
        chave = PERGUNTAS[indice]['chave']
        respostas[chave] = mensagem
        indice += 1
        estado['indice'] = indice
        estado['respostas'] = respostas
        session['jarvis'] = estado

    if indice < len(PERGUNTAS):
        proxima = PERGUNTAS[indice]['pergunta']
        return jsonify({'resposta': proxima, 'concluido': False})

    _salvar_perfil(respostas)
    nome = respostas.get('nome', 'senhor(a)')
    resposta_final = (
        f'Perfeito, {nome}. Registrei todas as informações:\n\n'
        f'{_resumo(respostas)}\n\n'
        'A partir de agora vou levar tudo isso em conta sempre que conversarmos. '
        'Como posso ajudá-lo?'
    )
    return jsonify({'resposta': resposta_final, 'concluido': True})


@jarvis_bp.route('/reiniciar', methods=['POST'])
def reiniciar():
    session['jarvis'] = _estado_inicial()
    primeira_pergunta = PERGUNTAS[0]['pergunta']
    return jsonify({
        'resposta': (
            f'{saudacao_por_horario()}. Vamos começar novamente. '
            f'{primeira_pergunta}'
        ),
        'concluido': False,
    })

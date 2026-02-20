# DP Inteligente — Sistema de Automação de Atendimento (WhatsApp + Painel Web)

Projeto autoral de automação de atendimento para Departamento Pessoal, integrado ao WhatsApp via API (Z-API), com backend em Google Apps Script e painel web para gestão em tempo real.

## 📌 Objetivo

Automatizar o fluxo de atendimento interno do Departamento Pessoal (Admissão, Rescisão, Folha e Ponto), estruturando dados de interação, organizando logs e permitindo rastreabilidade das demandas.

O sistema simula um ambiente corporativo de atendimento automatizado orientado a eficiência operacional e organização de dados.

---

## ⚙️ Arquitetura do Sistema

**Backend**
- Google Apps Script
- Webhook para recebimento de mensagens
- Processamento de fluxo por status
- Integração com API do WhatsApp (Z-API)
- Integração com IA (Groq API)

**Banco de Dados**
- Google Sheets estruturado:
  - ATENDIMENTOS
  - LOG_CONVERSAS
  - FAQ
  - SETORES
  - CONFIG

**Frontend**
- Painel Web em HTML, CSS e JavaScript
- Lista de atendimentos com filtros
- Indicador de mensagens não lidas
- Chat em tempo real
- Controle manual de respostas e encerramento

---

## 🔄 Fluxo de Atendimento

1. Recebimento da mensagem via Webhook
2. Deduplicação para evitar processamento duplicado
3. Identificação ou criação do atendimento
4. Triagem automática via FAQ
5. Direcionamento por setor
6. Registro de logs estruturados
7. Atendimento humano via painel

---

## 🧠 Funcionalidades Implementadas

- Automação do fluxo por status:
  - NOVO
  - AGUARDANDO_DUVIDA
  - AGUARDANDO_SETOR
  - EM_ATENDIMENTO
  - ENCERRADO

- Sistema de FAQ automatizado
- Registro completo de logs
- Deduplicação com CacheService
- Painel web com atualização periódica
- Integração com IA para respostas estruturadas

---

## 🛠️ Tecnologias Utilizadas

- Google Apps Script
- JavaScript
- HTML
- CSS
- API REST
- Z-API
- Google Sheets
- Integração com IA (Groq)

---

## 🎯 Competências Demonstradas

- Automação de processos
- Estruturação e organização de dados operacionais
- Integração com APIs
- Desenvolvimento de backend
- Lógica de sistemas
- Painéis operacionais em tempo real

---

## 📷 Demonstração

Painel de atendimento em tempo real com controle de status e histórico de mensagens.

(Adicionar print do painel na pasta /docs ou diretamente neste README)

---

## 📌 Observação

Projeto desenvolvido com foco em aplicação prática de automação e organização de dados para apoio à tomada de decisão e eficiência operacional.

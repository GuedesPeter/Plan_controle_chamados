# 📊 Controle de Chamados - Planilha Automatizada

Este projeto gera automaticamente uma **planilha Excel (.xlsx)** para **controle de chamados de suporte**, organizada por funcionário e com um **Dashboard visual em gráficos de pizza**, utilizando **Python + OpenPyXL**.

A planilha foi pensada para ter um **layout moderno**, uso de **cores padronizadas**, **validações de dados**, **formatação condicional** e **indicadores visuais claros** para facilitar a gestão dos chamados.

---

## 🚀 Funcionalidades

### 📁 Abas por Funcionário
São criadas abas individuais para cada colaborador:
- **Douglas**
- **Gerson**
- **Paulo**

Cada aba contém os seguintes campos:
- **Número do Chamado**
- **Status** (dropdown)
- **Descrição**
- **Início EM**
- **Finalizar?** (SIM / NÃO)

---

### 🎯 Status dos Chamados (Dropdown com cores)
O campo **Status** possui validação de dados com as opções:

| Status           | Cor |
|------------------|-----|
| Novo             | Verde |
| Em Atendimento   | Verde Claro |
| Pendente         | Laranja |
| Solucionado      | Azul |
| Finalizado       | Cinza Escuro |

📌 Status com cores mais escuras possuem **texto branco e em negrito**, garantindo melhor leitura.

---

### ✅ Destaque de Chamados Finalizados
Quando o campo **Finalizar?** for marcado como **SIM**:
- Toda a linha do chamado é automaticamente destacada em **Azul Claro**

Isso facilita a identificação visual de chamados encerrados.

---

## 📊 Aba Resumo (Dashboard)

A aba **Resumo** funciona como um **Dashboard Gerencial**, contendo:

### 📌 Tabela Consolidada
- Quantidade de chamados por **Status**
- Visão **Geral**
- Visão individual por funcionário

---

### 🥧 Gráficos de Pizza
São gerados automaticamente gráficos de pizza respeitando as cores dos status:

- **Chamados por Status – Geral**
- **Chamados por Status – Douglas**
- **Chamados por Status – Gerson**
- **Chamados por Status – Paulo**

Esses gráficos permitem uma análise rápida da distribuição dos chamados.

---

## 🛠️ Tecnologias Utilizadas

- **Python 3**
- **OpenPyXL**
  - Criação e manipulação de planilhas Excel
  - Validação de dados
  - Formatação condicional
  - Gráficos (PieChart)

---

## ▶️ Como Executar o Projeto

### 1️⃣ Clone o repositório
```bash
git clone https://github.com/GuedesPeter/Plan_controle_chamados.git
cd Plan_controle_chamados

"""
Example: Using fill_document_template to populate Word templates with dynamic data

This example demonstrates how to use the fill_document_template function
to fill Word templates using Jinja2 syntax with docxtpl.
"""

import json
import asyncio
from word_document_server.tools.document_tools import fill_document_template


async def example_simple_template():
    """Example 1: Simple variable replacement"""
    
    # Data for simple variables
    data = {
        "assunto": "Relatório Mensal de Vendas",
        "codigo": "REL-2024-001",
        "data": "19/11/2024",
        "autor": "João Silva"
    }
    
    # Convert to JSON string
    data_json = json.dumps(data, ensure_ascii=False)
    
    # Fill the template
    result = await fill_document_template(
        template_path="template_simples.docx",
        output_path="output_simples.docx",
        data_json=data_json
    )
    
    print(result)


async def example_table_loop():
    """Example 2: Template with table loops"""
    
    # Data with list for table rows
    data = {
        "titulo": "Lista de Produtos",
        "data": "19/11/2024",
        "produtos": [
            {"nome": "Produto A", "quantidade": 10, "preco": 50.00},
            {"nome": "Produto B", "quantidade": 5, "preco": 120.00},
            {"nome": "Produto C", "quantidade": 15, "preco": 30.00}
        ],
        "total": 1550.00
    }
    
    data_json = json.dumps(data, ensure_ascii=False)
    
    result = await fill_document_template(
        template_path="template_tabela.docx",
        output_path="output_tabela.docx",
        data_json=data_json
    )
    
    print(result)


async def example_complex_template():
    """Example 3: Complex template with nested data and conditionals"""
    
    data = {
        "empresa": {
            "nome": "Empresa XYZ Ltda",
            "cnpj": "12.345.678/0001-90",
            "endereco": "Rua Exemplo, 123"
        },
        "relatorio": {
            "titulo": "Relatório Trimestral",
            "periodo": "Q4 2024",
            "aprovado": True
        },
        "secoes": [
            {
                "titulo": "Vendas",
                "descricao": "Análise de vendas do período",
                "items": [
                    {"descricao": "Vendas Online", "valor": 150000},
                    {"descricao": "Vendas Físicas", "valor": 80000}
                ]
            },
            {
                "titulo": "Despesas",
                "descricao": "Despesas operacionais",
                "items": [
                    {"descricao": "Salários", "valor": 50000},
                    {"descricao": "Aluguel", "valor": 10000}
                ]
            }
        ],
        "observacoes": "Documento gerado automaticamente"
    }
    
    data_json = json.dumps(data, ensure_ascii=False)
    
    result = await fill_document_template(
        template_path="template_complexo.docx",
        output_path="output_complexo.docx",
        data_json=data_json
    )
    
    print(result)


async def example_header_footer():
    """Example 4: Template with header and footer variables"""
    
    data = {
        "header_titulo": "CONFIDENCIAL",
        "header_codigo": "DOC-2024-001",
        "corpo_texto": "Este é o conteúdo principal do documento.",
        "footer_empresa": "Empresa XYZ Ltda",
        "footer_pagina": "Página {{page}} de {{total_pages}}"
    }
    
    data_json = json.dumps(data, ensure_ascii=False)
    
    result = await fill_document_template(
        template_path="template_header_footer.docx",
        output_path="output_header_footer.docx",
        data_json=data_json
    )
    
    print(result)


# Template syntax examples for reference:
"""
TEMPLATE SYNTAX EXAMPLES (to use in your .docx templates):

1. Simple variable:
   {{assunto}}
   {{codigo}}

2. Nested object:
   {{empresa.nome}}
   {{relatorio.titulo}}

3. Table loop:
   {% for produto in produtos %}
   {{produto.nome}}  {{produto.quantidade}}  {{produto.preco}}
   {% endfor %}

4. Conditional:
   {% if relatorio.aprovado %}
   Status: Aprovado
   {% else %}
   Status: Pendente
   {% endif %}

5. Nested loops:
   {% for secao in secoes %}
   {{secao.titulo}}
   {% for item in secao.items %}
   - {{item.descricao}}: {{item.valor}}
   {% endfor %}
   {% endfor %}
"""


if __name__ == "__main__":
    # Run examples
    print("Example 1: Simple template")
    asyncio.run(example_simple_template())
    
    print("\nExample 2: Table loop")
    asyncio.run(example_table_loop())
    
    print("\nExample 3: Complex template")
    asyncio.run(example_complex_template())
    
    print("\nExample 4: Header/Footer")
    asyncio.run(example_header_footer())

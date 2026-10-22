"""
Generator for: 10 integracao frete.pptx
"""

import os
from PIL import Image, ImageDraw, ImageFont
import pptx
from pptx import Presentation
from pptx.util import Inches, Pt

from slide_builder import (
    TEMP_DIR, FONT_HERO, FONT_TITLE, FONT_SUBTITLE, FONT_HEADING,
    FONT_BODY, FONT_BODY_BOLD, FONT_SMALL, FONT_SMALL_BOLD, FONT_CODE,
    C_DARK_BG, C_CARD_BG, C_CARD_BORDER, C_TEXT_LIGHT, C_TEXT_MUTED, C_WHITE,
    create_base_canvas, draw_header, draw_browser_chrome,
    draw_wp_sidebar, draw_arrow_pointer, draw_check_icon, draw_cross_icon,
    draw_bullet_dot, draw_star_badge
)

PRIMARY_TEAL = (14, 116, 144)      # Cyan 700
ACCENT_AMBER = (245, 158, 11)      # Amber 500
CORREIOS_YELLOW = (255, 204, 0)
CORREIOS_BLUE = (0, 75, 145)

def create_slide_1():
    img, draw = create_base_canvas()
    draw.rounded_rectangle([120, 160, 520, 215], radius=12, fill=PRIMARY_TEAL)
    draw.text((140, 172), "MODULO 10 • LOGISTICA", font=FONT_HEADING, fill=C_WHITE)
    
    draw.text((120, 250), "Integracao de Frete", font=FONT_HERO, fill=C_WHITE)
    draw.text((120, 335), "& Logistica no WooCommerce", font=FONT_HERO, fill=ACCENT_AMBER)
    
    draw.text((120, 440), "Configurando Envios Automaticos, Correios, Melhor Envio, Zonas de Entrega e Cubagem", font=FONT_SUBTITLE, fill=C_TEXT_MUTED)
    
    features = [
        ("Agregadores de Frete", "Melhor Envio e Frenet oferecem descontos de ate 80% nos Correios e Jadlog sem exigir contrato nem mensalidade."),
        ("Calculo Automatico em Tempo Real", "O cliente digita o CEP e a loja consulta as tabelas das transportadoras instantaneamente."),
        ("Cubagem Inteligente", "Como peso, largura, altura e comprimento evitam prejuizos no frete cobrado do comprador.")
    ]
    for i, (title, desc) in enumerate(features):
        x = 120 + i * 570
        y = 540
        draw.rounded_rectangle([x, y, x + 530, y + 260], radius=16, fill=C_CARD_BG, outline=C_CARD_BORDER, width=2)
        draw.ellipse([x + 30, y + 30, x + 80, y + 80], fill=PRIMARY_TEAL)
        draw.text((x + 46, y + 38), str(i+1), font=FONT_HEADING, fill=C_WHITE)
        draw.text((x + 95, y + 42), title, font=FONT_HEADING, fill=C_WHITE)
        
        words = desc.split()
        line = ""
        ty = y + 105
        for w in words:
            if len(line + " " + w) > 34:
                draw.text((x + 30, ty), line, font=FONT_BODY, fill=C_TEXT_MUTED)
                ty += 28
                line = w
            else:
                line = (line + " " + w).strip()
        if line:
            draw.text((x + 30, ty), line, font=FONT_BODY, fill=C_TEXT_MUTED)

    draw.rounded_rectangle([120, 850, 1800, 960], radius=12, fill=(22, 78, 99))
    draw.text((160, 880), "OBJETIVO DA AULA:", font=FONT_BODY_BOLD, fill=ACCENT_AMBER)
    draw.text((360, 880), "Dominar a configuracao de fretes no WooCommerce para que a loja calcule taxas exatas e gere etiquetas de despacho.", font=FONT_BODY, fill=C_WHITE)

    path = os.path.join(TEMP_DIR, "mod10_slide1.png")
    img.save(path, quality=95)
    return path

def create_slide_2():
    img, draw = create_base_canvas()
    draw_header(draw, "Panorama Nacional", "O Papel Estrategico do Frete no E-Commerce", "Mais de 65% dos abandonos de carrinho no Brasil sao causados por frete caro ou prazo excessivo", PRIMARY_TEAL)
    
    # 3 Stat Cards
    stats = [
        ("66%", "Maior Causador de Abandono", "Consumidores desistem da compra se o frete for considerado desproporcional ao valor do produto.", (239, 68, 68)),
        ("78%", "Decisao por Frete Gratis", "Clientes compram mais itens para atingir o valor minimo estipulado para ganhar envio gratuito.", (16, 185, 129)),
        ("2 a 3", "Opcoes Ideais de Envio", "O consumidor busca escolher entre a opcao mais economica (PAC/Jadlog) e a mais rapida (SEDEX).", PRIMARY_TEAL)
    ]
    for i, (val, title, desc, col) in enumerate(stats):
        x = 120 + i * 570
        y = 260
        draw.rounded_rectangle([x, y, x + 530, y + 240], radius=16, fill=C_CARD_BG, outline=col, width=2)
        draw.text((x + 35, y + 25), val, font=FONT_HERO, fill=col)
        draw.text((x + 35, y + 105), title, font=FONT_HEADING, fill=C_WHITE)
        draw.text((x + 35, y + 155), desc, font=FONT_BODY, fill=C_TEXT_MUTED)

    # Comparison Table
    draw.rounded_rectangle([120, 540, 1800, 980], radius=16, fill=C_CARD_BG, outline=C_CARD_BORDER, width=2)
    draw.text((160, 570), "Modalidades Indispensaveis em Toda Loja Virtual", font=FONT_HEADING, fill=C_WHITE)
    
    rows = [
        ("Frete Economico (PAC / Jadlog Package)", "Foco em menor custo. Prazo tipico de 4 a 10 dias uteis. Ideal para clientes sensiveis a preco.", "Indispensavel"),
        ("Frete Expresso (SEDEX / Jadlog .Com / Aereo)", "Foco em urgencia. Prazo de 1 a 3 dias uteis. Ideal para presentes, reposicoes e capitais.", "Indispensavel"),
        ("Frete Gratis Condicional", "Regra com valor minimo no carrinho (ex: compras acima de R$ 199 ganham frete gratis).", "Aumenta Ticket Medio"),
        ("Retirada no Local (Local Pickup)", "Cliente retira na loja fisica ou ponto de coleta. Custo zero para ambas as partes.", "Excelente para cidades vizinhas")
    ]
    ry = 630
    for title, desc, tag in rows:
        draw.rectangle([160, ry, 1760, ry + 70], fill=(24, 32, 47))
        draw.text((180, ry + 15), title, font=FONT_BODY_BOLD, fill=C_WHITE)
        draw.text((640, ry + 15), desc, font=FONT_BODY, fill=C_TEXT_MUTED)
        draw.rounded_rectangle([1520, ry + 12, 1740, ry + 52], radius=8, fill=(22, 78, 99))
        draw.text((1535, ry + 18), tag, font=FONT_SMALL_BOLD, fill=ACCENT_AMBER)
        ry += 80

    path = os.path.join(TEMP_DIR, "mod10_slide2.png")
    img.save(path, quality=95)
    return path

def create_slide_3():
    img, draw = create_base_canvas()
    draw_header(draw, "Conceitos Basicos", "Os 3 Pilares do Frete no WooCommerce", "Como o WooCommerce organiza geograficamente e operacionalmente as entregas", PRIMARY_TEAL)
    
    pillars = [
        ("1. Areas de Entrega", "(Shipping Zones)", "Agrupamento geografico onde se aplicam determinadas regras de frete.", [
            "Pode ser um Pais inteiro (Brasil)",
            "Pode ser um Estado (ex: Sao Paulo)",
            "Pode ser delimitado por Faixas de CEP especificas (ex: 01000-000 a 09999-999)",
            "Uma loja pode ter multiplas zonas com precos diferentes"
        ], PRIMARY_TEAL),
        ("2. Metodos de Envio", "(Shipping Methods)", "As opcoes de transporte disponiveis para a zona selecionada.", [
            "Taxa Fixa (valor cobrado fixo por pedido ou item)",
            "Frete Gratis (com ou sem gatilho de valor)",
            "Retirada no Local (balcao da empresa)",
            "Calculo Dinamico via API (Melhor Envio, Frenet, Correios)"
        ], ACCENT_AMBER),
        ("3. Classes de Entrega", "(Shipping Classes)", "Agrupamento de produtos com caracteristicas logisticas especiais.", [
            "Produtos Volumosos ou Frágeis (ex: espelhos, geladeiras)",
            "Produtos Pequenos (ex: acessorios, bijuterias)",
            "Permite aplicar taxas extras ou regras especificas para produtos diferenciados"
        ], (16, 185, 129))
    ]
    
    for i, (title, sub, intro, bullets, col) in enumerate(pillars):
        x = 120 + i * 570
        y = 260
        draw.rounded_rectangle([x, y, x + 530, y + 680], radius=16, fill=C_CARD_BG, outline=col, width=2)
        draw.rounded_rectangle([x, y, x + 530, y + 80], radius=16, fill=col)
        draw.rectangle([x, y + 50, x + 530, y + 80], fill=col)
        draw.text((x + 25, y + 15), title, font=FONT_HEADING, fill=C_WHITE)
        draw.text((x + 25, y + 48), sub, font=FONT_SMALL, fill=(240, 240, 240))
        
        draw.text((x + 30, y + 110), intro, font=FONT_BODY_BOLD, fill=C_WHITE)
        draw.line([x + 30, y + 175, x + 500, y + 175], fill=C_CARD_BORDER, width=1)
        
        by = y + 200
        for b in bullets:
            draw_bullet_dot(draw, x + 40, by + 12, radius=4, color=col)
            words = b.split()
            line = ""
            for w in words:
                if len(line + " " + w) > 34:
                    draw.text((x + 60, by), line, font=FONT_BODY, fill=C_TEXT_LIGHT)
                    by += 28
                    line = w
                else:
                    line = (line + " " + w).strip()
            if line:
                draw.text((x + 60, by), line, font=FONT_BODY, fill=C_TEXT_LIGHT)
                by += 38

    path = os.path.join(TEMP_DIR, "mod10_slide3.png")
    img.save(path, quality=95)
    return path

def create_slide_4():
    img, draw = create_base_canvas()
    draw_header(draw, "Fisica do Frete", "O Calculo de Cubagem e Dimensoes", "Por que volume e peso sao cobrados de forma diferente pelas transportadoras", PRIMARY_TEAL)
    
    # Left: Formula explanation
    draw.rounded_rectangle([120, 260, 920, 980], radius=16, fill=C_CARD_BG, outline=C_CARD_BORDER, width=2)
    draw.text((160, 295), "O que e Peso Cubado?", font=FONT_HEADING, fill=C_WHITE)
    draw.text((160, 340), "Transportadoras cobram pelo espaco que o pacote ocupa no caminhao ou aviao, e nao apenas pelo peso na balanca.", font=FONT_BODY, fill=C_TEXT_MUTED)
    
    # Formula Box
    draw.rounded_rectangle([160, 420, 880, 560], radius=12, fill=(24, 32, 47), outline=ACCENT_AMBER, width=2)
    draw.text((200, 445), "FORMULA OFICIAL DE CUBAGEM (CORREIOS):", font=FONT_SMALL_BOLD, fill=ACCENT_AMBER)
    draw.text((200, 485), "Peso Cubado (kg) = ( C x L x A ) / 6.000", font=FONT_HEADING, fill=C_WHITE)
    draw.text((200, 525), "Onde Comprimento (C), Largura (L) e Altura (A) estao em centimetros.", font=FONT_SMALL, fill=C_TEXT_MUTED)

    rules = [
        ("Peso Real vs Peso Cubado", "A transportadora sempre cobra o MAIOR entre os dois pesos."),
        ("Isencao de Cubagem dos Correios", "Pacotes com peso cubado de ate 5 kg sao tarifados apenas pelo peso real na balanca."),
        ("Dimensao Minima e Maxima", "Correios: Minimo de 15 cm de comprimento, 10 cm de largura e 1 cm de altura. Limite de soma: C+L+A <= 200 cm.")
    ]
    ry = 600
    for r_title, r_desc in rules:
        draw.text((160, ry), r_title, font=FONT_BODY_BOLD, fill=C_WHITE)
        draw.text((160, ry + 28), r_desc, font=FONT_BODY, fill=C_TEXT_MUTED)
        ry += 95

    # Right: Practical Example
    draw.rounded_rectangle([960, 260, 1800, 980], radius=16, fill=C_CARD_BG, outline=PRIMARY_TEAL, width=2)
    draw.text((1000, 295), "Exemplo Pratico: O Erro que Causa Prejuizo", font=FONT_HEADING, fill=C_WHITE)
    
    # Case 1: Pillow / Almofada
    draw.rounded_rectangle([1000, 360, 1760, 620], radius=12, fill=(24, 32, 47), outline=(239, 68, 68), width=1)
    draw.rectangle([1000, 360, 1012, 620], fill=(239, 68, 68))
    draw.text((1035, 385), "Caso 1: Produto Leve mas Volumoso (Travesseiro)", font=FONT_BODY_BOLD, fill=C_WHITE)
    draw.text((1035, 420), "• Peso na balanca (Real): 0,5 kg (500g)", font=FONT_BODY, fill=C_TEXT_LIGHT)
    draw.text((1035, 455), "• Medidas da caixa: 50 cm x 40 cm x 30 cm", font=FONT_BODY, fill=C_TEXT_LIGHT)
    draw.text((1035, 490), "• Calculo Cubado: (50 x 40 x 30) / 6000 = 10,0 kg!", font=FONT_BODY_BOLD, fill=ACCENT_AMBER)
    draw.text((1035, 530), "• SE CADASTRAR APENAS 0,5 KG NO SITE: A loja cobra frete de 500g do cliente (ex: R$ 18),", font=FONT_BODY, fill=(252, 165, 165))
    draw.text((1035, 565), "  mas os Correios cobram 10 kg da loja (ex: R$ 68). Prejuizo direto de R$ 50!", font=FONT_BODY_BOLD, fill=(239, 68, 68))

    # Regra de Ouro
    draw.rounded_rectangle([1000, 660, 1760, 940], radius=12, fill=(22, 78, 99), outline=PRIMARY_TEAL, width=2)
    draw.text((1035, 690), "REGRA DE OURO DO CADASTRO DE PRODUTOS:", font=FONT_BODY_BOLD, fill=ACCENT_AMBER)
    draw.text((1035, 735), "Cadastre SEMPRE as dimensoes e o peso do produto JA DENTRO DA EMBALAGEM de envio.", font=FONT_BODY, fill=C_WHITE)
    draw.text((1035, 775), "Se você vende uma caneca que mede 10cm, mas usa uma caixa de 15x15x15 com plastico bolha,", font=FONT_BODY, fill=C_TEXT_LIGHT)
    draw.text((1035, 815), "o que deve estar no WooCommerce sao os 15x15x15 e o peso total da caixa lacrada!", font=FONT_BODY_BOLD, fill=C_WHITE)

    path = os.path.join(TEMP_DIR, "mod10_slide4.png")
    img.save(path, quality=95)
    return path

def create_slide_5():
    img, draw = create_base_canvas()
    draw_header(draw, "Configuracao Nativa", "Configurando Areas de Entrega no WooCommerce", "Acesse WooCommerce > Configuracoes > Entrega para definir onde sua loja atende", PRIMARY_TEAL)
    
    # Left browser mockup
    draw_browser_chrome(draw, 80, 250, 1080, 750, title="Areas de Entrega - WooCommerce", url="https://minhaloja.com.br/wp-admin/admin.php?page=wc-settings&tab=shipping")
    draw_wp_sidebar(draw, 82, 298, 200, 700, active_item="WooCommerce")
    
    # Content area
    draw.text((310, 320), "Entrega", font=FONT_HEADING, fill=C_WHITE)
    draw.text((310, 360), "Areas de entrega  |  Opcoes de entrega  |  Classes de entrega", font=FONT_SMALL, fill=PRIMARY_TEAL)
    
    # Button Adicionar area
    draw.rounded_rectangle([920, 320, 1120, 360], radius=6, fill=PRIMARY_TEAL)
    draw.text((945, 332), "+ Adicionar Area", font=FONT_SMALL_BOLD, fill=C_WHITE)
    
    # Table of Zones
    draw.rounded_rectangle([310, 400, 1120, 720], radius=8, fill=(34, 43, 60), outline=(56, 68, 88), width=1)
    # Header
    draw.rectangle([310, 400, 1120, 440], fill=(22, 28, 40))
    draw.text((330, 412), "Nome da Zona", font=FONT_SMALL_BOLD, fill=C_TEXT_LIGHT)
    draw.text((580, 412), "Regioes", font=FONT_SMALL_BOLD, fill=C_TEXT_LIGHT)
    draw.text((820, 412), "Metodos de Envio", font=FONT_SMALL_BOLD, fill=C_TEXT_LIGHT)
    
    rows = [
        ("Sao Paulo Capital", "SP (Faixas 01000-000 a 05999-999)", "Taxa Fixa R$ 15, Motoboy, Retirada"),
        ("Sudeste (Interior e Outros)", "SP, RJ, MG, ES", "Melhor Envio (PAC, SEDEX, Jadlog)"),
        ("Brasil Inteiro", "Todo o territorio nacional", "Melhor Envio + Frete Gratis acima de R$ 299"),
        ("Locais nao cobertos", "Restante do mundo", "Nenhum metodo disponivel")
    ]
    zy = 450
    for z_name, z_reg, z_meth in rows:
        draw.line([310, zy, 1120, zy], fill=(48, 58, 76), width=1)
        draw.text((330, zy + 15), z_name, font=FONT_BODY_BOLD, fill=C_WHITE)
        draw.text((580, zy + 15), z_reg, font=FONT_SMALL, fill=C_TEXT_MUTED)
        draw.text((820, zy + 15), z_meth, font=FONT_SMALL, fill=ACCENT_AMBER)
        zy += 65

    # Right explanations
    draw.rounded_rectangle([1200, 250, 1840, 1000], radius=16, fill=C_CARD_BG, outline=C_CARD_BORDER, width=2)
    draw.text((1235, 290), "Como Funciona a Hierarquia:", font=FONT_HEADING, fill=C_WHITE)
    
    steps = [
        ("1. O WooCommerce le de CIMA para BAIXO", "Ele testa a primeira zona que corresponder ao CEP do comprador e para ali. Por isso, zonas mais especificas (ex: Capital) devem ficar no topo."),
        ("2. Delimitacao por CEP", "Ao clicar em 'Limitar a CEPs especificos', você pode colar faixas de CEP como '01000-000...05999-999' para criar frete via motoboy exclusivo."),
        ("3. Flexibilidade Total", "Você pode oferecer Frete Gratis apenas para o seu Estado, enquanto o resto do Brasil paga frete normal."),
        ("4. Zona Padrao de Fallback", "A ultima linha 'Locais nao cobertos' define o que acontece se alguem de fora das zonas cadastradas tentar comprar (ex: 'Nenhum frete disponivel').")
    ]
    sy = 360
    for s_title, s_desc in steps:
        draw.text((1235, sy), s_title, font=FONT_BODY_BOLD, fill=ACCENT_AMBER)
        words = s_desc.split()
        line = ""
        ty = sy + 30
        for w in words:
            if len(line + " " + w) > 42:
                draw.text((1235, ty), line, font=FONT_BODY, fill=C_TEXT_MUTED)
                ty += 24
                line = w
            else:
                line = (line + " " + w).strip()
        if line:
            draw.text((1235, ty), line, font=FONT_BODY, fill=C_TEXT_MUTED)
        sy += 130

    path = os.path.join(TEMP_DIR, "mod10_slide5.png")
    img.save(path, quality=95)
    return path

def create_slide_6():
    img, draw = create_base_canvas()
    draw_header(draw, "Metodos Nativos", "Configurando Taxa Fixa, Frete Gratis e Retirada", "Os tres metodos que ja vem instalados de fabrica no WooCommerce", PRIMARY_TEAL)
    
    methods = [
        ("Taxa Fixa (Flat Rate)", "Cobrar um valor predeterminado", [
            "Valor fixo por pedido (ex: R$ 20,00)",
            "Formula avancada por item: digite '10 + ( 2 * [qty] )' para cobrar R$ 10 base + R$ 2 por produto extra",
            "Formula por porcentagem: digite '15 + ( [cost] * 0.05 )' para adicionar 5% do valor do pedido",
            "Muito usado para entrega expressa local ou transportadora regional"
        ], PRIMARY_TEAL),
        ("Frete Gratis (Free Shipping)", "O maior incentivo de compras", [
            "Condicao 1: Um cupom valido de frete gratis",
            "Condicao 2: Um valor minimo no carrinho (ex: 'Compras acima de R$ 199')",
            "Condicao 3: Valor minimo OU cupom",
            "Condicao 4: Valor minimo E cupom",
            "Dica: Mostre na barra do site 'Faltam R$ 35 para frete gratis!'"
        ], (16, 185, 129)),
        ("Retirada no Local (Local Pickup)", "Retirar no balcao / loja fisica", [
            "Ideal para negocios que possuem ponto fisico ou sede propria",
            "Custo configuravel (geralmente R$ 0,00 gratuito)",
            "Elimina custo de frete e permite entrega imediata",
            "No checkout, exibe instrucoes de endereco para retirada apos o pagamento"
        ], ACCENT_AMBER)
    ]
    
    for i, (m_title, m_sub, bullets, col) in enumerate(methods):
        x = 120 + i * 570
        y = 260
        draw.rounded_rectangle([x, y, x + 530, y + 700], radius=16, fill=C_CARD_BG, outline=col, width=2)
        draw.rounded_rectangle([x, y, x + 530, y + 80], radius=16, fill=col)
        draw.rectangle([x, y + 50, x + 530, y + 80], fill=col)
        draw.text((x + 25, y + 16), m_title, font=FONT_HEADING, fill=C_WHITE)
        draw.text((x + 25, y + 48), m_sub, font=FONT_SMALL, fill=(240, 240, 240))
        
        by = y + 115
        for b in bullets:
            draw_bullet_dot(draw, x + 35, by + 12, radius=4, color=col)
            words = b.split()
            line = ""
            for w in words:
                if len(line + " " + w) > 34:
                    draw.text((x + 55, by), line, font=FONT_BODY, fill=C_TEXT_LIGHT)
                    by += 28
                    line = w
                else:
                    line = (line + " " + w).strip()
            if line:
                draw.text((x + 55, by), line, font=FONT_BODY, fill=C_TEXT_LIGHT)
                by += 44

    path = os.path.join(TEMP_DIR, "mod10_slide6.png")
    img.save(path, quality=95)
    return path

def create_slide_7():
    img, draw = create_base_canvas()
    draw_header(draw, "Agregadores de Frete", "Melhor Envio vs Contrato Direto dos Correios", "Por que quase todo e-commerce iniciante e medio utiliza agregadores logisticos", PRIMARY_TEAL)
    
    # Left: Contrato Direto Correios
    draw.rounded_rectangle([120, 260, 920, 980], radius=16, fill=C_CARD_BG, outline=CORREIOS_BLUE, width=2)
    draw.rounded_rectangle([120, 260, 920, 330], radius=16, fill=CORREIOS_BLUE)
    draw.rectangle([120, 300, 920, 330], fill=CORREIOS_BLUE)
    draw.text((160, 280), "Contrato Proprio com os Correios", font=FONT_HEADING, fill=CORREIOS_YELLOW)
    
    c_direct = [
        ("Exige CNPJ e Analise de Credito", False, "Requer abertura de conta juridica (Correios Facil) e documentacao."),
        ("Cotas Minimas Mensais de Faturamento", False, "Em contratos avancados, ha consumo minimo mensal obrigatorio."),
        ("Apenas 1 Transportadora (Correios)", False, "Se houver greve ou atraso nos Correios, a loja fica sem alternativa."),
        ("Integracao Manual de WebServices", False, "Exige plugin de terceiros com codigo de servico e senha do contrato."),
        ("Coleta no Endereco", True, "Disponivel apenas para volumes diarios elevados ou taxas extras.")
    ]
    dy = 360
    for title, is_good, desc in c_direct:
        if is_good:
            draw_check_icon(draw, 160, dy + 15, radius=12)
        else:
            draw_cross_icon(draw, 160, dy + 15, radius=12)
        draw.text((190, dy), title, font=FONT_BODY_BOLD, fill=C_WHITE)
        draw.text((190, dy + 28), desc, font=FONT_SMALL, fill=C_TEXT_MUTED)
        dy += 95

    # Right: Melhor Envio
    draw.rounded_rectangle([960, 260, 1800, 980], radius=16, fill=C_CARD_BG, outline=ACCENT_AMBER, width=2)
    draw.rounded_rectangle([960, 260, 1800, 330], radius=16, fill=ACCENT_AMBER)
    draw.rectangle([960, 300, 1800, 330], fill=ACCENT_AMBER)
    draw.text((1000, 280), "Agregador: Melhor Envio / Frenet", font=FONT_HEADING, fill=(20, 20, 20))
    
    m_agg = [
        ("Zero Mensalidade e Sem Contrato Fixo", True, "Voce so paga o frete dos pedidos que realmente vender."),
        ("Descontos Coletivos de ate 80%", True, "Por movimentar milhoes de encomendas, repassam tabelas corporativas gigantes."),
        ("Multiplas Transportadoras em 1 so Conta", True, "Correios (SEDEX, PAC), Jadlog, Loggi, LATAM Cargo e Azul Cargo."),
        ("Funciona para CPF e CNPJ", True, "Ideal tanto para MEI/Pessoa Fisica iniciando quanto grandes empresas."),
        ("Geracao e Impressao de Etiquetas em Lote", True, "Gere etiquetas prontas com codigo de barras e poste na agencia ou ponto de coleta.")
    ]
    my = 360
    for title, is_good, desc in m_agg:
        draw_check_icon(draw, 1000, my + 15, radius=12, bg=(16, 185, 129))
        draw.text((1030, my), title, font=FONT_BODY_BOLD, fill=C_WHITE)
        draw.text((1030, my + 28), desc, font=FONT_SMALL, fill=C_TEXT_MUTED)
        my += 95

    path = os.path.join(TEMP_DIR, "mod10_slide7.png")
    img.save(path, quality=95)
    return path

def create_slide_8():
    img, draw = create_base_canvas()
    draw_header(draw, "Instalacao de Plugin", "Instalando o Plugin Oficial do Melhor Envio", "Passo a passo no painel do WordPress para conectar o Melhor Envio ao WooCommerce", PRIMARY_TEAL)
    
    # Left browser
    draw_browser_chrome(draw, 80, 250, 1100, 750, title="Adicionar Plugins - WordPress", url="https://minhaloja.com.br/wp-admin/plugin-install.php?s=melhor+envio")
    draw_wp_sidebar(draw, 82, 298, 200, 700, active_item="Plugins")
    
    # Search Box
    draw.rounded_rectangle([700, 320, 1150, 360], radius=6, fill=(34, 43, 60), outline=(56, 68, 88), width=1)
    draw.text((720, 332), "melhor envio", font=FONT_SMALL, fill=C_WHITE)
    
    # Plugin Card
    draw.rounded_rectangle([310, 400, 1140, 680], radius=12, fill=(30, 38, 52), outline=PRIMARY_TEAL, width=2)
    draw.rounded_rectangle([335, 430, 435, 530], radius=8, fill=PRIMARY_TEAL)
    draw.text((345, 465), "MELHOR", font=FONT_SMALL_BOLD, fill=C_WHITE)
    draw.text((350, 485), "ENVIO", font=FONT_SMALL_BOLD, fill=ACCENT_AMBER)
    
    draw.text((460, 430), "Melhor Envio para WooCommerce", font=FONT_HEADING, fill=C_WHITE)
    draw.text((460, 475), "Por Melhor Envio | Integracao Oficial", font=FONT_SMALL, fill=C_TEXT_MUTED)
    draw.text((460, 505), "Calcule cotacoes com Correios, Jadlog, Loggi e Azul Cargo em tempo real no checkout.", font=FONT_BODY, fill=C_TEXT_LIGHT)
    
    draw_star_badge(draw, 460, 560, count=5)
    draw.text((580, 558), "(Mais de 40.000 instalacoes ativas • Compativel com sua versao)", font=FONT_SMALL, fill=C_TEXT_MUTED)
    
    draw.rounded_rectangle([460, 610, 640, 655], radius=6, fill=(16, 185, 129))
    draw.text((485, 622), "Instalar Agora", font=FONT_BODY_BOLD, fill=C_WHITE)
    draw_arrow_pointer(draw, (700, 720), (550, 655), color=(239, 68, 68))
    draw.text((680, 735), "1. Clique em Instalar e depois Ativar", font=FONT_BODY_BOLD, fill=C_WHITE)

    # Right side instructions
    draw.rounded_rectangle([1220, 250, 1840, 1000], radius=16, fill=C_CARD_BG, outline=C_CARD_BORDER, width=2)
    draw.text((1255, 290), "Passos no WordPress:", font=FONT_HEADING, fill=C_WHITE)
    
    steps = [
        ("1. Acessar o Menu", "No menu lateral esquerdo, va em Plugins > Adicionar Novo Plugin."),
        ("2. Campo de Pesquisa", "No canto superior direito, digite 'Melhor Envio'."),
        ("3. Identificar o Oficial", "Verifique o autor 'Melhor Envio' e o selo de compatibilidade com WooCommerce."),
        ("4. Instalar e Ativar", "Aguarde o download e clique no botao azul 'Ativar'."),
        ("5. Novo Menu Lateral", "Apos ativar, um novo item 'Melhor Envio' aparecera no menu do painel.")
    ]
    iy = 360
    for s_num, s_desc in steps:
        draw.text((1255, iy), s_num, font=FONT_BODY_BOLD, fill=ACCENT_AMBER)
        words = s_desc.split()
        line = ""
        ty = iy + 30
        for w in words:
            if len(line + " " + w) > 42:
                draw.text((1255, ty), line, font=FONT_BODY, fill=C_TEXT_MUTED)
                ty += 24
                line = w
            else:
                line = (line + " " + w).strip()
        if line:
            draw.text((1255, ty), line, font=FONT_BODY, fill=C_TEXT_MUTED)
        iy += 115

    path = os.path.join(TEMP_DIR, "mod10_slide8.png")
    img.save(path, quality=95)
    return path

def create_slide_9():
    img, draw = create_base_canvas()
    draw_header(draw, "Autenticacao & Token", "Conectando o Token da Conta do Melhor Envio", "Como gerar e colar o token de API para comunicar sua loja com o Melhor Envio", PRIMARY_TEAL)
    
    # Left side: Token Generation
    draw.rounded_rectangle([120, 260, 920, 980], radius=16, fill=C_CARD_BG, outline=C_CARD_BORDER, width=2)
    draw.text((160, 295), "1. No Painel do Melhor Envio", font=FONT_HEADING, fill=C_WHITE)
    draw.text((160, 340), "Acesse menv.io e faca login na sua conta gratuita:", font=FONT_BODY, fill=C_TEXT_MUTED)
    
    steps_me = [
        ("Acesse 'Painel de Controle > Integracoes'", "Va na secao de desenvolvedores / aplicacoes."),
        ("Clique em 'Gerar Novo Token'", "Crie uma chave com permissoes de Cotacao de Frete e Geracao de Etiquetas."),
        ("Ambiente de Producao vs Sandbox", "Para testes, pode usar sandbox.melhorenvio.com.br; para vender de verdade, gere no painel oficial."),
        ("Copie o Token JWT", "Um codigo criptografado longo comecando com 'eyJhbGciOi...'. Mantenha-o seguro.")
    ]
    py = 400
    for p_title, p_desc in steps_me:
        draw_bullet_dot(draw, 175, py + 12, radius=4, color=ACCENT_AMBER)
        draw.text((195, py), p_title, font=FONT_BODY_BOLD, fill=C_WHITE)
        draw.text((195, py + 26), p_desc, font=FONT_SMALL, fill=C_TEXT_MUTED)
        py += 90

    # Right side: WooCommerce Configuration
    draw.rounded_rectangle([960, 260, 1800, 980], radius=16, fill=C_CARD_BG, outline=PRIMARY_TEAL, width=2)
    draw.text((1000, 295), "2. No WooCommerce (Melhor Envio > Configuracoes)", font=FONT_HEADING, fill=C_WHITE)
    
    fields = [
        ("Ambiente:", "Producao (Loja Real)"),
        ("Token de Acesso:", "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpZCI6NDU5MDI... (Cole aqui)"),
        ("CEP de Origem (Postagem):", "01310-100 (Endereco fisico de onde as encomendas saem)"),
        ("Documento do Remetente:", "CNPJ ou CPF cadastrado no Melhor Envio"),
        ("Telefone e Email de Contato:", "contato@minhaloja.com.br")
    ]
    fy = 370
    for label, val in fields:
        draw.text((1000, fy), label, font=FONT_BODY_BOLD, fill=C_TEXT_LIGHT)
        draw.rounded_rectangle([1000, fy + 32, 1760, fy + 80], radius=8, fill=(24, 32, 47), outline=(56, 68, 88), width=1)
        draw.text((1020, fy + 45), val, font=FONT_CODE if "eyJ" in val else FONT_BODY, fill=ACCENT_AMBER if "eyJ" in val else C_WHITE)
        fy += 105

    draw.rounded_rectangle([1000, 890, 1760, 955], radius=8, fill=(16, 185, 129))
    draw.text((1030, 910), "Status: CONECTADO COM SUCESSO A CONTA MELHOR ENVIO", font=FONT_BODY_BOLD, fill=C_WHITE)

    path = os.path.join(TEMP_DIR, "mod10_slide9.png")
    img.save(path, quality=95)
    return path

def create_slide_10():
    img, draw = create_base_canvas()
    draw_header(draw, "Selecao de Transportadoras", "Habilitando e Escolhendo os Servicos de Entrega", "Selecione quais servicos ficarao visiveis para os clientes no carrinho de compras", PRIMARY_TEAL)
    
    carriers = [
        ("Correios - PAC", "Economico", "Atende 100% dos municipios brasileiros. Mais barato para encomendas normais.", (16, 185, 129)),
        ("Correios - SEDEX", "Expresso", "Entrega rapida em 24h a 48h para a maioria das capitais e regioes metropolitanas.", PRIMARY_TEAL),
        ("Jadlog - .Package", "Economico Privado", "Excelente tarifa para capitais e interior do Sudeste/Sul. Alternativa perfeita ao PAC.", (147, 51, 234)),
        ("Jadlog - .Com", "Expresso Privado", "Opcao expressa de transportadora privada com rastreamento ponto a ponto.", ACCENT_AMBER),
        ("Loggi", "Urbano / Expresso", "Agilidade maxima em capitais com entrega motorizada.", (59, 130, 246)),
        ("Azul Cargo Express", "Aereo Rapido", "Velocidade maxima para distancias interestaduais longas (Norte/Nordeste para Sul/Sudeste).", (236, 72, 153))
    ]
    
    # 2 columns x 3 rows
    for i, (name, tag, desc, col) in enumerate(carriers):
        col_idx = i % 2
        row_idx = i // 2
        x = 120 + col_idx * 840
        y = 260 + row_idx * 230
        
        draw.rounded_rectangle([x, y, x + 800, y + 200], radius=16, fill=C_CARD_BG, outline=col, width=2)
        draw_check_icon(draw, x + 40, y + 45, radius=16, bg=col)
        draw.text((x + 75, y + 30), name, font=FONT_HEADING, fill=C_WHITE)
        
        # Tag badge
        draw.rounded_rectangle([x + 620, y + 28, x + 760, y + 68], radius=8, fill=(24, 32, 47), outline=col, width=1)
        draw.text((x + 635, y + 36), tag, font=FONT_SMALL_BOLD, fill=col)
        
        draw.text((x + 40, y + 100), desc, font=FONT_BODY, fill=C_TEXT_LIGHT)
        draw.text((x + 40, y + 145), "Status: HABILITADO NO CHECKOUT", font=FONT_SMALL_BOLD, fill=(16, 185, 129))

    path = os.path.join(TEMP_DIR, "mod10_slide10.png")
    img.save(path, quality=95)
    return path

def create_slide_11():
    img, draw = create_base_canvas()
    draw_header(draw, "Ajustes Operacionais", "Prazos Adicionais e Margem de Seguranca de Preco", "Configurando dias extras de preparacao (Lead Time) e margens nos valores calculados", PRIMARY_TEAL)
    
    # Left: Lead Time
    draw.rounded_rectangle([120, 260, 920, 980], radius=16, fill=C_CARD_BG, outline=C_CARD_BORDER, width=2)
    draw.text((160, 295), "Prazo Adicional de Separacao (Lead Time)", font=FONT_HEADING, fill=C_WHITE)
    draw.text((160, 340), "Transportadoras calculam o prazo apenas a partir da postagem no balcao. Mas e o tempo que sua equipe leva para embalar?", font=FONT_BODY, fill=C_TEXT_MUTED)
    
    lead_items = [
        ("Prazo de Separacao e Embalagem", "Recomenda-se adicionar 1 a 2 dias uteis no plugin. Se o SEDEX demora 2 dias, o cliente vera 'Entrega em 3 a 4 dias uteis'."),
        ("Evitar Reclamacoes de Atraso", "E muito melhor entregar ANTES do prazo prometido do que 1 dia depois. Gere surpresa positiva no cliente!"),
        ("Produtos sob Encomenda ou Artesanais", "Se seu produto leva 5 dias para ser fabricado, informe esse prazo adicional para nao frustrar o comprador.")
    ]
    ly = 440
    for l_title, l_desc in lead_items:
        draw_bullet_dot(draw, 175, ly + 12, radius=4, color=PRIMARY_TEAL)
        draw.text((195, ly), l_title, font=FONT_BODY_BOLD, fill=C_WHITE)
        words = l_desc.split()
        line = ""
        ty = ly + 28
        for w in words:
            if len(line + " " + w) > 42:
                draw.text((195, ty), line, font=FONT_BODY, fill=C_TEXT_MUTED)
                ty += 24
                line = w
            else:
                line = (line + " " + w).strip()
        if line:
            draw.text((195, ty), line, font=FONT_BODY, fill=C_TEXT_MUTED)
        ly += 140

    # Right: Cost margins
    draw.rounded_rectangle([960, 260, 1800, 980], radius=16, fill=C_CARD_BG, outline=ACCENT_AMBER, width=2)
    draw.text((1000, 295), "Margem de Custo de Embalagem", font=FONT_HEADING, fill=C_WHITE)
    draw.text((1000, 340), "Caixa de papelao, fita adesiva, plastico bolha e nota fiscal custam dinheiro!", font=FONT_BODY, fill=C_TEXT_MUTED)
    
    margin_options = [
        ("Taxa Fixa Adicional por Envio", "Ex: Inserir + R$ 3,00 em todo frete calculado. Se o PAC der R$ 22, o cliente paga R$ 25 (cobrindo a caixa de papelao).", ACCENT_AMBER),
        ("Taxa Percentual Adicional", "Ex: Inserir + 5% sobre a cotacao para cobrir oscilacoes ou taxas de combustivel das transportadoras.", (16, 185, 129)),
        ("Desconto Promocional de Frete", "Ex: Inserir - 10% para subsidiar parte do frete com a margem do produto e tornar seu frete mais competitivo.", PRIMARY_TEAL)
    ]
    my = 440
    for m_title, m_desc, m_col in margin_options:
        draw.rounded_rectangle([1000, my, 1760, my + 130], radius=12, fill=(24, 32, 47), outline=m_col, width=1)
        draw.text((1025, my + 18), m_title, font=FONT_BODY_BOLD, fill=C_WHITE)
        words = m_desc.split()
        line = ""
        ty = my + 52
        for w in words:
            if len(line + " " + w) > 52:
                draw.text((1025, ty), line, font=FONT_BODY, fill=C_TEXT_MUTED)
                ty += 24
                line = w
            else:
                line = (line + " " + w).strip()
        if line:
            draw.text((1025, ty), line, font=FONT_BODY, fill=C_TEXT_MUTED)
        my += 160

    path = os.path.join(TEMP_DIR, "mod10_slide11.png")
    img.save(path, quality=95)
    return path

def create_slide_12():
    img, draw = create_base_canvas()
    draw_header(draw, "Experiencia do Cliente", "Simulador de Frete na Pagina do Produto", "Permita que o cliente calcule o frete antes de adicionar ao carrinho", PRIMARY_TEAL)
    
    # Left: Mockup Product Page
    draw_browser_chrome(draw, 120, 250, 960, 750, title="Pagina de Produto - Calculo de Frete", url="https://minhaloja.com.br/produto/tenis-running-pro")
    
    # Product area
    draw.rounded_rectangle([160, 330, 460, 630], radius=8, fill=(34, 43, 60))
    draw.text((250, 470), "[FOTO PRODUTO]", font=FONT_BODY_BOLD, fill=C_TEXT_MUTED)
    
    draw.text((500, 330), "Tenis Esportivo Running Pro", font=FONT_HEADING, fill=C_WHITE)
    draw.text((500, 375), "R$ 299,90", font=FONT_HERO, fill=(16, 185, 129))
    draw.text((500, 445), "Em ate 6x de R$ 49,98 sem juros no cartao", font=FONT_SMALL, fill=C_TEXT_MUTED)
    
    # Shipping Box Simulator
    draw.rounded_rectangle([500, 490, 1020, 740], radius=10, fill=(24, 32, 47), outline=PRIMARY_TEAL, width=2)
    draw.text((525, 510), "Consulte o frete e prazo de entrega:", font=FONT_BODY_BOLD, fill=C_WHITE)
    draw.rounded_rectangle([525, 550, 750, 595], radius=6, fill=(34, 43, 60), outline=(56, 68, 88), width=1)
    draw.text((545, 562), "01310-100", font=FONT_BODY, fill=C_WHITE)
    draw.rounded_rectangle([770, 550, 920, 595], radius=6, fill=PRIMARY_TEAL)
    draw.text((795, 562), "Calcular", font=FONT_BODY_BOLD, fill=C_WHITE)
    
    # Results
    draw.text((525, 620), "• SEDEX: R$ 18,90 (Chega em ate 2 dias uteis)", font=FONT_SMALL_BOLD, fill=(16, 185, 129))
    draw.text((525, 655), "• PAC: R$ 12,40 (Chega em ate 5 dias uteis)", font=FONT_SMALL, fill=C_TEXT_LIGHT)
    draw.text((525, 690), "• Jadlog Package: R$ 11,20 (Chega em ate 4 dias uteis)", font=FONT_SMALL, fill=C_TEXT_LIGHT)

    # Right: Benefits
    draw.rounded_rectangle([1120, 250, 1840, 1000], radius=16, fill=C_CARD_BG, outline=C_CARD_BORDER, width=2)
    draw.text((1155, 290), "Por que ter o Simulador na Pagina?", font=FONT_HEADING, fill=C_WHITE)
    
    benefits = [
        ("Reducao Drastica de Carrinhos Abandonados", "O cliente ja descobre o custo de entrega ANTES de iniciar o checkout, evitando a frustracao na etapa final."),
        ("Aumento de Confianca na Loja", "Transparencia de prazos e precos passa credibilidade e profissionalismo."),
        ("Plugins Populares para Adicionar", "• 'WooCommerce Shipping Simulator' (gratis no repositorio do WP)\n• 'Calculador de Frete na Pagina do Produto' de autores brasileiros"),
        ("Memorizacao do CEP", "Ao calcular na pagina do produto, o CEP ja fica salvo na sessao do navegador e e preenchido automaticamente no checkout!")
    ]
    by = 360
    for b_title, b_desc in benefits:
        draw.text((1155, by), b_title, font=FONT_BODY_BOLD, fill=ACCENT_AMBER)
        words = b_desc.split()
        line = ""
        ty = by + 28
        for w in words:
            if len(line + " " + w) > 42:
                draw.text((1155, ty), line, font=FONT_BODY, fill=C_TEXT_MUTED)
                ty += 24
                line = w
            else:
                line = (line + " " + w).strip()
        if line:
            draw.text((1155, ty), line, font=FONT_BODY, fill=C_TEXT_MUTED)
        by += 135

    path = os.path.join(TEMP_DIR, "mod10_slide12.png")
    img.save(path, quality=95)
    return path

def create_slide_13():
    img, draw = create_base_canvas()
    draw_header(draw, "Validacao & Testes", "Testando o Frete no Carrinho e Checkout", "Checklist pratico de testes para garantir que nenhuma cotacao falhe em producao", PRIMARY_TEAL)
    
    # 4 Test scenarios
    tests = [
        ("Cenario 1: CEP da Mesma Cidade (Local)", "01000-000 (SP Capital)", "Verificar se exibe motoboy ou taxa local mais barata.", "Sucesso", (16, 185, 129)),
        ("Cenario 2: CEP de Capital Interestadual", "20040-002 (Rio de Janeiro)", "Validar se PAC, SEDEX e Jadlog sao listados com valores distintos.", "Sucesso", (16, 185, 129)),
        ("Cenario 3: CEP do Interior / Distante", "69005-040 (Manaus / Norte)", "Verificar prazos maiores e opcoes aereas (Azul Cargo / SEDEX).", "Sucesso", (16, 185, 129)),
        ("Cenario 4: CEP Invalido / Digitacao Errada", "00000-000", "Conferir se o sistema avisa amigavelmente 'Nenhum frete disponivel para este CEP'.", "Alerta", ACCENT_AMBER)
    ]
    
    draw.rounded_rectangle([120, 260, 1800, 680], radius=16, fill=C_CARD_BG, outline=C_CARD_BORDER, width=2)
    draw.text((160, 295), "Matriz de Testes Obrigatorios Antes de Abrir a Loja", font=FONT_HEADING, fill=C_WHITE)
    
    ty = 360
    for name, cep, action, status, col in tests:
        draw.rounded_rectangle([160, ty, 1760, ty + 65], radius=8, fill=(24, 32, 47), outline=col, width=1)
        draw.text((185, ty + 18), name, font=FONT_BODY_BOLD, fill=C_WHITE)
        draw.text((640, ty + 18), "CEP: " + cep, font=FONT_CODE, fill=ACCENT_AMBER)
        draw.text((950, ty + 18), action, font=FONT_BODY, fill=C_TEXT_LIGHT)
        draw.rounded_rectangle([1600, ty + 14, 1735, ty + 50], radius=6, fill=col)
        draw.text((1625, ty + 20), status, font=FONT_SMALL_BOLD, fill=C_WHITE)
        ty += 78

    # Troubleshooting box
    draw.rounded_rectangle([120, 720, 1800, 980], radius=16, fill=(22, 78, 99), outline=PRIMARY_TEAL, width=2)
    draw.text((160, 750), "PROBLEMAS COMUNS E COMO RESOLVER:", font=FONT_HEADING, fill=ACCENT_AMBER)
    
    troubles = [
        ("• O frete nao calcula:", "Verifique se o produto cadastrado tem peso e dimensoes maiores que zero."),
        ("• Erro de conexao com API:", "Verifique se o Token no painel do Melhor Envio nao expirou ou foi digitado com espacos extras."),
        ("• Frete muito caro:", "Verifique se você colocou dimensoes em centímetros (ex: 20cm) e nao em metros (0.2m) ou milímetros.")
    ]
    for i, (tb_t, tb_d) in enumerate(troubles):
        draw.text((160, 805 + i * 45), tb_t, font=FONT_BODY_BOLD, fill=C_WHITE)
        draw.text((450, 805 + i * 45), tb_d, font=FONT_BODY, fill=C_TEXT_LIGHT)

    path = os.path.join(TEMP_DIR, "mod10_slide13.png")
    img.save(path, quality=95)
    return path

def create_slide_14():
    img, draw = create_base_canvas()
    draw_header(draw, "Rotina de Expedicao", "Da Venda Aprovada ao Despacho das Encomendas", "Como emitir etiquetas, preparar o pacote e postar nos pontos de envio", PRIMARY_TEAL)
    
    workflow = [
        ("1. Pedido Pago", "Status 'Processando' no WooCommerce confirma pagamento aprovado pelo Mercado Pago/Pix."),
        ("2. Sincronizacao", "O pedido aparece automaticamente no carrinho do Melhor Envio com peso e endereco preenchidos."),
        ("3. Pagamento do Frete", "Pague o valor do frete com o saldo da carteira Melhor Envio (recarga via Pix ou Cartao)."),
        ("4. Impressao da Etiqueta", "Imprima a etiqueta com codigo de barras + Declaracao de Conteudo ou Nota Fiscal (DANFE Simplificada)."),
        ("5. Embalagem e Colagem", "Embale o produto com protecao, cole a etiqueta e coloque a NF/Declaracao em saquinho plastico externo canguru."),
        ("6. Postagem na Agencia", "Entregue o pacote na agencia dos Correios, filial Jadlog ou Ponto Melhor Ponto da sua regiao sem pagar nada no balcao!")
    ]
    
    # 2 columns of 3
    for i, (title, desc) in enumerate(workflow):
        col_idx = i % 2
        row_idx = i // 2
        x = 120 + col_idx * 840
        y = 260 + row_idx * 230
        
        draw.rounded_rectangle([x, y, x + 800, y + 200], radius=16, fill=C_CARD_BG, outline=PRIMARY_TEAL, width=2)
        draw.ellipse([x + 35, y + 35, x + 85, y + 85], fill=PRIMARY_TEAL)
        draw.text((x + 52, y + 42), str(i+1), font=FONT_HEADING, fill=C_WHITE)
        draw.text((x + 105, y + 45), title, font=FONT_HEADING, fill=C_WHITE)
        
        words = desc.split()
        line = ""
        ty = y + 105
        for w in words:
            if len(line + " " + w) > 50:
                draw.text((x + 35, ty), line, font=FONT_BODY, fill=C_TEXT_MUTED)
                ty += 26
                line = w
            else:
                line = (line + " " + w).strip()
        if line:
            draw.text((x + 35, ty), line, font=FONT_BODY, fill=C_TEXT_MUTED)

    path = os.path.join(TEMP_DIR, "mod10_slide14.png")
    img.save(path, quality=95)
    return path

def create_slide_15():
    img, draw = create_base_canvas()
    draw_header(draw, "Checklist Final", "Checklist de Conclusao do Modulo 10", "Tudo pronto para enviar seus produtos para todo o territorio nacional com seguranca", PRIMARY_TEAL)
    
    checklist = [
        ("Areas de Entrega Definidas", "Configuradas no WooCommerce para abranger o Brasil inteiro e regioes prioritarias."),
        ("Peso e Dimensoes Cadastrados", "Produtos com peso real e dimensoes de embalagem preenchidos para calculo de cubagem."),
        ("Plugin Melhor Envio Instalado e Ativo", "Integracao oficial instalada via diretorio de plugins do WordPress."),
        ("Token de API Vinculado", "Chave de seguranca gerada no painel do Melhor Envio e colada nas configuracoes do WooCommerce."),
        ("CEP de Origem (Postagem) Preenchido", "Endereco do seu estoque/loja fisica inserido como ponto de partida das cotacoes."),
        ("Transportadoras Selecionadas", "Habilitadas opcoes economicas (PAC / Jadlog) e expressas (SEDEX)."),
        ("Lead Time e Margem Adicionados", "Prazo extra de preparacao de 1 a 2 dias uteis configurado para evitar atrasos."),
        ("Testes Realizados com Sucesso", "Cotacoes simuladas para capitais e interior sem erros de calculo.")
    ]
    
    # 2 columns x 4 rows
    for i, (title, desc) in enumerate(checklist):
        col_idx = i % 2
        row_idx = i // 4
        x = 120 + col_idx * 840
        y = 260 + row_idx * 165
        
        draw.rounded_rectangle([x, y, x + 800, y + 145], radius=12, fill=C_CARD_BG, outline=(16, 185, 129), width=2)
        draw_check_icon(draw, x + 40, y + 40, radius=16, bg=(16, 185, 129))
        draw.text((x + 75, y + 25), title, font=FONT_HEADING, fill=C_WHITE)
        draw.text((x + 75, y + 75), desc, font=FONT_BODY, fill=C_TEXT_MUTED)

    # Next steps footer
    draw.rounded_rectangle([120, 910, 1800, 980], radius=12, fill=(24, 32, 47), outline=ACCENT_AMBER, width=1)
    draw.text((150, 930), "PROXIMO MODULO (MODULO 11):", font=FONT_BODY_BOLD, fill=ACCENT_AMBER)
    draw.text((450, 930), "Cadastro Completo de Produtos no WooCommerce: Produtos Simples, Variaveis, Atributos e SEO.", font=FONT_BODY, fill=C_WHITE)

    path = os.path.join(TEMP_DIR, "mod10_slide15.png")
    img.save(path, quality=95)
    return path


SLIDE_BUILDERS = [
    (create_slide_1, "Abertura do Modulo 10: Integracao de Frete e Logistica no WooCommerce. Apresentar o papel critico do frete no e-commerce brasileiro e os topicos que serao abordados na aula."),
    (create_slide_2, "Panorama de Logistica: Explicar aos alunos que frete caro e prazo longo sao responsaveis por 66% dos abandonos de carrinho. Destacar a importancia de sempre oferecer uma opcao economica e uma expressa."),
    (create_slide_3, "Conceitos Basicos do WooCommerce: Detalhar a diferenca entre Areas de Entrega (geografia), Metodos de Envio (como sera entregue) e Classes de Entrega (tipos de produto volumoso/pequeno)."),
    (create_slide_4, "Cubagem e Dimensoes: Explicar a formula oficial (C x L x A / 6000). Demonstrar com o exemplo pratico do travesseiro o enorme prejuizo que acontece quando a loja nao cadastra dimensoes."),
    (create_slide_5, "Areas de Entrega Nativas: Navegar pelo painel WordPress > WooCommerce > Configuracoes > Entrega. Explicar como criar zonas por estado e faixas de CEP especificas."),
    (create_slide_6, "Metodos Nativos: Ensinar a configuracao de Taxa Fixa, as 4 regras de ativacao do Frete Gratis (valor minimo no carrinho) e a Retirada no Local sem custos."),
    (create_slide_7, "Comparativo Correios x Agregadores: Mostrar aos alunos porque o Melhor Envio e a melhor opcao para pequenas e medias lojas (descontos de 80%, sem contrato nem mensalidade, varias transportadoras)."),
    (create_slide_8, "Instalacao do Plugin Melhor Envio: Demonstrar a busca na aba Plugins > Adicionar Novo e a ativacao do plugin oficial."),
    (create_slide_9, "Autenticacao e Token: Explicar o que e um Token JWT, como gera-lo no painel do Melhor Envio e como colar no WooCommerce configurando o CEP de postagem."),
    (create_slide_10, "Selecao de Transportadoras: Apresentar as vantagens de cada parceiro (PAC, SEDEX, Jadlog Package, Jadlog .Com, Loggi e Azul Cargo)."),
    (create_slide_11, "Lead Time e Margem: Ensinar os alunos a adicionar 1 a 2 dias uteis no prazo exibido para dar tempo de separar e embalar a encomenda, alem de margem para custear a caixa de papelao."),
    (create_slide_12, "Simulador de Frete na Pagina do Produto: Apresentar o beneficio de permitir o calculo antecipado de frete antes do carrinho para evitar surpresas no checkout."),
    (create_slide_13, "Validacao e Testes: Roteiro pratico de testes simulando CEPs locais, interestaduais e do interior, alem de solucionar problemas comuns de calculo zerado."),
    (create_slide_14, "Rotina de Expedicao: Do pedido pago no WooCommerce a emissao da etiqueta no Melhor Envio, embalagem do pacote e despacho na agencia dos Correios ou Jadlog."),
    (create_slide_15, "Checklist Final: Revisao dos 8 pontos criticos para garantir que a loja esta 100% pronta para calcular fretes e enviar produtos.")
]

def build_presentation_10():
    print("Generating slides for 10 integracao frete.pptx...")
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]
    
    for idx, (func, notes_text) in enumerate(SLIDE_BUILDERS):
        img_path = func()
        slide = prs.slides.add_slide(blank_layout)
        slide.shapes.add_picture(img_path, 0, 0, width=Inches(13.333), height=Inches(7.5))
        
        # Add instructor notes
        notes_slide = slide.notes_slide
        notes_slide.notes_text_frame.text = f"[AULA E-COMMERCE - MODULO 10 - SLIDE {idx+1}]\n{notes_text}"
        print(f"  Slide {idx+1}/{len(SLIDE_BUILDERS)} generated.")

    out_name = "10 integracao frete.pptx"
    prs.save(out_name)
    print(f"Presentation saved successfully as '{out_name}' ({os.path.getsize(out_name)} bytes).")

if __name__ == "__main__":
    build_presentation_10()

"""
Generator for: 11 cadastro produtos.pptx
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

PRIMARY_INDIGO = (79, 70, 229)     # Indigo 600
ACCENT_GREEN = (16, 185, 129)      # Emerald 500
ACCENT_AMBER = (245, 158, 11)      # Amber 500

def create_slide_1():
    img, draw = create_base_canvas()
    draw.rounded_rectangle([120, 160, 520, 215], radius=12, fill=PRIMARY_INDIGO)
    draw.text((140, 172), "MODULO 11 • CATALOGO", font=FONT_HEADING, fill=C_WHITE)
    
    draw.text((120, 250), "Cadastro de Produtos", font=FONT_HERO, fill=C_WHITE)
    draw.text((120, 335), "& Gestao de Catalogo no WooCommerce", font=FONT_HERO, fill=ACCENT_GREEN)
    
    draw.text((120, 440), "Produtos Simples, Produtos Variaveis com Atributos, Controle de Estoque, SEO e Fotos de Alta Conversao", font=FONT_SUBTITLE, fill=C_TEXT_MUTED)
    
    features = [
        ("Produtos Simples & Variaveis", "Aprenda a cadastrar desde um item unico ate produtos com variacoes de cor, tamanho, voltagem e precos distintos."),
        ("Gestao Precisa de Estoque", "Controle de SKU, quantidade disponivel, notificacoes de estoque baixo e bloqueio automatico de itens esgotados."),
        ("Copywriting & SEO para Vendas", "Tecnicas de descricao curta, descricao detalhada, imagens otimizadas e meta tags para atrair compradores do Google.")
    ]
    for i, (title, desc) in enumerate(features):
        x = 120 + i * 570
        y = 540
        draw.rounded_rectangle([x, y, x + 530, y + 260], radius=16, fill=C_CARD_BG, outline=C_CARD_BORDER, width=2)
        draw.ellipse([x + 30, y + 30, x + 80, y + 80], fill=PRIMARY_INDIGO)
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

    draw.rounded_rectangle([120, 850, 1800, 960], radius=12, fill=(49, 46, 129))
    draw.text((160, 880), "OBJETIVO DA AULA:", font=FONT_BODY_BOLD, fill=ACCENT_GREEN)
    draw.text((360, 880), "Capacitar o aluno a estruturar um catalogo comercial profissional, com fotos impecaveis, atributos e estoque controlado.", font=FONT_BODY, fill=C_WHITE)

    path = os.path.join(TEMP_DIR, "mod11_slide1.png")
    img.save(path, quality=95)
    return path

def create_slide_2():
    img, draw = create_base_canvas()
    draw_header(draw, "Estrutura Visual", "Anatomia de uma Pagina de Produto de Alta Conversao", "Os elementos visuais e informativos essenciais que transformam visitantes em compradores", PRIMARY_INDIGO)
    
    # Left: Mockup Product Page
    draw_browser_chrome(draw, 120, 250, 1050, 750, title="Pagina de Produto Modelo - WooCommerce", url="https://minhaloja.com.br/produto/mochila-executiva-impermeavel")
    
    # Photo box
    draw.rounded_rectangle([160, 320, 520, 680], radius=12, fill=(34, 43, 60), outline=(56, 68, 88), width=1)
    draw.text((250, 480), "[FOTO PRINCIPAL]", font=FONT_BODY_BOLD, fill=C_TEXT_MUTED)
    draw.text((260, 510), "1000 x 1000 px", font=FONT_SMALL, fill=(100, 116, 139))
    
    # Thumbnails
    for i in range(4):
        draw.rounded_rectangle([160 + i * 92, 700, 240 + i * 92, 780], radius=6, fill=(24, 32, 47), outline=(56, 68, 88), width=1)
        draw.text((185 + i * 92, 730), f"F{i+1}", font=FONT_SMALL_BOLD, fill=C_TEXT_MUTED)
        
    # Product info
    draw.text((560, 320), "Mochila Executiva Impermeavel Slim", font=FONT_HEADING, fill=C_WHITE)
    draw.text((560, 365), "Ref/SKU: MOC-SLIM-01  |  Categoria: Acessorios", font=FONT_SMALL, fill=C_TEXT_MUTED)
    
    draw_star_badge(draw, 560, 400, count=5)
    draw.text((680, 398), "4.9 (84 avaliacoes de clientes)", font=FONT_SMALL, fill=ACCENT_AMBER)
    
    draw.text((560, 440), "De R$ 289,90 por", font=FONT_SMALL, fill=(239, 68, 68))
    draw.text((560, 465), "R$ 199,90", font=FONT_HERO, fill=ACCENT_GREEN)
    draw.text((560, 535), "em ate 6x de R$ 33,31 sem juros ou R$ 189,90 no Pix (5% OFF)", font=FONT_SMALL_BOLD, fill=C_TEXT_LIGHT)
    
    # Short description
    draw.rounded_rectangle([560, 570, 1130, 660], radius=8, fill=(24, 32, 47))
    draw.text((580, 585), "Compartimento acolchoado para notebook ate 15.6', tecido resistente", font=FONT_SMALL, fill=C_TEXT_LIGHT)
    draw.text((580, 615), "a agua, entrada USB externa e design ergonomico antifurto.", font=FONT_SMALL, fill=C_TEXT_LIGHT)
    
    # Buy Button
    draw.rounded_rectangle([560, 680, 840, 740], radius=8, fill=ACCENT_GREEN)
    draw.text((615, 695), "COMPRAR AGORA", font=FONT_HEADING, fill=C_WHITE)
    
    # Trust badges
    draw.text((560, 765), "Envio Imediato • Garantia de 1 Ano • Devolucao Gratis em 7 dias", font=FONT_SMALL_BOLD, fill=ACCENT_AMBER)

    # Right: Principles
    draw.rounded_rectangle([1210, 250, 1840, 1000], radius=16, fill=C_CARD_BG, outline=C_CARD_BORDER, width=2)
    draw.text((1245, 290), "6 Elementos de Conversao:", font=FONT_HEADING, fill=C_WHITE)
    
    principles = [
        ("1. Titulo Especifico", "Contenha marca, modelo e principal diferencial (evite titulos genericos como apenas 'Mochila')."),
        ("2. Multiplas Fotos em Fundo Limpo", "Fotos de frente, lado, detalhes internos e foto de uso no dia a dia com modelo."),
        ("3. Preco com Ancoragem", "Mostrar o preco original 'De' riscado ao lado do 'Por' ativa o gatilho da oportunidade."),
        ("4. Parcelamento & Desconto Pix", "Destaque o valor da parcela sem juros e desconto para pagamento a vista."),
        ("5. Botao CTA em Destaque", "Cor chamativa (verde ou laranja) com texto focado em acao direta."),
        ("6. Selos de Confianca e Garantia", "Reduz a ansiedade de compra eliminando medo de golpe ou produto com defeito.")
    ]
    py = 350
    for p_t, p_d in principles:
        draw.text((1245, py), p_t, font=FONT_BODY_BOLD, fill=ACCENT_GREEN)
        words = p_d.split()
        line = ""
        ty = py + 26
        for w in words:
            if len(line + " " + w) > 42:
                draw.text((1245, ty), line, font=FONT_BODY, fill=C_TEXT_MUTED)
                ty += 22
                line = w
            else:
                line = (line + " " + w).strip()
        if line:
            draw.text((1245, ty), line, font=FONT_BODY, fill=C_TEXT_MUTED)
        py += 105

    path = os.path.join(TEMP_DIR, "mod11_slide2.png")
    img.save(path, quality=95)
    return path

def create_slide_3():
    img, draw = create_base_canvas()
    draw_header(draw, "Tipos de Produtos", "Os 4 Tipos de Produtos Nativos no WooCommerce", "Escolha a estrutura correta para o tipo de mercadoria que você comercializa", PRIMARY_INDIGO)
    
    types = [
        ("Produto Simples", "O tipo mais comum de e-commerce", [
            "Item unico, sem opcoes de escolha para o comprador",
            "Exemplos: Livro fisico, caneca unica, cabo USB de 1 metro",
            "Possui apenas 1 preco, 1 SKU e 1 controle de estoque",
            "Pode ser fisico, virtual ou baixavel"
        ], PRIMARY_INDIGO),
        ("Produto Variavel", "Produtos com opcoes configuraveis", [
            "Permite ao cliente escolher tamanho, cor, voltagem, sabor, etc.",
            "Exemplos: Camiseta (P, M, G), Tenis (38 ao 44), Ventilador (110v/220v)",
            "Cada combinacao tem seu proprio SKU, preco e estoque individual",
            "O mais utilizado no varejo de moda e tecnologia"
        ], ACCENT_GREEN),
        ("Produto Agrupado", "Colecoes de itens relacionados", [
            "Conjunto de produtos simples vendidos em uma unica tela",
            "Exemplo: Kit Escritorio (Mesa + Cadeira + Luminaria)",
            "O cliente escolhe quais itens do kit deseja incluir no carrinho",
            "Estimula a compra casada (Cross-Selling)"
        ], ACCENT_AMBER),
        ("Produto Externo / Afiliado", "Redirecionamento para terceiros", [
            "O produto e exibido na loja, mas o botao leva para outro site",
            "Exemplo: Vender como afiliado Amazon, Hotmart, Shopee",
            "Voce nao processa o pagamento nem envia a mercadoria",
            "Recebe comissao pela indicacao da venda"
        ], (236, 72, 153))
    ]
    
    # 4 cards
    for i, (p_title, p_sub, bullets, col) in enumerate(types):
        x = 80 + i * 445
        y = 260
        draw.rounded_rectangle([x, y, x + 420, y + 700], radius=16, fill=C_CARD_BG, outline=col, width=2)
        draw.rounded_rectangle([x, y, x + 420, y + 80], radius=16, fill=col)
        draw.rectangle([x, y + 50, x + 420, y + 80], fill=col)
        draw.text((x + 20, y + 16), p_title, font=FONT_HEADING, fill=C_WHITE)
        draw.text((x + 20, y + 48), p_sub, font=FONT_SMALL, fill=(240, 240, 240))
        
        by = y + 115
        for b in bullets:
            draw_bullet_dot(draw, x + 25, by + 12, radius=4, color=col)
            words = b.split()
            line = ""
            for w in words:
                if len(line + " " + w) > 28:
                    draw.text((x + 45, by), line, font=FONT_BODY, fill=C_TEXT_LIGHT)
                    by += 28
                    line = w
                else:
                    line = (line + " " + w).strip()
            if line:
                draw.text((x + 45, by), line, font=FONT_BODY, fill=C_TEXT_LIGHT)
                by += 44

    path = os.path.join(TEMP_DIR, "mod11_slide3.png")
    img.save(path, quality=95)
    return path

def create_slide_4():
    img, draw = create_base_canvas()
    draw_header(draw, "Painel de Cadastro", "Acessando a Tela de Adicionar Novo Produto", "Visao geral da interface de criacao de produtos no painel administrativo do WordPress", PRIMARY_INDIGO)
    
    # Browser
    draw_browser_chrome(draw, 80, 250, 1150, 750, title="Adicionar Novo Produto - WordPress", url="https://minhaloja.com.br/wp-admin/post-new.php?post_type=product")
    draw_wp_sidebar(draw, 82, 298, 200, 700, active_item="Produtos")
    
    # Post Title input
    draw.text((310, 320), "Adicionar novo produto", font=FONT_HEADING, fill=C_WHITE)
    draw.rounded_rectangle([310, 365, 1190, 420], radius=6, fill=(34, 43, 60), outline=(56, 68, 88), width=1)
    draw.text((330, 380), "Digite o nome do produto aqui (Ex: Fone de Ouvido Bluetooth Noise Cancelling)", font=FONT_BODY, fill=C_TEXT_MUTED)
    
    # Description editor
    draw.rounded_rectangle([310, 445, 1190, 680], radius=8, fill=(24, 32, 47), outline=(56, 68, 88), width=1)
    draw.rectangle([310, 445, 1190, 485], fill=(18, 24, 36))
    draw.text((330, 455), "B  I  U  |  Lista  |  Link  |  Adicionar Midia", font=FONT_SMALL_BOLD, fill=C_TEXT_LIGHT)
    draw.text((330, 510), "Descricao Completa e Detalhada do Produto...", font=FONT_BODY, fill=C_TEXT_LIGHT)
    draw.text((330, 545), "• Especificacoes tecnicas, materiais utilizados, modo de uso e cuidados.", font=FONT_BODY, fill=C_TEXT_MUTED)
    draw.text((330, 580), "• Tabela de medidas (caso vestuario) e duvidas frequentes.", font=FONT_BODY, fill=C_TEXT_MUTED)
    
    # Right sidebar boxes
    draw.rounded_rectangle([310, 710, 1190, 970], radius=8, fill=(34, 43, 60), outline=PRIMARY_INDIGO, width=2)
    draw.text((330, 730), "Dados do Produto — [ Produto Simples v ]", font=FONT_HEADING, fill=C_WHITE)
    draw.text((330, 775), "Geral  |  Inventario  |  Entrega  |  Produtos Relacionados  |  Atributos  |  Avancado", font=FONT_SMALL_BOLD, fill=ACCENT_GREEN)

    # Right explanations
    draw.rounded_rectangle([1260, 250, 1840, 1000], radius=16, fill=C_CARD_BG, outline=C_CARD_BORDER, width=2)
    draw.text((1295, 290), "Zonas da Tela de Edicao:", font=FONT_HEADING, fill=C_WHITE)
    
    zones = [
        ("1. Titulo Principal", "O nome oficial do produto. Sera usado na URL amigavel (slug) e na tag H1 da pagina."),
        ("2. Descricao Longa (Editor Superior)", "Fica nas abas inferiores da pagina. Perfeita para textos completos, tabelas de medidas, videos e fotos de detalhes."),
        ("3. Caixa 'Dados do Produto'", "O coracao do cadastro. Onde se define preco, estoque, dimensoes para frete e variacoes."),
        ("4. Descricao Curta (Abaixo)", "O texto de resumo que aparece logo ao lado da foto e abaixo do preco. Foco em beneficios diretos."),
        ("5. Coluna Lateral Direita", "Onde se define a Imagem Principal, Galeria de Fotos Adicionais, Categorias e Tags.")
    ]
    zy = 360
    for z_t, z_d in zones:
        draw.text((1295, zy), z_t, font=FONT_BODY_BOLD, fill=ACCENT_GREEN)
        words = z_d.split()
        line = ""
        ty = zy + 28
        for w in words:
            if len(line + " " + w) > 38:
                draw.text((1295, ty), line, font=FONT_BODY, fill=C_TEXT_MUTED)
                ty += 24
                line = w
            else:
                line = (line + " " + w).strip()
        if line:
            draw.text((1295, ty), line, font=FONT_BODY, fill=C_TEXT_MUTED)
        zy += 120

    path = os.path.join(TEMP_DIR, "mod11_slide4.png")
    img.save(path, quality=95)
    return path

def create_slide_5():
    img, draw = create_base_canvas()
    draw_header(draw, "Descricoes & Copywriting", "Descricao Curta vs Descricao Longa: O Que Colocar?", "Aprenda a distribuir as informacoes para prender a atencao do cliente e sanar todas as duvidas", PRIMARY_INDIGO)
    
    # Left: Short Desc
    draw.rounded_rectangle([120, 260, 920, 980], radius=16, fill=C_CARD_BG, outline=ACCENT_GREEN, width=2)
    draw.rounded_rectangle([120, 260, 920, 330], radius=16, fill=ACCENT_GREEN)
    draw.rectangle([120, 300, 920, 330], fill=ACCENT_GREEN)
    draw.text((160, 280), "Descricao Breve (Curta)", font=FONT_HEADING, fill=C_WHITE)
    
    draw.text((160, 355), "ONDE FICA: Logo ao lado da imagem e acima do botao de comprar.", font=FONT_SMALL_BOLD, fill=ACCENT_AMBER)
    draw.text((160, 390), "OBJETIVO: Convencer em menos de 5 segundos de leitura.", font=FONT_SMALL_BOLD, fill=C_WHITE)
    
    short_items = [
        ("Bullet Points Diretos", "Use de 3 a 5 tópicos destacando os principais diferenciais."),
        ("Beneficio Principal", "Responda a pergunta: 'Qual dor esse produto resolve?'"),
        ("Destaque de Material", "Ex: '100% Algodao Egipcio 400 fios' ou 'Aco Inoxidavel Cirurgico'."),
        ("Dimensao Resumida", "Capacidade ou tamanho principal (ex: 'Capacidade 1.5 Litros')."),
        ("Frase de Urgencia/Garantia", "Ex: 'Pronta entrega com despacho em 24h uteis'.")
    ]
    sy = 450
    for s_t, s_d in short_items:
        draw_check_icon(draw, 160, sy + 15, radius=12, bg=ACCENT_GREEN)
        draw.text((190, sy), s_t, font=FONT_BODY_BOLD, fill=C_WHITE)
        draw.text((190, sy + 28), s_d, font=FONT_BODY, fill=C_TEXT_MUTED)
        sy += 95

    # Right: Long Desc
    draw.rounded_rectangle([960, 260, 1800, 980], radius=16, fill=C_CARD_BG, outline=PRIMARY_INDIGO, width=2)
    draw.rounded_rectangle([960, 260, 1800, 330], radius=16, fill=PRIMARY_INDIGO)
    draw.rectangle([960, 300, 1800, 330], fill=PRIMARY_INDIGO)
    draw.text((1000, 280), "Descricao Completa (Longa)", font=FONT_HEADING, fill=C_WHITE)
    
    draw.text((1000, 355), "ONDE FICA: Nas abas inferiores da pagina ('Descricao' e 'Informacoes Adicionais').", font=FONT_SMALL_BOLD, fill=ACCENT_AMBER)
    draw.text((1000, 390), "OBJETIVO: Tirar todas as duvidas tecnicas e ranquear no Google (SEO).", font=FONT_SMALL_BOLD, fill=C_WHITE)
    
    long_items = [
        ("Storytelling da Marca/Produto", "Explique a origem, o conceito e os padroes de qualidade da fabricacao."),
        ("Ficha Tecnica Completa", "Peso, dimensoes exatas, tensao eletrica, composicao textil, validade."),
        ("Tabela de Medidas Ilustrada", "Fundamental em moda e calcados para reduzir taxa de troca e devolucao."),
        ("Guia de Cuidados e Lavagem", "Instrucoes de manutencao que aumentam a durabilidade do item."),
        ("Perguntas Frequentes (FAQ)", "Responda as duvidas mais comuns recebidas pelo suporte/WhatsApp.")
    ]
    ly = 450
    for l_t, l_d in long_items:
        draw_check_icon(draw, 1000, ly + 15, radius=12, bg=PRIMARY_INDIGO)
        draw.text((1030, ly), l_t, font=FONT_BODY_BOLD, fill=C_WHITE)
        draw.text((1030, ly + 28), l_d, font=FONT_BODY, fill=C_TEXT_MUTED)
        ly += 95

    path = os.path.join(TEMP_DIR, "mod11_slide5.png")
    img.save(path, quality=95)
    return path

def create_slide_6():
    img, draw = create_base_canvas()
    draw_header(draw, "Dados do Produto: Geral", "Configurando Preco Normal, Promocional e Agendamento", "Aba 'Geral' da caixa de dados do produto no WooCommerce", PRIMARY_INDIGO)
    
    # Left: Form mockup
    draw.rounded_rectangle([120, 260, 1050, 980], radius=16, fill=C_CARD_BG, outline=PRIMARY_INDIGO, width=2)
    draw.text((160, 295), "Dados do Produto — Aba Geral", font=FONT_HEADING, fill=C_WHITE)
    
    fields = [
        ("Preco Normal (R$):", "199,90", "O preco regular de venda do produto."),
        ("Preco Promocional (R$):", "149,90", "Preco com desconto. Ao preencher, o preco normal ganha risco tachado."),
        ("Agendar Promocao:", "De: 25/11/2026  Ate: 30/11/2026", "Ativa o preco promocional automaticamente no periodo da Black Friday."),
        ("Status do Imposto:", "Tributavel", "Se o item esta sujeito a cobranca de impostos."),
        ("Classe de Imposto:", "Padrao", "Regra fiscal configurada na loja.")
    ]
    
    fy = 360
    for label, val, desc in fields:
        draw.text((160, fy), label, font=FONT_BODY_BOLD, fill=C_TEXT_LIGHT)
        draw.rounded_rectangle([160, fy + 32, 1000, fy + 80], radius=8, fill=(24, 32, 47), outline=(56, 68, 88), width=1)
        draw.text((180, fy + 45), val, font=FONT_BODY, fill=ACCENT_GREEN if "149" in val else C_WHITE)
        draw.text((160, fy + 90), desc, font=FONT_SMALL, fill=C_TEXT_MUTED)
        fy += 120

    # Right: Marketing Tips
    draw.rounded_rectangle([1120, 260, 1840, 1000], radius=16, fill=C_CARD_BG, outline=C_CARD_BORDER, width=2)
    draw.text((1155, 290), "Estrategias de Precificacao:", font=FONT_HEADING, fill=C_WHITE)
    
    tips = [
        ("Efeito Ancoragem de Preco", "Sempre que possivel, cadastre um Preco Normal mais alto e um Preco Promocional ativo. O cerebro humano percebe ganho imediato com o selo 'OFERTA' ou '25% OFF'."),
        ("Agendamento Automatico de Campanhas", "Use o agendador para datas comemorativas (Dia das Maes, Natal, Black Friday). A promocao entra no ar a meia-noite e encerra sozinha sem precisar de intervencao manual."),
        ("Precos Terminados em ,90 ou ,99", "A precificacao psicologica comprovadamente converte mais do que numeros redondos (R$ 149,90 vende mais do que R$ 150,00)."),
        ("Politica de Margem Minima", "Certifique-se de que o preco promocional cobre o custo da mercadoria + impostos + taxas de pagamento (3% a 5%) + taxa de embalagem.")
    ]
    ty = 360
    for t_title, t_desc in tips:
        draw.text((1155, ty), t_title, font=FONT_BODY_BOLD, fill=ACCENT_AMBER)
        words = t_desc.split()
        line = ""
        ly = ty + 28
        for w in words:
            if len(line + " " + w) > 42:
                draw.text((1155, ly), line, font=FONT_BODY, fill=C_TEXT_MUTED)
                ly += 24
                line = w
            else:
                line = (line + " " + w).strip()
        if line:
            draw.text((1155, ly), line, font=FONT_BODY, fill=C_TEXT_MUTED)
        ty += 140

    path = os.path.join(TEMP_DIR, "mod11_slide6.png")
    img.save(path, quality=95)
    return path

def create_slide_7():
    img, draw = create_base_canvas()
    draw_header(draw, "Dados do Produto: Inventario", "Controle de Estoque, SKU e Notificacoes de Falta", "Aba 'Inventario' para evitar vender itens esgotados e controlar seu almoxarifado", PRIMARY_INDIGO)
    
    # Left: Form mockup
    draw.rounded_rectangle([120, 260, 1050, 980], radius=16, fill=C_CARD_BG, outline=PRIMARY_INDIGO, width=2)
    draw.text((160, 295), "Dados do Produto — Aba Inventario", font=FONT_HEADING, fill=C_WHITE)
    
    inv_fields = [
        ("SKU (Identificador Unico):", "TENIS-RUN-AZUL-41", "Stock Keeping Unit. Codigo unico para localizar a mercadoria no estoque."),
        ("Gerenciar Estoque?", "[ X ] Ativar gerenciamento de estoque em nivel de produto", "Quando marcado, o WooCommerce passa a dar baixa automatica nas vendas."),
        ("Quantidade em Estoque:", "25 unidades", "O saldo fisico real presente nas prateleiras."),
        ("Permitir Encomendas?", "Nao permitir  (ou Permitir com aviso ao cliente)", "Se permite que o cliente compre mesmo quando o estoque zerar."),
        ("Limiar de Estoque Baixo:", "3 unidades", "Quando o estoque chegar a 3, a loja envia um email de alerta ao administrador.")
    ]
    iy = 360
    for label, val, desc in inv_fields:
        draw.text((160, iy), label, font=FONT_BODY_BOLD, fill=C_TEXT_LIGHT)
        draw.rounded_rectangle([160, iy + 30, 1000, iy + 75], radius=8, fill=(24, 32, 47), outline=(56, 68, 88), width=1)
        draw.text((180, iy + 42), val, font=FONT_CODE if "TENIS" in val else FONT_BODY, fill=ACCENT_AMBER if "TENIS" in val else C_WHITE)
        draw.text((160, iy + 85), desc, font=FONT_SMALL, fill=C_TEXT_MUTED)
        iy += 120

    # Right: SKU Good Practices
    draw.rounded_rectangle([1120, 260, 1840, 1000], radius=16, fill=C_CARD_BG, outline=C_CARD_BORDER, width=2)
    draw.text((1155, 290), "Boas Praticas de Gestao de Estoque:", font=FONT_HEADING, fill=C_WHITE)
    
    practices = [
        ("Como Criar um SKU Inteligente", "Estruture com: [CATEGORIA]-[MODELO]-[COR]-[TAMANHO].\nExemplo: CAM-POLO-AZL-M ou FONE-BT-PRT-UNI. Facilita muito a expedicao e empacotamento."),
        ("Evite Furos de Estoque a Todo Custo", "Vender um produto que nao existe no estoque gera reclamacao no Reclame Aqui, cancelamento de compra e perda irrevogavel de confianca."),
        ("Limite de 1 Item por Pedido", "O campo 'Vendido Individualmente' impede que o comprador adicione mais de 1 unidade no carrinho (otimo para pecas exclusivas ou cursos)."),
        ("Sincronizacao Multicanal", "O mesmo SKU cadastrado aqui sera usado pelo plugin do Mercado Livre para dar baixa quando uma venda ocorrer por la!")
    ]
    py = 360
    for p_title, p_desc in practices:
        draw.text((1155, py), p_title, font=FONT_BODY_BOLD, fill=ACCENT_GREEN)
        words = p_desc.split()
        line = ""
        ly = py + 28
        for w in words:
            if len(line + " " + w) > 42:
                draw.text((1155, ly), line, font=FONT_BODY, fill=C_TEXT_MUTED)
                ly += 24
                line = w
            else:
                line = (line + " " + w).strip()
        if line:
            draw.text((1155, ly), line, font=FONT_BODY, fill=C_TEXT_MUTED)
        py += 140

    path = os.path.join(TEMP_DIR, "mod11_slide7.png")
    img.save(path, quality=95)
    return path

def create_slide_8():
    img, draw = create_base_canvas()
    draw_header(draw, "Dados do Produto: Entrega", "Configurando Peso, Dimensoes e Classe de Frete", "Aba 'Entrega' para integrar com o Melhor Envio e calcular cotacoes exatas", PRIMARY_INDIGO)
    
    # Left: Form mockup
    draw.rounded_rectangle([120, 260, 1050, 980], radius=16, fill=C_CARD_BG, outline=PRIMARY_INDIGO, width=2)
    draw.text((160, 295), "Dados do Produto — Aba Entrega", font=FONT_HEADING, fill=C_WHITE)
    
    deliv_fields = [
        ("Peso (kg):", "0.650", "Peso total do produto COM a embalagem de despacho (em quilogramas)."),
        ("Dimensoes (cm):", "Comprimento: 25 cm   Largura: 18 cm   Altura: 12 cm", "As medidas externas da caixa de papelao ou envelope de postagem."),
        ("Classe de Entrega:", "Nenhuma (Padrao)  [ou 'Itens Frágeis', 'Volumosos']", "Agrupa produtos que exigem fretes ou taxas diferenciadas.")
    ]
    dy = 370
    for label, val, desc in deliv_fields:
        draw.text((160, dy), label, font=FONT_BODY_BOLD, fill=C_TEXT_LIGHT)
        draw.rounded_rectangle([160, dy + 32, 1000, dy + 80], radius=8, fill=(24, 32, 47), outline=(56, 68, 88), width=1)
        draw.text((180, dy + 45), val, font=FONT_BODY, fill=ACCENT_AMBER)
        draw.text((160, dy + 92), desc, font=FONT_SMALL, fill=C_TEXT_MUTED)
        dy += 130

    # Warning Box
    draw.rounded_rectangle([160, 770, 1000, 940], radius=12, fill=(22, 78, 99), outline=ACCENT_AMBER, width=2)
    draw.text((190, 795), "LEMBRE-SE DA REGRA DO MODULO 10:", font=FONT_BODY_BOLD, fill=ACCENT_AMBER)
    draw.text((190, 830), "Nunca cadastre as medidas do produto 'nu' (fora da caixa).", font=FONT_BODY, fill=C_WHITE)
    draw.text((190, 865), "Se você vende sapatos, as medidas devem ser da caixa do sapato + envelope,", font=FONT_BODY, fill=C_TEXT_LIGHT)
    draw.text((190, 895), "pois e o volume externo que sera cubado pelos Correios e Jadlog!", font=FONT_BODY_BOLD, fill=C_WHITE)

    # Right: Virtual & Downloadable Products
    draw.rounded_rectangle([1120, 260, 1840, 1000], radius=16, fill=C_CARD_BG, outline=C_CARD_BORDER, width=2)
    draw.text((1155, 290), "E para Produtos Digitais ou Servicos?", font=FONT_HEADING, fill=C_WHITE)
    
    digitals = [
        ("Opcao 'Virtual' (Sem Frete)", "Ao marcar a caixa 'Virtual' no topo dos dados do produto, a aba de Entrega DESAPARECE completamente e o WooCommerce nao cobra frete no checkout (ideal para consultorias e ingressos)."),
        ("Opcao 'Baixavel' (Download)", "Permite fazer upload de arquivos (PDF, ZIP, MP3, Video). Apos a aprovacao do pagamento, o cliente recebe o link seguro para baixar o arquivo na area do cliente."),
        ("Limite de Downloads e Expiracao", "Voce pode limitar quantos downloads o cliente pode fazer (ex: maximo 3 vezes) e definir prazo de validade do link (ex: expira em 30 dias).")
    ]
    py = 370
    for p_title, p_desc in digitals:
        draw.text((1155, py), p_title, font=FONT_BODY_BOLD, fill=ACCENT_GREEN)
        words = p_desc.split()
        line = ""
        ly = py + 28
        for w in words:
            if len(line + " " + w) > 42:
                draw.text((1155, ly), line, font=FONT_BODY, fill=C_TEXT_MUTED)
                ly += 24
                line = w
            else:
                line = (line + " " + w).strip()
        if line:
            draw.text((1155, ly), line, font=FONT_BODY, fill=C_TEXT_MUTED)
        py += 170

    path = os.path.join(TEMP_DIR, "mod11_slide8.png")
    img.save(path, quality=95)
    return path

def create_slide_9():
    img, draw = create_base_canvas()
    draw_header(draw, "Produtos Variaveis", "Criando Atributos Globais de Variacao", "Acesse Produtos > Atributos para cadastrar Tamanhos, Cores e Voltagens", PRIMARY_INDIGO)
    
    # Left: Attributes creation
    draw_browser_chrome(draw, 80, 250, 1100, 750, title="Atributos de Produtos - WooCommerce", url="https://minhaloja.com.br/wp-admin/edit.php?post_type=product&page=product_attributes")
    draw_wp_sidebar(draw, 82, 298, 200, 700, active_item="Produtos")
    
    # Form
    draw.text((310, 320), "Atributos", font=FONT_HEADING, fill=C_WHITE)
    draw.text((310, 360), "Nome do Atributo: (Ex: Tamanho, Cor, Voltagem)", font=FONT_SMALL, fill=C_TEXT_LIGHT)
    draw.rounded_rectangle([310, 390, 680, 435], radius=6, fill=(34, 43, 60), outline=(56, 68, 88), width=1)
    draw.text((330, 402), "Cor", font=FONT_BODY_BOLD, fill=C_WHITE)
    
    draw.rounded_rectangle([310, 460, 480, 505], radius=6, fill=PRIMARY_INDIGO)
    draw.text((335, 472), "Adicionar Atributo", font=FONT_SMALL_BOLD, fill=C_WHITE)
    
    # Table of existing attributes
    draw.rounded_rectangle([720, 360, 1150, 720], radius=8, fill=(24, 32, 47), outline=(56, 68, 88), width=1)
    draw.rectangle([720, 360, 1150, 400], fill=(18, 24, 36))
    draw.text((740, 372), "Nome", font=FONT_SMALL_BOLD, fill=C_TEXT_LIGHT)
    draw.text((880, 372), "Termos / Valores", font=FONT_SMALL_BOLD, fill=C_TEXT_LIGHT)
    
    attrs = [
        ("Tamanho", "P, M, G, GG, XG"),
        ("Cor", "Preto, Branco, Azul Marinho, Vermelho"),
        ("Voltagem", "110V, 220V, Bivolt"),
        ("Sabor", "Chocolate, Baunilha, Morango")
    ]
    ay = 410
    for a_name, a_vals in attrs:
        draw.line([720, ay, 1150, ay], fill=(48, 58, 76), width=1)
        draw.text((740, ay + 15), a_name, font=FONT_BODY_BOLD, fill=C_WHITE)
        draw.text((880, ay + 15), a_vals, font=FONT_SMALL, fill=ACCENT_GREEN)
        ay += 65

    # Right: Attributes concept
    draw.rounded_rectangle([1220, 250, 1840, 1000], radius=16, fill=C_CARD_BG, outline=C_CARD_BORDER, width=2)
    draw.text((1255, 290), "Atributos Globais vs Locais:", font=FONT_HEADING, fill=C_WHITE)
    
    concepts = [
        ("Atributos Globais (Recomendado)", "Criados no menu Produtos > Atributos. Ficam salvos para sempre e podem ser reutilizados em centenas de produtos diferentes."),
        ("Facilidade de Manutencao", "Se no futuro você passar a vender o tamanho 'XGG', basta adicionar o termo uma vez no Atributo Global e ele estara disponivel para toda a loja."),
        ("Filtros na Loja (Widgets)", "Apenas atributos globais permitem criar filtros na lateral da loja (ex: 'Filtrar por Cor: Azul' ou 'Filtrar por Tamanho: M')."),
        ("Atributos Locais (Especificos)", "Criados diretamente dentro de um unico produto. Uteis apenas para caracteristicas raras que nao se repetem em outros itens.")
    ]
    cy = 360
    for c_t, c_d in concepts:
        draw.text((1255, cy), c_t, font=FONT_BODY_BOLD, fill=ACCENT_AMBER)
        words = c_d.split()
        line = ""
        ty = cy + 28
        for w in words:
            if len(line + " " + w) > 40:
                draw.text((1255, ty), line, font=FONT_BODY, fill=C_TEXT_MUTED)
                ty += 24
                line = w
            else:
                line = (line + " " + w).strip()
        if line:
            draw.text((1255, ty), line, font=FONT_BODY, fill=C_TEXT_MUTED)
        cy += 135

    path = os.path.join(TEMP_DIR, "mod11_slide9.png")
    img.save(path, quality=95)
    return path

def create_slide_10():
    img, draw = create_base_canvas()
    draw_header(draw, "Geracao de Variacoes", "Criando as Variacoes Individuais do Produto", "Como transformar atributos em combinacoes compráveis com precos e estoques proprios", PRIMARY_INDIGO)
    
    # Left browser / variation matrix
    draw.rounded_rectangle([120, 260, 1050, 980], radius=16, fill=C_CARD_BG, outline=PRIMARY_INDIGO, width=2)
    draw.text((160, 295), "Aba 'Variacoes' do Produto Variavel", font=FONT_HEADING, fill=C_WHITE)
    draw.text((160, 335), "Apos selecionar Atributos 'Usado para variacoes', clique em 'Gerar Variacoes':", font=FONT_SMALL, fill=C_TEXT_MUTED)
    
    # Variation rows
    vars_list = [
        ("Cor: Preto | Tamanho: P", "SKU: CAM-PRT-P", "R$ 79,90", "Estoque: 12 un"),
        ("Cor: Preto | Tamanho: M", "SKU: CAM-PRT-M", "R$ 79,90", "Estoque: 20 un"),
        ("Cor: Preto | Tamanho: G", "SKU: CAM-PRT-G", "R$ 89,90", "Estoque: 8 un"),
        ("Cor: Branco | Tamanho: P", "SKU: CAM-BRC-P", "R$ 79,90", "Estoque: 5 un"),
        ("Cor: Branco | Tamanho: M", "SKU: CAM-BRC-M", "R$ 79,90", "Estoque: 15 un")
    ]
    vy = 390
    for v_name, v_sku, v_price, v_stock in vars_list:
        draw.rounded_rectangle([160, vy, 1000, vy + 95], radius=8, fill=(24, 32, 47), outline=(56, 68, 88), width=1)
        draw.rounded_rectangle([180, vy + 15, 245, vy + 80], radius=6, fill=PRIMARY_INDIGO)
        draw.text((195, vy + 38), "[FOTO]", font=FONT_SMALL_BOLD, fill=C_WHITE)
        draw.text((265, vy + 18), v_name, font=FONT_BODY_BOLD, fill=C_WHITE)
        draw.text((265, vy + 52), v_sku, font=FONT_CODE, fill=C_TEXT_MUTED)
        draw.text((680, vy + 32), v_price, font=FONT_HEADING, fill=ACCENT_GREEN)
        draw.text((840, vy + 35), v_stock, font=FONT_SMALL_BOLD, fill=ACCENT_AMBER)
        vy += 110

    # Right: Rules
    draw.rounded_rectangle([1120, 260, 1840, 1000], radius=16, fill=C_CARD_BG, outline=C_CARD_BORDER, width=2)
    draw.text((1155, 290), "Regras Criticas de Variacoes:", font=FONT_HEADING, fill=C_WHITE)
    
    rules = [
        ("Preco Individual Obrigatorio", "SE UMA VARIACAO NAO TIVER PRECO CADASTRADO, ELA NAO APARECERA NA LOJA. Este e o erro mais comum de alunos iniciantes."),
        ("Fotos Especificas por Variacao", "Quando o cliente clica na bolinha 'Cor: Vermelho', a foto principal da tela troca instantaneamente para o produto vermelho!"),
        ("Precos Diferenciados por Tamanho", "Tamanhos especiais (ex: GG, XG) ou materiais premium podem custar mais caro do que os tamanhos convencionais."),
        ("Estoque Isolado", "Se acabar o tamanho 'M', apenas o botao 'M' ficara indisponivel; os tamanhos P e G continuarao sendo vendidos normalmente.")
    ]
    ry = 360
    for r_t, r_d in rules:
        draw.text((1155, ry), r_t, font=FONT_BODY_BOLD, fill=ACCENT_GREEN)
        words = r_d.split()
        line = ""
        ly = ry + 28
        for w in words:
            if len(line + " " + w) > 42:
                draw.text((1155, ly), line, font=FONT_BODY, fill=C_TEXT_MUTED)
                ly += 24
                line = w
            else:
                line = (line + " " + w).strip()
        if line:
            draw.text((1155, ly), line, font=FONT_BODY, fill=C_TEXT_MUTED)
        ry += 135

    path = os.path.join(TEMP_DIR, "mod11_slide10.png")
    img.save(path, quality=95)
    return path

def create_slide_11():
    img, draw = create_base_canvas()
    draw_header(draw, "Fotografia & Imagens", "Padrao de Imagens para E-Commerce de Alta Performance", "A foto e o unico contato visual que o cliente tem com o produto antes de comprar", PRIMARY_INDIGO)
    
    specs = [
        ("Dimensao Padrao Quadrada", "1000 x 1000 pixels (Proporcao 1:1)", "Garante que o zoom do WooCommerce funcione com nitidez cristalina sem distorcer o layout."),
        ("Formato Otimizado WebP / JPEG", "Peso de arquivo < 200 KB por foto", "Imagens pesadas deixam a loja lenta no celular e espantam compradores. Use conversores WebP."),
        ("Fundo Neutro / Branco", "Destaque absoluto para a peca", "Facilita a visualizacao de detalhes, texturas e cores reais sem distracoes de fundo."),
        ("Galeria com Minimo de 4 Fotos", "Frente, Verso, Detalhes e Uso Real", "Mostre a etiqueta, costura, embalagem aberta e modelo utilizando o produto no dia a dia.")
    ]
    
    # 4 spec cards
    for i, (title, sub, desc) in enumerate(specs):
        x = 120 + i * 420
        y = 260
        draw.rounded_rectangle([x, y, x + 395, y + 360], radius=16, fill=C_CARD_BG, outline=PRIMARY_INDIGO, width=2)
        draw.rounded_rectangle([x, y, x + 395, y + 80], radius=16, fill=PRIMARY_INDIGO)
        draw.rectangle([x, y + 50, x + 395, y + 80], fill=PRIMARY_INDIGO)
        draw.text((x + 20, y + 16), title, font=FONT_HEADING, fill=C_WHITE)
        draw.text((x + 20, y + 48), sub, font=FONT_SMALL, fill=(240, 240, 240))
        
        words = desc.split()
        line = ""
        ty = y + 110
        for w in words:
            if len(line + " " + w) > 26:
                draw.text((x + 20, ty), line, font=FONT_BODY, fill=C_TEXT_LIGHT)
                ty += 28
                line = w
            else:
                line = (line + " " + w).strip()
        if line:
            draw.text((x + 20, ty), line, font=FONT_BODY, fill=C_TEXT_LIGHT)

    # Bottom tips box
    draw.rounded_rectangle([120, 660, 1800, 980], radius=16, fill=(24, 32, 47), outline=ACCENT_AMBER, width=2)
    draw.text((160, 690), "NOMENCLATURA DE ARQUIVO E TEXTO ALTERNATIVO (ALT TEXT):", font=FONT_HEADING, fill=ACCENT_AMBER)
    
    img_rules = [
        ("• Nome do Arquivo no Computador:", "Nunca suba fotos com nomes genericos como 'IMG_4829.jpg'. Renomeie antes para 'tenis-running-pro-preto-lateral.webp'."),
        ("• Texto Alternativo (Alt Text):", "Descreva a imagem detalhadamente no painel de midia (ex: 'Tenis preto esportivo com sola branca visto de lado')."),
        ("• Importancia para Acessibilidade:", "Leitores de tela para deficientes visuais leem o Alt Text para explicar o produto."),
        ("• Ranqueamento no Google Imagens:", "Milhares de pessoas buscam produtos pelo Google Imagens e entram na sua loja pelo Alt Text!")
    ]
    iy = 745
    for ir_t, ir_d in img_rules:
        draw.text((160, iy), ir_t, font=FONT_BODY_BOLD, fill=C_WHITE)
        draw.text((560, iy), ir_d, font=FONT_BODY, fill=C_TEXT_LIGHT)
        iy += 52

    path = os.path.join(TEMP_DIR, "mod11_slide11.png")
    img.save(path, quality=95)
    return path

def create_slide_12():
    img, draw = create_base_canvas()
    draw_header(draw, "Organizacao de Catalogo", "Categorias, Subcategorias e Tags no WooCommerce", "Como organizar a arvore mercadologica para uma navegacao fluida e intuitiva", PRIMARY_INDIGO)
    
    # Left: Category tree mockup
    draw.rounded_rectangle([120, 260, 920, 980], radius=16, fill=C_CARD_BG, outline=PRIMARY_INDIGO, width=2)
    draw.text((160, 295), "Arvore de Categorias (Exemplo Moda)", font=FONT_HEADING, fill=C_WHITE)
    
    tree = [
        ("Roupas (Categoria Pai)", 0, (16, 185, 129)),
        ("├── Camisetas (Subcategoria)", 1, C_TEXT_LIGHT),
        ("│   ├── Camisetas Basicas", 2, C_TEXT_MUTED),
        ("│   └── Camisetas Estampadas", 2, C_TEXT_MUTED),
        ("├── Calcas (Subcategoria)", 1, C_TEXT_LIGHT),
        ("│   ├── Jeans", 2, C_TEXT_MUTED),
        ("│   └── Moletom", 2, C_TEXT_MUTED),
        ("Acessorios (Categoria Pai)", 0, (16, 185, 129)),
        ("├── Mochilas & Bolsas", 1, C_TEXT_LIGHT),
        ("└── Bones & Chapeus", 1, C_TEXT_LIGHT)
    ]
    try_y = 360
    for name, level, col in tree:
        draw.text((160 + level * 20, try_y), name, font=FONT_BODY_BOLD if level == 0 else FONT_BODY, fill=col)
        try_y += 45

    # Right: Categorias vs Tags
    draw.rounded_rectangle([960, 260, 1800, 980], radius=16, fill=C_CARD_BG, outline=ACCENT_AMBER, width=2)
    draw.text((1000, 295), "Diferenca entre Categoria e Tag:", font=FONT_HEADING, fill=C_WHITE)
    
    cats_vs_tags = [
        ("Categorias: Hierarquicas e Amplas", "Estruturam o menu principal da loja. Um produto deve pertencer a poucas categorias (idealmente 1 ou 2). Ex: 'Tenis Esportivo'."),
        ("Tags (Etiquetas): Nao-Hierarquicas e Transversais", "Termos descritivos que conectam produtos de categorias diferentes por uma caracteristica comum. Ex: '#impermeavel', '#algodao-organico', '#black-friday', '#presente-namorados'."),
        ("Evite Excesso de Categorias", "Lojas com 50 categorias no menu geram confusao mental. Mantenha no maximo de 5 a 7 categorias principais no topo."),
        ("Migalhas de Pao (Breadcrumbs)", "A hierarquia correta gera o caminho de navegacao no topo do site: 'Home > Roupas > Camisetas > Camiseta Polo'." )
    ]
    cy = 360
    for c_title, c_desc in cats_vs_tags:
        draw.text((1000, cy), c_title, font=FONT_BODY_BOLD, fill=ACCENT_AMBER)
        words = c_desc.split()
        line = ""
        ty = cy + 28
        for w in words:
            if len(line + " " + w) > 48:
                draw.text((1000, ty), line, font=FONT_BODY, fill=C_TEXT_MUTED)
                ty += 24
                line = w
            else:
                line = (line + " " + w).strip()
        if line:
            draw.text((1000, ty), line, font=FONT_BODY, fill=C_TEXT_MUTED)
        cy += 140

    path = os.path.join(TEMP_DIR, "mod11_slide12.png")
    img.save(path, quality=95)
    return path

def create_slide_13():
    img, draw = create_base_canvas()
    draw_header(draw, "SEO para E-Commerce", "Otimizacao para Motores de Busca no Cadastro de Produtos", "Como atrair clientes do Google sem gastar nada com anuncios pagos", PRIMARY_INDIGO)
    
    # Left Google SERP snippet
    draw.rounded_rectangle([120, 260, 1050, 980], radius=16, fill=C_CARD_BG, outline=PRIMARY_INDIGO, width=2)
    draw.text((160, 295), "Previa de Exibicao no Google (SERP)", font=FONT_HEADING, fill=C_WHITE)
    
    # Google snippet box
    draw.rounded_rectangle([160, 350, 1000, 520], radius=8, fill=(24, 32, 47), outline=(56, 68, 88), width=1)
    draw.text((185, 375), "https://minhaloja.com.br › produto › mochila-executiva-slim", font=FONT_SMALL, fill=(148, 163, 184))
    draw.text((185, 405), "Mochila Executiva Impermeavel Slim - Frete Gratis | Minha Loja", font=FONT_HEADING, fill=(138, 180, 248))
    draw.text((185, 455), "Compre Mochila Executiva Slim com tecido impermeavel e entrada USB.", font=FONT_BODY, fill=(189, 193, 198))
    draw.text((185, 482), "Em ate 6x sem juros no cartao ou 5% OFF no Pix. Envio imediato para todo o Brasil!", font=FONT_BODY, fill=(189, 193, 198))

    seo_items = [
        ("URL Amigavel (Slug):", "curta, limpa e com palavra-chave principal (evite datas e codigos)."),
        ("Titulo SEO (<title>):", "limite de 60 caracteres. Coloque palavra-chave + beneficio + nome da loja."),
        ("Meta Description:", "limite de 155 caracteres. Texto persuasivo com gatilhos de clique (frete, parcelamento).")
    ]
    sy = 560
    for s_t, s_d in seo_items:
        draw.text((160, sy), s_t, font=FONT_BODY_BOLD, fill=ACCENT_GREEN)
        draw.text((160, sy + 28), s_d, font=FONT_BODY, fill=C_TEXT_LIGHT)
        sy += 75

    # Right: Plugins SEO
    draw.rounded_rectangle([1120, 260, 1840, 1000], radius=16, fill=C_CARD_BG, outline=C_CARD_BORDER, width=2)
    draw.text((1155, 290), "Plugins Recomendados de SEO:", font=FONT_HEADING, fill=C_WHITE)
    
    seo_plugins = [
        ("Rank Math SEO ou Yoast SEO", "Permitem preencher os campos de Meta Title e Meta Description diretamente na tela de edicao de cada produto do WooCommerce."),
        ("Schema Markup de Produto (Rich Snippets)", "Adiciona automaticamente o preco, estoque e avaliacoes com estrelinhas douradas diretamente no resultado do Google!"),
        ("Evite Conteudo Duplicado", "Nunca copie e cole a descricao oficial do fabricante que ja esta presente em dezenas de outros sites. Crie um texto unico e autentico."),
        ("Sitemap XML Automatico", "Informa instantaneamente aos robos do Google toda vez que um novo produto for publicado na loja.")
    ]
    py = 360
    for p_title, p_desc in seo_plugins:
        draw.text((1155, py), p_title, font=FONT_BODY_BOLD, fill=ACCENT_AMBER)
        words = p_desc.split()
        line = ""
        ly = py + 28
        for w in words:
            if len(line + " " + w) > 42:
                draw.text((1155, ly), line, font=FONT_BODY, fill=C_TEXT_MUTED)
                ly += 24
                line = w
            else:
                line = (line + " " + w).strip()
        if line:
            draw.text((1155, ly), line, font=FONT_BODY, fill=C_TEXT_MUTED)
        py += 140

    path = os.path.join(TEMP_DIR, "mod11_slide13.png")
    img.save(path, quality=95)
    return path

def create_slide_14():
    img, draw = create_base_canvas()
    draw_header(draw, "Publicacao & Status", "Opcoes de Publicacao, Visibilidade e Destaques", "Caixa 'Publicar' e controles de catalogo no canto superior direito", PRIMARY_INDIGO)
    
    # Left Publish box
    draw.rounded_rectangle([120, 260, 920, 980], radius=16, fill=C_CARD_BG, outline=PRIMARY_INDIGO, width=2)
    draw.text((160, 295), "Caixa 'Publicar' no WordPress", font=FONT_HEADING, fill=C_WHITE)
    
    pub_states = [
        ("Status: Publicado / Rascunho", "Permite salvar o produto como 'Rascunho' para continuar editando depois sem exibi-lo ao publico."),
        ("Visibilidade: Publico / Privado / Protegido", "Publico fica aberto a todos; Privado visivel apenas a administradores; Protegido exige senha para abrir a pagina."),
        ("Publicar Imediatamente ou Agendado", "Voce pode programar o lancamento de um produto para uma data e hora especifica no futuro."),
        ("Visibilidade do Catalogo", "Escolha se o item aparece na Loja e Busca, apenas na Busca, ou oculto (acessivel apenas por link direto)."),
        ("Produto em Destaque (Featured Star)", "Ao clicar na estrelinha de destaque, o item ganha prioridade na vitrine da Home e sessoes promocionais.")
    ]
    py = 360
    for p_t, p_d in pub_states:
        draw_bullet_dot(draw, 175, py + 12, radius=4, color=ACCENT_GREEN)
        draw.text((195, py), p_t, font=FONT_BODY_BOLD, fill=C_WHITE)
        words = p_d.split()
        line = ""
        ty = py + 28
        for w in words:
            if len(line + " " + w) > 42:
                draw.text((195, ty), line, font=FONT_BODY, fill=C_TEXT_MUTED)
                ty += 22
                line = w
            else:
                line = (line + " " + w).strip()
        if line:
            draw.text((195, ty), line, font=FONT_BODY, fill=C_TEXT_MUTED)
        py += 115

    # Right: Pre-launch checklist
    draw.rounded_rectangle([960, 260, 1800, 980], radius=16, fill=C_CARD_BG, outline=ACCENT_AMBER, width=2)
    draw.text((1000, 295), "Antes de Apertar o Botao 'Publicar':", font=FONT_HEADING, fill=C_WHITE)
    
    checks = [
        ("Clique no botao 'Pre-visualizar'", "Veja como a pagina ficara exatamente na tela de um computador e no smartphone."),
        ("Confira os Precos Digitados", "Evite o erro classico de digitar R$ 19,90 ao inves de R$ 199,00."),
        ("Verifique os Vinculos de Variacoes", "Se for produto variavel, certifique-se de que TODAS as variacoes possuem preco cadastrado."),
        ("Teste a Calculadora de Frete", "Digite um CEP na pagina de teste para conferir se o peso e as dimensoes estao retornando valores corretos."),
        ("Confira a Categoria Selecionada", "Evite que o produto caia na categoria padrao 'Sem Categoria' (Uncategorized).")
    ]
    cy = 360
    for c_t, c_d in checks:
        draw_check_icon(draw, 1000, cy + 15, radius=12, bg=ACCENT_AMBER)
        draw.text((1030, cy), c_t, font=FONT_BODY_BOLD, fill=C_WHITE)
        words = c_d.split()
        line = ""
        ty = cy + 28
        for w in words:
            if len(line + " " + w) > 44:
                draw.text((1030, ty), line, font=FONT_BODY, fill=C_TEXT_MUTED)
                ty += 22
                line = w
            else:
                line = (line + " " + w).strip()
        if line:
            draw.text((1030, ty), line, font=FONT_BODY, fill=C_TEXT_MUTED)
        cy += 115

    path = os.path.join(TEMP_DIR, "mod11_slide14.png")
    img.save(path, quality=95)
    return path

def create_slide_15():
    img, draw = create_base_canvas()
    draw_header(draw, "Checklist de Qualidade", "Checklist de Conclusao do Modulo 11", "Garantia de que seu catalogo de produtos segue o padrao ouro do comercio eletronico", PRIMARY_INDIGO)
    
    checklist = [
        ("Titulo Claro e Comercial", "Com nome de modelo, marca e diferencial sem termos genericos."),
        ("Descricao Curta em Bullets", "Resumo de 3 a 5 pontos fortes focado em beneficios diretos ao comprador."),
        ("Descricao Longa & Ficha Tecnica", "Especificacoes tecnicas completas, cuidados e medidas para SEO."),
        ("Preco & Promocao Agendada", "Preco com ancoragem psicologica e margem financeira garantida."),
        ("Controle de Estoque & SKU", "Codigo unico cadastrado e baixa automatica habilitada."),
        ("Peso e Dimensoes Preenchidos", "Medidas externas da caixa cadastradas para calculo de frete."),
        ("Variacoes Configuradas Corretamente", "Fotos, precos e estoques individuais para cada tamanho ou cor."),
        ("Fotos em Alta Resolucao (WebP)", "Imagens quadradas 1000x1000px com Alt Text e fundo neutro.")
    ]
    
    # 2 columns x 4 rows
    for i, (title, desc) in enumerate(checklist):
        col_idx = i % 2
        row_idx = i // 4
        x = 120 + col_idx * 840
        y = 260 + row_idx * 165
        
        draw.rounded_rectangle([x, y, x + 800, y + 145], radius=12, fill=C_CARD_BG, outline=ACCENT_GREEN, width=2)
        draw_check_icon(draw, x + 40, y + 40, radius=16, bg=ACCENT_GREEN)
        draw.text((x + 75, y + 25), title, font=FONT_HEADING, fill=C_WHITE)
        draw.text((x + 75, y + 75), desc, font=FONT_BODY, fill=C_TEXT_MUTED)

    # Next steps footer
    draw.rounded_rectangle([120, 910, 1800, 980], radius=12, fill=(49, 46, 129), outline=ACCENT_AMBER, width=1)
    draw.text((150, 930), "PROXIMO MODULO (MODULO 12 - FINAL):", font=FONT_BODY_BOLD, fill=ACCENT_AMBER)
    draw.text((500, 930), "Simulacao Completa de Compra: Da Vitrine ao Checkout, Pagamento Pix/Cartao e Gestao de Pedidos.", font=FONT_BODY, fill=C_WHITE)

    path = os.path.join(TEMP_DIR, "mod11_slide15.png")
    img.save(path, quality=95)
    return path


SLIDE_BUILDERS = [
    (create_slide_1, "Abertura do Modulo 11: Cadastro e Gestao de Produtos no WooCommerce. Apresentar os objetivos da aula: aprender a cadastrar produtos simples e variaveis, estoque, dimensoes e otimizacao para vendas."),
    (create_slide_2, "Anatomia da Pagina de Produto: Explicar aos alunos os 6 elementos essenciais que aumentam a taxa de conversao (titulo, fotos, ancoragem de preco, parcelamento, CTA e selos de confianca)."),
    (create_slide_3, "Os 4 Tipos de Produtos: Detalhar Produto Simples (item unico), Produto Variavel (opcoes de tamanho/cor), Produto Agrupado (kits) e Produto Externo/Afiliado."),
    (create_slide_4, "Interface de Cadastro: Navegar pelo painel WordPress > Produtos > Adicionar Novo e apresentar a estrutura de campos."),
    (create_slide_5, "Descricao Curta x Descricao Longa: Ensinar onde cada texto aparece na loja e quais informacoes colocar em cada um para maximizar conversoes e SEO."),
    (create_slide_6, "Dados do Produto: Geral: Ensinar a configurar Preco Normal, Preco Promocional e o agendamento automatico de promocoes para datas comemorativas."),
    (create_slide_7, "Dados do Produto: Inventario: Ensinar o conceito de SKU, controle de estoque em nivel de produto e alertas de estoque baixo."),
    (create_slide_8, "Dados do Produto: Entrega: Relembrar a regra de ouro das dimensoes de embalagem e explicar produtos virtuais e baixaveis (digitais)."),
    (create_slide_9, "Atributos Globais: Demonstrar a criacao de atributos de Tamanho, Cor e Voltagem em Produtos > Atributos."),
    (create_slide_10, "Geracao de Variacoes: Mostrar como vincular atributos e alertar sobre o erro classico de esquecer de colocar preco nas variacoes."),
    (create_slide_11, "Fotografia e Imagens: Explicar padroes profissionais de fotos (1000x1000px, WebP, fundo neutro, fotos de angulos variados e Alt Text)."),
    (create_slide_12, "Categorias e Tags: Ensinar a montar a arvore mercadologica da loja para navegacao facil e breadcrumbs."),
    (create_slide_13, "SEO para Produtos: Ensinar a configurar URL amigavel, Meta Title e Meta Description para atrair clientes do Google organicamente."),
    (create_slide_14, "Publicacao e Status: Ensinar o uso de Rascunho, Pre-visualizacao, Produto em Destaque e verificacoes antes de colocar no ar."),
    (create_slide_15, "Checklist Final: Revisao dos 8 pontos de qualidade para garantir que cada produto cadastrado seja uma maquina de vendas.")
]

def build_presentation_11():
    print("Generating slides for 11 cadastro produtos.pptx...")
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
        notes_slide.notes_text_frame.text = f"[AULA E-COMMERCE - MODULO 11 - SLIDE {idx+1}]\n{notes_text}"
        print(f"  Slide {idx+1}/{len(SLIDE_BUILDERS)} generated.")

    out_name = "11 cadastro produtos.pptx"
    prs.save(out_name)
    print(f"Presentation saved successfully as '{out_name}' ({os.path.getsize(out_name)} bytes).")

if __name__ == "__main__":
    build_presentation_11()
